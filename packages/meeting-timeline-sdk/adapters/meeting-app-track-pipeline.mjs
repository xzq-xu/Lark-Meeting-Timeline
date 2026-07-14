import { compactObject, normalizeAbsoluteMs } from '../index.mjs';
import { normalizeMeetingAppSnapshots } from './meeting-apps.mjs';
import { buildMeetingPlatformParticipantTrack } from './platform-participant-track.mjs';
import { normalizeMeetingPlatform } from './platform-setup.mjs';
import { buildMeetingPlatformSpeakerTrack } from './platform-speaker-track.mjs';

export const MEETING_APP_TRACK_PIPELINE_SCHEMA = 'meeting_app_track_pipeline';
export const MEETING_APP_TRACK_PIPELINE_SCHEMA_VERSION = 1;

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

function snapshotArray(input = {}) {
  if (Array.isArray(input)) return input;
  return asArray(firstNonEmpty(
    input.snapshots,
    input.samples,
    input.captures,
    input.rows,
    input.items,
    input.observations,
    input,
  ));
}

function observedAtMs(snapshot = {}, fallback) {
  const value = firstNonEmpty(
    snapshot.observedAtMs,
    snapshot.observed_at_ms,
    snapshot.timestamp_ms,
    snapshot.timestampMs,
    snapshot.detected_at_ms,
    snapshot.detectedAtMs,
    fallback,
  );
  return value == null ? undefined : normalizeAbsoluteMs(value, 'meeting_app_track_snapshot_time');
}

function meetingFromSnapshot(snapshot = {}, options = {}) {
  const meeting = snapshot.meeting ?? {};
  const platform = firstNonEmpty(
    options.platform,
    snapshot.platform,
    snapshot.provider,
    meeting.platform,
  );
  const meetingId = firstNonEmpty(
    options.meeting_id,
    options.meetingId,
    snapshot.meeting_id,
    snapshot.meetingId,
    snapshot.external_meeting_id,
    snapshot.externalMeetingId,
    meeting.meeting_id,
    meeting.meetingId,
    meeting.external_meeting_id,
    meeting.externalMeetingId,
  );
  if (!platform || !meetingId) return null;
  return compactObject({
    platform: normalizeMeetingPlatform(platform),
    meeting_id: String(meetingId),
    external_meeting_id: firstNonEmpty(
      snapshot.external_meeting_id,
      snapshot.externalMeetingId,
      meeting.external_meeting_id,
      meeting.externalMeetingId,
    ),
    meeting_url: firstNonEmpty(snapshot.meeting_url, snapshot.meetingUrl, snapshot.url, meeting.meeting_url, meeting.url),
    title: firstNonEmpty(snapshot.title, snapshot.topic, snapshot.name, meeting.title, meeting.topic, meeting.name),
  });
}

function snapshotSortValue(snapshot = {}) {
  return Number(observedAtMs(snapshot, Number.MAX_SAFE_INTEGER));
}

function normalizedSnapshots(input = {}, options = {}) {
  return normalizeMeetingAppSnapshots(snapshotArray(input), options)
    .map((snapshot, index) => compactObject({
      ...snapshot,
      observedAtMs: observedAtMs(snapshot, options.observedAtMs ?? options.observed_at_ms),
      __pipeline_index: index,
    }))
    .filter((snapshot) => meetingFromSnapshot(snapshot, options))
    .sort((left, right) => snapshotSortValue(left) - snapshotSortValue(right) || left.__pipeline_index - right.__pipeline_index)
    .map(({ __pipeline_index, ...snapshot }) => snapshot);
}

function primaryPlatform(snapshots = [], options = {}) {
  const platform = firstNonEmpty(options.platform, options.provider, snapshots[0]?.platform, snapshots[0]?.meeting?.platform);
  return platform ? normalizeMeetingPlatform(platform) : null;
}

function samePipelineMeeting(snapshot = {}, meeting = {}) {
  const current = meetingFromSnapshot(snapshot);
  return Boolean(current
    && current.platform === meeting.platform
    && String(current.meeting_id) === String(meeting.meeting_id));
}

function participantRows(snapshot = {}) {
  return asArray(firstNonEmpty(snapshot.participants, snapshot.roster, snapshot.members, snapshot.people, []));
}

function participantSnapshot(snapshot = {}, options = {}) {
  const meeting = meetingFromSnapshot(snapshot, options);
  const atMs = observedAtMs(snapshot, options.observedAtMs ?? options.observed_at_ms);
  return compactObject({
    meeting,
    observed_at_ms: atMs,
    participants: participantRows(snapshot),
    source_event_id: `${meeting?.platform}:${meeting?.meeting_id}:${atMs}:roster`,
    raw: snapshot,
  });
}

function speakerSample(snapshot = {}, options = {}) {
  const meeting = meetingFromSnapshot(snapshot, options);
  const atMs = observedAtMs(snapshot, options.observedAtMs ?? options.observed_at_ms);
  const speaker = snapshot.activeSpeaker ?? snapshot.active_speaker ?? null;
  return compactObject({
    meeting,
    observedAtMs: atMs,
    activeSpeaker: speaker ? compactObject({
      id: firstNonEmpty(speaker.id, speaker.speaker_id, speaker.speakerId, speaker.participant_id, speaker.participantId),
      name: firstNonEmpty(speaker.name, speaker.display_name, speaker.displayName, speaker.speaker_name, speaker.speakerName),
      speaking: speaker.speaking ?? speaker.isSpeaking ?? speaker.is_speaking ?? true,
    }) : undefined,
    source_event_id: `${meeting?.platform}:${meeting?.meeting_id}:${atMs}:speaker`,
    raw: snapshot,
  });
}

function optionGroup(options = {}, camel, snake) {
  return {
    ...(options[camel] ?? {}),
    ...(options[snake] ?? {}),
  };
}

function sortMarks(marks = []) {
  return [...marks].sort((left, right) => (
    Number(left.captured_at_ms ?? 0) - Number(right.captured_at_ms ?? 0)
    || String(left.intent ?? '').localeCompare(String(right.intent ?? ''))
    || String(left.id ?? '').localeCompare(String(right.id ?? ''))
  ));
}

export function buildMeetingAppTrackPipeline(input = {}, options = {}) {
  const snapshots = normalizedSnapshots(input, options);
  const platform = primaryPlatform(snapshots, options);
  const firstMeeting = snapshots.map((snapshot) => meetingFromSnapshot(snapshot, options)).find(Boolean);
  const scopedSnapshots = firstMeeting
    ? snapshots.filter((snapshot) => samePipelineMeeting(snapshot, firstMeeting))
    : snapshots;
  const speakerSamples = scopedSnapshots.map((snapshot) => speakerSample(snapshot, options));
  const rosterSnapshots = scopedSnapshots.map((snapshot) => participantSnapshot(snapshot, options));
  const speakerTrack = platform
    ? buildMeetingPlatformSpeakerTrack(platform, {
      samples: speakerSamples,
    }, {
      source: 'meeting_app_track_pipeline',
      ...optionGroup(options, 'speakerTrackOptions', 'speaker_track_options'),
    })
    : null;
  const participantTrack = platform
    ? buildMeetingPlatformParticipantTrack(platform, {
      snapshots: rosterSnapshots,
    }, {
      source: 'meeting_app_track_pipeline',
      ...optionGroup(options, 'participantTrackOptions', 'participant_track_options'),
    })
    : null;
  const marks = sortMarks([
    ...(speakerTrack?.marks ?? []),
    ...(participantTrack?.marks ?? []),
  ]);
  const coverage = {
    normalized_snapshot_count: snapshots.length,
    scoped_snapshot_count: scopedSnapshots.length,
    platform_detected: Boolean(platform),
    meeting_bound: Boolean(firstMeeting),
    speaker_sample_count: speakerSamples.length,
    roster_snapshot_count: rosterSnapshots.length,
    speaker_track_mark_count: speakerTrack?.mark_count ?? 0,
    participant_track_mark_count: participantTrack?.mark_count ?? 0,
    mark_count: marks.length,
    transcript_required: false,
    provider_event_required: false,
  };
  return compactObject({
    type: 'meeting_app_track_pipeline',
    schema: MEETING_APP_TRACK_PIPELINE_SCHEMA,
    schema_version: MEETING_APP_TRACK_PIPELINE_SCHEMA_VERSION,
    status: marks.length > 0 ? 'track_marks_ready' : 'no_track_marks',
    platform,
    meeting: firstMeeting,
    snapshot_count: scopedSnapshots.length,
    mark_count: marks.length,
    snapshots: scopedSnapshots,
    speaker_samples: speakerSamples,
    participant_snapshots: rosterSnapshots,
    speaker_track: speakerTrack,
    participant_track: participantTrack,
    marks,
    coverage,
    next_actions: marks.length > 0
      ? ['insert_track_marks_on_current_axis']
      : ['collect_more_meeting_app_snapshots_with_active_speaker_or_roster_changes'],
  });
}

export function createMeetingAppTrackPipeline(options = {}) {
  let snapshots = [];
  return {
    push(input = {}, pushOptions = {}) {
      snapshots.push(...snapshotArray(input));
      return buildMeetingAppTrackPipeline(snapshots, {
        ...options,
        ...pushOptions,
      });
    },
    build(buildOptions = {}) {
      return buildMeetingAppTrackPipeline(snapshots, {
        ...options,
        ...buildOptions,
      });
    },
    getSnapshots() {
      return [...snapshots];
    },
    reset(nextSnapshots = []) {
      snapshots = snapshotArray(nextSnapshots);
      return this.build();
    },
  };
}
