import { MeetingTimelineSdkError, compactObject, normalizeAbsoluteMs } from '../index.mjs';
import {
  MEETING_APP_FIXTURE_PLATFORMS,
  buildMeetingAppFixtureSnapshot,
  diagnoseMeetingAppFixture,
} from './meeting-app-fixtures.mjs';
import { normalizeMeetingAppSnapshot } from './meeting-apps.mjs';
import { buildMeetingPlatformParticipantTrack } from './platform-participant-track.mjs';
import { normalizeMeetingPlatform } from './platform-setup.mjs';
import { buildMeetingPlatformSpeakerTrack } from './platform-speaker-track.mjs';

export const MEETING_APP_FIXTURE_TRACK_READINESS_SCHEMA = 'meeting_app_fixture_track_readiness';
export const MEETING_APP_FIXTURE_TRACK_READINESS_REPORT_SCHEMA = 'meeting_app_fixture_track_readiness_report';
export const MEETING_APP_FIXTURE_TRACK_READINESS_SCHEMA_VERSION = 1;

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

function unique(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function selectedPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_APP_FIXTURE_PLATFORMS))
    .map((platform) => normalizeFixturePlatform(platform)));
}

function normalizeFixturePlatform(platform) {
  const key = normalizeMeetingPlatform(platform);
  if (!MEETING_APP_FIXTURE_PLATFORMS.includes(key)) {
    throw new MeetingTimelineSdkError(`Unsupported meeting app fixture track platform: ${String(platform || '(empty)')}`, {
      platform,
      supported_platforms: MEETING_APP_FIXTURE_PLATFORMS,
    });
  }
  return key;
}

function cloneJson(value) {
  return JSON.parse(JSON.stringify(value));
}

function snapshotObservedAtMs(snapshot = {}, fallback) {
  return normalizeAbsoluteMs(firstNonEmpty(
    snapshot.observedAtMs,
    snapshot.observed_at_ms,
    snapshot.timestamp_ms,
    fallback,
  ), 'meeting_app_fixture_track_observed_at_ms');
}

function participantKey(participant = {}) {
  return firstNonEmpty(
    participant.id,
    participant.participant_id,
    participant.participantId,
    participant.user_id,
    participant.userId,
    participant.name,
    participant.display_name,
    participant.displayName,
    participant.label,
  );
}

function rosterSnapshotFromFixture(snapshot = {}, options = {}) {
  const normalized = normalizeMeetingAppSnapshot(snapshot, options);
  const atMs = snapshotObservedAtMs(normalized ?? snapshot, snapshot.observedAtMs);
  return compactObject({
    meeting: {
      platform: normalized?.platform,
      meeting_id: normalized?.meeting_id,
      meeting_url: normalized?.meeting_url ?? normalized?.url ?? snapshot.meeting_url ?? snapshot.url,
      start_time_ms: atMs,
    },
    observed_at_ms: atMs,
    participants: normalized?.participants ?? snapshot.page?.participants ?? snapshot.dom?.participants ?? [],
    source_event_id: `${snapshot.id ?? normalized?.platform}:${atMs}:roster`,
  });
}

function fixtureWithParticipantRemoved(snapshot = {}, removedKey) {
  const next = cloneJson(snapshot);
  const shouldKeep = (participant = {}) => String(participantKey(participant)) !== String(removedKey);
  if (Array.isArray(next.page?.participants)) next.page.participants = next.page.participants.filter(shouldKeep);
  if (Array.isArray(next.page?.tiles)) next.page.tiles = next.page.tiles.filter(shouldKeep);
  if (Array.isArray(next.dom?.participants)) next.dom.participants = next.dom.participants.filter(shouldKeep);
  return next;
}

function fixtureStartMs(options = {}) {
  return normalizeAbsoluteMs(firstNonEmpty(
    options.observedAtMs,
    options.observed_at_ms,
    options.startMs,
    options.start_ms,
    1_783_356_000_000,
  ), 'meeting_app_fixture_track_start_ms');
}

function coverageKeys() {
  return [
    'participant_roster_snapshot',
    'speaker_track_mark',
    'participant_track_mark',
    'speaker_track_payload_text_free',
    'participant_track_payload_text_free',
  ];
}

function markPayloadTextFree(track = {}) {
  return asArray(track.marks).every((mark) => !mark.payload?.transcript && !mark.payload?.text);
}

export function diagnoseMeetingAppFixtureTrackReadiness(platform, options = {}) {
  const key = normalizeFixturePlatform(platform);
  const startMs = fixtureStartMs(options);
  const endMs = normalizeAbsoluteMs(firstNonEmpty(
    options.endObservedAtMs,
    options.end_observed_at_ms,
    startMs + Number(firstNonEmpty(options.endOffsetMs, options.end_offset_ms, 2_000)),
  ), 'meeting_app_fixture_track_end_observed_at_ms');
  const activeSnapshot = buildMeetingAppFixtureSnapshot(key, {
    ...options,
    state: 'active',
    observedAtMs: startMs,
  });
  const diagnosis = diagnoseMeetingAppFixture(key, {
    ...options,
    state: 'active',
    observedAtMs: startMs,
  });
  const normalized = diagnosis.normalized;
  const speakerTrack = buildMeetingPlatformSpeakerTrack(key, {
    signals: diagnosis.signals,
  }, {
    minSegmentMs: 0,
    closeOpenSegmentsAtMs: startMs + 1_500,
    ...(options.speakerTrackOptions ?? {}),
    ...(options.speaker_track_options ?? {}),
  });
  const participants = normalized?.participants ?? [];
  const activeSpeakerKey = firstNonEmpty(normalized?.activeSpeaker?.id, normalized?.activeSpeaker?.name);
  const removedParticipant = participants.find((participant) => String(participantKey(participant)) !== String(activeSpeakerKey))
    ?? participants[participants.length - 1];
  const participantChangedSnapshot = fixtureWithParticipantRemoved(activeSnapshot, participantKey(removedParticipant));
  participantChangedSnapshot.observedAtMs = endMs;
  participantChangedSnapshot.observed_at_ms = endMs;
  const participantTrack = buildMeetingPlatformParticipantTrack(key, {
    snapshots: [
      rosterSnapshotFromFixture(activeSnapshot, options),
      rosterSnapshotFromFixture(participantChangedSnapshot, options),
    ],
  }, {
    leaveStableMs: 0,
    closePendingLeavesAtMs: endMs + 1,
    ...(options.participantTrackOptions ?? {}),
    ...(options.participant_track_options ?? {}),
  });
  const coverage = {
    participant_roster_snapshot: participants.length > 0,
    participant_roster_count: participants.length,
    speaker_track_mark: speakerTrack.mark_count > 0,
    participant_track_mark: participantTrack.mark_count > 0,
    speaker_track_payload_text_free: markPayloadTextFree(speakerTrack),
    participant_track_payload_text_free: markPayloadTextFree(participantTrack),
  };
  const required = coverageKeys();
  const missing = required.filter((keyName) => coverage[keyName] !== true);
  return {
    type: 'meeting_app_fixture_track_readiness',
    schema: MEETING_APP_FIXTURE_TRACK_READINESS_SCHEMA,
    schema_version: MEETING_APP_FIXTURE_TRACK_READINESS_SCHEMA_VERSION,
    platform: key,
    ready: missing.length === 0,
    required,
    missing,
    snapshots: {
      active: activeSnapshot,
      participant_changed: participantChangedSnapshot,
    },
    normalized,
    speaker_track: speakerTrack,
    participant_track: participantTrack,
    coverage,
  };
}

export function buildMeetingAppFixtureTrackReadinessReport(options = {}) {
  const platforms = selectedPlatforms(options);
  const rows = platforms.map((platform) => diagnoseMeetingAppFixtureTrackReadiness(platform, options));
  const missing = rows.flatMap((row) => row.missing.map((item) => `${row.platform}:${item}`));
  return {
    type: 'meeting_app_fixture_track_readiness_report',
    schema: MEETING_APP_FIXTURE_TRACK_READINESS_REPORT_SCHEMA,
    schema_version: MEETING_APP_FIXTURE_TRACK_READINESS_SCHEMA_VERSION,
    accepted: missing.length === 0,
    required: coverageKeys(),
    missing,
    platform_count: rows.length,
    accepted_count: rows.filter((row) => row.ready === true).length,
    coverage_by_platform: Object.fromEntries(rows.map((row) => [row.platform, row.coverage])),
    rows,
  };
}
