import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import { MEETING_PLATFORM_KEYS, normalizeMeetingPlatform, platformEventEndpoint } from './platform-setup.mjs';

export const PLATFORM_FIXTURE_SIGNAL_TYPES = Object.freeze([
  'meeting_start',
  'meeting_end',
  'participant_join',
  'participant_left',
  'speaker_activity',
  'transcript_ready',
  'recording_ready',
  'subscription_lifecycle',
]);

export const DEFAULT_PLATFORM_FIXTURE_SIGNAL_TYPES = Object.freeze({
  local_detector: Object.freeze(['meeting_start', 'meeting_end', 'participant_join', 'participant_left', 'speaker_activity']),
  lark: Object.freeze(['meeting_start', 'meeting_end', 'participant_join', 'participant_left', 'transcript_ready']),
  google_meet: Object.freeze(['meeting_start', 'meeting_end', 'participant_join', 'participant_left', 'transcript_ready', 'subscription_lifecycle']),
  microsoft_teams: Object.freeze(['meeting_start', 'meeting_end', 'participant_join', 'participant_left', 'transcript_ready', 'subscription_lifecycle']),
  zoom: Object.freeze(['meeting_start', 'meeting_end', 'participant_join', 'participant_left', 'transcript_ready']),
  webex: Object.freeze(['meeting_start', 'meeting_end', 'participant_join', 'participant_left', 'transcript_ready']),
});

const DEFAULT_BASE_URL = 'https://timeline.example.com';
const DEFAULT_START_MS = 1_782_442_800_000;
const DEFAULT_DURATION_MS = 30 * 60 * 1000;

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function normalizeSignalType(value) {
  const text = String(value || '')
    .trim()
    .replace(/([a-z0-9])([A-Z])/g, '$1_$2')
    .replace(/[\s.-]+/g, '_')
    .toLowerCase();
  const aliases = {
    start: 'meeting_start',
    started: 'meeting_start',
    meeting_started: 'meeting_start',
    meeting_start: 'meeting_start',
    end: 'meeting_end',
    ended: 'meeting_end',
    meeting_ended: 'meeting_end',
    meeting_end: 'meeting_end',
    join: 'participant_join',
    joined: 'participant_join',
    participant_joined: 'participant_join',
    participant_join: 'participant_join',
    leave: 'participant_left',
    left: 'participant_left',
    participant_leave: 'participant_left',
    participant_left: 'participant_left',
    speaker: 'speaker_activity',
    active_speaker: 'speaker_activity',
    speaker_started: 'speaker_activity',
    speaker_activity: 'speaker_activity',
    transcript: 'transcript_ready',
    transcript_created: 'transcript_ready',
    transcript_ready: 'transcript_ready',
    recording: 'recording_ready',
    recording_created: 'recording_ready',
    recording_ready: 'recording_ready',
    lifecycle: 'subscription_lifecycle',
    subscription: 'subscription_lifecycle',
    subscription_lifecycle: 'subscription_lifecycle',
  };
  const normalized = aliases[text] ?? text;
  if (!PLATFORM_FIXTURE_SIGNAL_TYPES.includes(normalized)) {
    throw new MeetingTimelineSdkError(`Unsupported fixture signal type: ${String(value || '(empty)')}`, {
      signal_type: value,
      supported_signal_types: PLATFORM_FIXTURE_SIGNAL_TYPES,
    });
  }
  return normalized;
}

function fixtureClock(options = {}) {
  const startMs = Number(firstNonEmpty(options.startMs, options.start_ms, DEFAULT_START_MS));
  const durationMs = Number(firstNonEmpty(options.durationMs, options.duration_ms, DEFAULT_DURATION_MS));
  return {
    startMs,
    endMs: startMs + durationMs,
    participantJoinMs: startMs + Number(firstNonEmpty(options.participantOffsetMs, options.participant_offset_ms, 30_000)),
    participantLeftMs: startMs + Number(firstNonEmpty(options.participantLeftOffsetMs, options.participant_left_offset_ms, durationMs - 30_000)),
    speakerMs: startMs + Number(firstNonEmpty(options.speakerOffsetMs, options.speaker_offset_ms, 60_000)),
    artifactMs: startMs + durationMs + Number(firstNonEmpty(options.artifactOffsetMs, options.artifact_offset_ms, 120_000)),
    lifecycleMs: startMs + Number(firstNonEmpty(options.lifecycleOffsetMs, options.lifecycle_offset_ms, 5 * 60 * 1000)),
  };
}

function iso(ms) {
  return new Date(ms).toISOString();
}

function fixtureIds(platform, options = {}) {
  const normalized = normalizeMeetingPlatform(platform);
  return {
    title: firstNonEmpty(options.title, 'SDK fixture meeting'),
    participantId: firstNonEmpty(options.participantId, options.participant_id, 'fixture-user-001'),
    participantName: firstNonEmpty(options.participantName, options.participant_name, 'Ada Lovelace'),
    speakerId: firstNonEmpty(options.speakerId, options.speaker_id, 'fixture-speaker-001'),
    speakerName: firstNonEmpty(options.speakerName, options.speaker_name, 'Ada Lovelace'),
    localMeetingId: firstNonEmpty(options.localMeetingId, options.local_meeting_id, 'local-fixture-001'),
    larkMeetingId: firstNonEmpty(options.larkMeetingId, options.lark_meeting_id, 'lark-fixture-001'),
    googleRecordId: firstNonEmpty(options.googleRecordId, options.google_record_id, 'google-fixture-001'),
    teamsMeetingId: firstNonEmpty(options.teamsMeetingId, options.teams_meeting_id, 'teams-fixture-001'),
    zoomUuid: firstNonEmpty(options.zoomUuid, options.zoom_uuid, 'zoom-fixture-001'),
    zoomNumericId: firstNonEmpty(options.zoomNumericId, options.zoom_numeric_id, 987654321),
    webexMeetingId: firstNonEmpty(options.webexMeetingId, options.webex_meeting_id, 'webex-fixture-001'),
    minuteToken: firstNonEmpty(options.minuteToken, options.minute_token, `${normalized}-minute-001`),
    transcriptId: firstNonEmpty(options.transcriptId, options.transcript_id, `${normalized}-transcript-001`),
    recordingId: firstNonEmpty(options.recordingId, options.recording_id, `${normalized}-recording-001`),
  };
}

function meetingUrls(options = {}) {
  return {
    local: firstNonEmpty(options.localUrl, options.local_url, 'https://meet.google.com/abc-defg-hij'),
    lark: firstNonEmpty(options.larkUrl, options.lark_url, 'https://vc.feishu.cn/j/123456789'),
    google: firstNonEmpty(options.googleUrl, options.google_url, 'https://meet.google.com/abc-defg-hij'),
    teams: firstNonEmpty(options.teamsUrl, options.teams_url, 'https://teams.microsoft.com/l/meetup-join/fixture'),
    zoom: firstNonEmpty(options.zoomUrl, options.zoom_url, 'https://zoom.us/j/987654321'),
    webex: firstNonEmpty(options.webexUrl, options.webex_url, 'https://example.webex.com/meet/fixture'),
  };
}

function localDetectorFixture(signalType, options) {
  const ids = fixtureIds('local_detector', options);
  const urls = meetingUrls(options);
  const clock = fixtureClock(options);
  const meeting = {
    detected_platform: firstNonEmpty(options.detectedPlatform, options.detected_platform, 'google_meet'),
    meeting_id: ids.localMeetingId,
    meeting_url: urls.local,
    title: ids.title,
  };
  const base = {
    id: `local-${signalType}-1`,
    source_event_id: `local-${signalType}-1`,
    ...meeting,
  };
  if (signalType === 'meeting_start') {
    return { ...base, type: 'meeting_started', start_time_ms: clock.startMs };
  }
  if (signalType === 'meeting_end') {
    return { ...base, type: 'meeting_ended', end_time_ms: clock.endMs };
  }
  if (signalType === 'participant_join') {
    return {
      ...base,
      type: 'participant_joined',
      occurred_at_ms: clock.participantJoinMs,
      participant: { id: ids.participantId, displayName: ids.participantName },
    };
  }
  if (signalType === 'participant_left') {
    return {
      ...base,
      type: 'participant_left',
      occurred_at_ms: clock.participantLeftMs,
      participant: { id: ids.participantId, displayName: ids.participantName },
    };
  }
  if (signalType === 'speaker_activity') {
    return {
      ...base,
      type: 'speaker_started',
      occurred_at_ms: clock.speakerMs,
      speaker: { id: ids.speakerId, displayName: ids.speakerName },
      participant: { id: ids.participantId, displayName: ids.participantName },
    };
  }
  return null;
}

function larkFixture(signalType, options) {
  const ids = fixtureIds('lark', options);
  const urls = meetingUrls(options);
  const clock = fixtureClock(options);
  const eventTypes = {
    meeting_start: 'vc.meeting.all_meeting_started_v1',
    meeting_end: 'vc.meeting.all_meeting_ended_v1',
    participant_join: 'vc.meeting.join_meeting_v1',
    participant_left: 'vc.meeting.leave_meeting_v1',
    transcript_ready: 'vc.meeting.minute_ready_v1',
    recording_ready: 'vc.meeting.recording_ready_v1',
  };
  const occurredAtMs = {
    meeting_start: clock.startMs,
    meeting_end: clock.endMs,
    participant_join: clock.participantJoinMs,
    participant_left: clock.participantLeftMs,
    transcript_ready: clock.artifactMs,
    recording_ready: clock.artifactMs,
  }[signalType];
  return {
    header: {
      event_id: `lark-${signalType}-1`,
      event_type: eventTypes[signalType],
      create_time: iso(occurredAtMs),
    },
    event: compactObject({
      meeting: {
        meeting_id: ids.larkMeetingId,
        open_meeting_id: ids.larkMeetingId,
        meeting_url: urls.lark,
        topic: ids.title,
        start_time: iso(clock.startMs),
        end_time: signalType === 'meeting_end' ? iso(clock.endMs) : undefined,
      },
      participant: ['participant_join', 'participant_left'].includes(signalType) ? {
        id: ids.participantId,
        display_name: ids.participantName,
      } : undefined,
      minute: signalType === 'transcript_ready' ? {
        token: ids.minuteToken,
        url: 'https://minutes.feishu.cn/minutes/fixture',
      } : undefined,
      recording: signalType === 'recording_ready' ? {
        id: ids.recordingId,
        url: 'https://minutes.feishu.cn/recordings/fixture',
      } : undefined,
    }),
  };
}

function googleMeetFixture(signalType, options) {
  const ids = fixtureIds('google_meet', options);
  const urls = meetingUrls(options);
  const clock = fixtureClock(options);
  const conferenceRecordName = `conferenceRecords/${ids.googleRecordId}`;
  const eventTypes = {
    meeting_start: 'google.workspace.meet.conference.v2.started',
    meeting_end: 'google.workspace.meet.conference.v2.ended',
    participant_join: 'google.workspace.meet.participant.v2.joined',
    participant_left: 'google.workspace.meet.participant.v2.left',
    transcript_ready: 'google.workspace.meet.transcript.v2.fileGenerated',
    recording_ready: 'google.workspace.meet.recording.v2.fileGenerated',
    subscription_lifecycle: 'google.workspace.events.subscription.v1.expirationReminder',
  };
  const occurredAtMs = {
    meeting_start: clock.startMs,
    meeting_end: clock.endMs,
    participant_join: clock.participantJoinMs,
    participant_left: clock.participantLeftMs,
    transcript_ready: clock.artifactMs,
    recording_ready: clock.artifactMs,
    subscription_lifecycle: clock.lifecycleMs,
  }[signalType];
  if (signalType === 'subscription_lifecycle') {
    return {
      id: `google-${signalType}-1`,
      type: eventTypes[signalType],
      time: iso(occurredAtMs),
      subject: `subscriptions/google-fixture-subscription`,
      data: {
        subscription: {
          name: 'subscriptions/google-fixture-subscription',
          targetResource: '//meet.googleapis.com/conferenceRecords/*',
          expireTime: iso(clock.endMs + 24 * 60 * 60 * 1000),
        },
      },
    };
  }
  const data = {
    conferenceRecord: { name: conferenceRecordName },
    meetingUri: urls.google,
    title: ids.title,
  };
  if (['participant_join', 'participant_left'].includes(signalType)) {
    data.participantSession = {
      name: `${conferenceRecordName}/participants/${ids.participantId}/participantSessions/session-1`,
      participant: { displayName: ids.participantName },
    };
  }
  if (signalType === 'transcript_ready') {
    data.transcript = {
      name: `${conferenceRecordName}/transcripts/${ids.transcriptId}`,
      docsDestination: { document: 'https://docs.google.com/document/d/google-fixture-transcript' },
    };
  }
  if (signalType === 'recording_ready') {
    data.recording = {
      name: `${conferenceRecordName}/recordings/${ids.recordingId}`,
      driveDestination: { exportUri: 'https://drive.google.com/file/d/google-fixture-recording' },
    };
  }
  return {
    id: `google-${signalType}-1`,
    type: eventTypes[signalType],
    time: iso(occurredAtMs),
    data,
  };
}

function teamsFixture(signalType, options) {
  const ids = fixtureIds('microsoft_teams', options);
  const urls = meetingUrls(options);
  const clock = fixtureClock(options);
  if (signalType === 'subscription_lifecycle') {
    return {
      id: 'teams-subscription-lifecycle-1',
      lifecycleEvent: 'reauthorizationRequired',
      subscriptionId: 'teams-fixture-subscription',
      resource: `/communications/onlineMeetings(joinWebUrl='${encodeURIComponent(urls.teams)}')/meetingCallEvents`,
      expirationDateTime: iso(clock.endMs + 60 * 60 * 1000),
    };
  }
  const eventTypes = {
    meeting_start: 'callStarted',
    meeting_end: 'callEnded',
    participant_join: 'rosterUpdated',
    participant_left: 'rosterUpdated',
    transcript_ready: 'callTranscriptCreated',
    recording_ready: 'callRecordingAvailable',
  };
  const occurredAtMs = {
    meeting_start: clock.startMs,
    meeting_end: clock.endMs,
    participant_join: clock.participantJoinMs,
    participant_left: clock.participantLeftMs,
    transcript_ready: clock.artifactMs,
    recording_ready: clock.artifactMs,
  }[signalType];
  const participant = compactObject({
    id: ids.participantId,
    displayName: ids.participantName,
    joinDateTime: signalType === 'participant_join' ? iso(clock.participantJoinMs) : undefined,
    leaveDateTime: signalType === 'participant_left' ? iso(clock.participantLeftMs) : undefined,
  });
  return {
    id: `teams-${signalType}-1`,
    subscriptionId: 'teams-fixture-subscription',
    resource: `/communications/onlineMeetings(joinWebUrl='${encodeURIComponent(urls.teams)}')/meetingCallEvents`,
    resourceData: compactObject({
      id: ['transcript_ready', 'recording_ready'].includes(signalType) ? (signalType === 'transcript_ready' ? ids.transcriptId : ids.recordingId) : undefined,
      eventType: eventTypes[signalType],
      eventDateTime: iso(occurredAtMs),
      onlineMeetingId: ids.teamsMeetingId,
      joinWebUrl: urls.teams,
      subject: ids.title,
      participants: ['participant_join', 'participant_left'].includes(signalType) ? [participant] : undefined,
    }),
  };
}

function zoomFixture(signalType, options) {
  const ids = fixtureIds('zoom', options);
  const urls = meetingUrls(options);
  const clock = fixtureClock(options);
  const eventNames = {
    meeting_start: 'meeting.started',
    meeting_end: 'meeting.ended',
    participant_join: 'meeting.participant_joined',
    participant_left: 'meeting.participant_left',
    transcript_ready: 'recording.transcript_completed',
    recording_ready: 'recording.completed',
  };
  const occurredAtMs = {
    meeting_start: clock.startMs,
    meeting_end: clock.endMs,
    participant_join: clock.participantJoinMs,
    participant_left: clock.participantLeftMs,
    transcript_ready: clock.artifactMs,
    recording_ready: clock.artifactMs,
  }[signalType];
  const object = compactObject({
    uuid: ids.zoomUuid,
    id: ids.zoomNumericId,
    topic: ids.title,
    join_url: urls.zoom,
    start_time: iso(clock.startMs),
    end_time: signalType === 'meeting_end' ? iso(clock.endMs) : undefined,
    participant: ['participant_join', 'participant_left'].includes(signalType) ? compactObject({
      id: ids.participantId,
      user_name: ids.participantName,
      join_time: signalType === 'participant_join' ? iso(clock.participantJoinMs) : undefined,
      leave_time: signalType === 'participant_left' ? iso(clock.participantLeftMs) : undefined,
    }) : undefined,
    transcript: signalType === 'transcript_ready' ? {
      download_url: 'https://zoom.example/transcript.vtt',
    } : undefined,
    recording_files: signalType === 'recording_ready' ? [{
      id: ids.recordingId,
      file_type: 'MP4',
      download_url: 'https://zoom.example/recording.mp4',
    }] : undefined,
  });
  return {
    event: eventNames[signalType],
    event_ts: occurredAtMs,
    payload: { object },
  };
}

function webexFixture(signalType, options) {
  const ids = fixtureIds('webex', options);
  const urls = meetingUrls(options);
  const clock = fixtureClock(options);
  const resources = {
    meeting_start: ['meetings', 'started'],
    meeting_end: ['meetings', 'ended'],
    participant_join: ['meetingParticipants', 'joined'],
    participant_left: ['meetingParticipants', 'left'],
    transcript_ready: ['meetingTranscripts', 'created'],
    recording_ready: ['recordings', 'created'],
  };
  const [resource, event] = resources[signalType];
  const occurredAtMs = {
    meeting_start: clock.startMs,
    meeting_end: clock.endMs,
    participant_join: clock.participantJoinMs,
    participant_left: clock.participantLeftMs,
    transcript_ready: clock.artifactMs,
    recording_ready: clock.artifactMs,
  }[signalType];
  return {
    id: `webex-${signalType}-1`,
    resource,
    event,
    created: iso(occurredAtMs),
    data: compactObject({
      id: resource === 'meetings' ? ids.webexMeetingId : `${resource}-fixture-001`,
      meetingId: ids.webexMeetingId,
      title: ids.title,
      webLink: urls.webex,
      startTime: signalType === 'meeting_start' ? iso(clock.startMs) : undefined,
      endTime: signalType === 'meeting_end' ? iso(clock.endMs) : undefined,
      joinTime: signalType === 'participant_join' ? iso(clock.participantJoinMs) : undefined,
      leaveTime: signalType === 'participant_left' ? iso(clock.participantLeftMs) : undefined,
      personId: ['participant_join', 'participant_left'].includes(signalType) ? ids.participantId : undefined,
      displayName: ['participant_join', 'participant_left'].includes(signalType) ? ids.participantName : undefined,
      transcriptId: signalType === 'transcript_ready' ? ids.transcriptId : undefined,
      txtDownloadLink: signalType === 'transcript_ready' ? 'https://webex.example/transcript.vtt' : undefined,
      recordingId: signalType === 'recording_ready' ? ids.recordingId : undefined,
      downloadUrl: signalType === 'recording_ready' ? 'https://webex.example/recording.mp4' : undefined,
    }),
  };
}

const BUILDERS = Object.freeze({
  local_detector: localDetectorFixture,
  lark: larkFixture,
  google_meet: googleMeetFixture,
  microsoft_teams: teamsFixture,
  zoom: zoomFixture,
  webex: webexFixture,
});

export function buildPlatformFixtureEvent(platform, signalType, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const type = normalizeSignalType(signalType);
  const supported = DEFAULT_PLATFORM_FIXTURE_SIGNAL_TYPES[key] ?? [];
  if (!supported.includes(type)) {
    throw new MeetingTimelineSdkError(`Fixture signal type ${type} is not supported for ${key}`, {
      platform: key,
      signal_type: type,
      supported_signal_types: supported,
    });
  }
  const body = BUILDERS[key](type, options);
  if (!body) {
    throw new MeetingTimelineSdkError(`Fixture signal type ${type} did not produce a payload for ${key}`, {
      platform: key,
      signal_type: type,
    });
  }
  return body;
}

export function buildPlatformFixtureSamples(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const signalTypes = (options.signalTypes ?? options.signal_types ?? DEFAULT_PLATFORM_FIXTURE_SIGNAL_TYPES[key])
    .map((item) => normalizeSignalType(item));
  return signalTypes.map((signalType) => ({
    platform: key,
    label: `${key}:${signalType}`,
    body: buildPlatformFixtureEvent(key, signalType, options),
  }));
}

export function buildAllPlatformFixtureSamples(options = {}) {
  const platforms = options.platforms ?? MEETING_PLATFORM_KEYS;
  return Object.fromEntries(platforms.map((platform) => {
    const key = normalizeMeetingPlatform(platform);
    return [key, buildPlatformFixtureSamples(key, options)];
  }));
}

export function buildPlatformFixtureEnv(options = {}) {
  const baseUrl = options.baseUrl ?? options.base_url ?? DEFAULT_BASE_URL;
  return {
    GOOGLE_PUBSUB_OIDC_AUDIENCE: platformEventEndpoint(baseUrl, 'google_meet'),
    GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL: 'pubsub-fixture@timeline-fixture.iam.gserviceaccount.com',
    MICROSOFT_GRAPH_CLIENT_STATE: 'timeline-fixture-client-state',
    ZOOM_WEBHOOK_SECRET_TOKEN: 'timeline-fixture-zoom-secret',
    WEBEX_WEBHOOK_SECRET: 'timeline-fixture-webex-secret',
    ...(options.env ?? {}),
  };
}

export function buildPlatformFixtureAcceptanceInput(options = {}) {
  const baseUrl = options.baseUrl ?? options.base_url ?? DEFAULT_BASE_URL;
  return {
    baseUrl,
    env: buildPlatformFixtureEnv({ ...options, baseUrl }),
    samples: buildAllPlatformFixtureSamples(options),
  };
}
