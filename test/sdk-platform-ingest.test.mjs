import assert from 'node:assert/strict';

import {
  assertMeetingPlatformProviderReplayMatrix,
  assertMeetingPlatformProviderReplayReport,
  buildMeetingPlatformProviderReplayAcceptanceReport,
  buildMeetingPlatformProviderReplayMatrix,
  buildMeetingPlatformProviderReplayReport,
  diagnosePlatformEvent,
  ingestPlatformEvent,
  normalizePlatformEvent,
  sampleMeetingPlatformProviderEvents,
} from '../packages/meeting-timeline-sdk/adapters/platform-ingest.mjs';

const startMs = 1_782_442_800_000;
const startIso = new Date(startMs).toISOString();

const larkStartEvent = {
  header: {
    event_id: 'lark-start-1',
    event_type: 'vc.meeting.all_meeting_started_v1',
    create_time: String(startMs),
  },
  event: {
    meeting: {
      id: 'lark-meeting-001',
      meeting_no: '123456789',
      topic: 'Lark product review',
      url: 'https://vc.feishu.cn/j/lark-meeting-001',
      start_time: String(Math.round(startMs / 1000)),
    },
    minute_token: 'minute-token-001',
  },
};

const googleStartEvent = {
  id: 'google-start-1',
  type: 'google.workspace.meet.conference.v2.started',
  time: startIso,
  data: {
    conferenceRecord: { name: 'conferenceRecords/google-record-001' },
    title: 'Google product review',
  },
};

const localDetectorStartEvent = {
  id: 'local-detector-start-1',
  type: 'meeting_started',
  detected_platform: 'google_meet',
  meeting_id: 'local-google-meet-001',
  meeting_url: 'https://meet.google.com/local-demo',
  title: 'Local detector Google Meet',
  start_time_ms: startMs,
};

const normalizedLocalDetector = normalizePlatformEvent('desktop-observer', localDetectorStartEvent);
assert.equal(normalizedLocalDetector.platform, 'local_detector');
assert.equal(normalizedLocalDetector.source, 'local_detector');
assert.equal(normalizedLocalDetector.signals[0].type, 'meeting_started');
assert.equal(normalizedLocalDetector.signals[0].meeting.platform, 'google_meet');
assert.equal(normalizedLocalDetector.signals[0].meeting.meeting_id, 'local-google-meet-001');
assert.equal(normalizedLocalDetector.signals[0].occurred_at_ms, startMs);

const normalized = normalizePlatformEvent('meet', googleStartEvent);
assert.equal(normalized.platform, 'google_meet');
assert.equal(normalized.source, 'google_meet_webhook');
assert.equal(normalized.signals.length, 1);
assert.equal(normalized.signals[0].type, 'meeting_started');
assert.equal(normalized.signals[0].meeting.meeting_id, 'google-record-001');

const googleDiagnostics = diagnosePlatformEvent('google-meet', googleStartEvent);
assert.equal(googleDiagnostics.ok, true);
assert.equal(googleDiagnostics.supported, true);
assert.equal(googleDiagnostics.actionable, true);
assert.equal(googleDiagnostics.signal_count, 1);
assert.deepEqual(googleDiagnostics.signal_types, ['meeting_started']);
assert.equal(googleDiagnostics.coverage.meeting_start, true);
assert.equal(googleDiagnostics.meetings[0].meeting_id, 'google-record-001');
assert.equal(googleDiagnostics.signals[0].type, 'meeting_started');
assert.equal(
  googleDiagnostics.issues.some((item) => item.code === 'provider_event_missing_meeting_url'),
  true,
);

const emptyLocalDiagnostics = diagnosePlatformEvent('local-detector', {
  type: 'unsupported_local_signal',
  meeting_id: 'local-empty-001',
  occurred_at_ms: startMs,
});
assert.equal(emptyLocalDiagnostics.ok, true);
assert.equal(emptyLocalDiagnostics.actionable, false);
assert.equal(emptyLocalDiagnostics.signal_count, 0);
assert.equal(emptyLocalDiagnostics.issues[0].code, 'no_supported_signals');

const googleLifecycleDiagnostics = diagnosePlatformEvent('google-meet', {
  id: 'google-lifecycle-diagnostic-1',
  type: 'google.workspace.events.subscription.v1.expirationReminder',
  time: startIso,
  data: {
    subscription: {
      name: 'subscriptions/google-diagnostic-sub',
      expireTime: new Date(startMs + 60_000).toISOString(),
    },
  },
});
assert.equal(googleLifecycleDiagnostics.ok, true);
assert.equal(googleLifecycleDiagnostics.actionable, true);
assert.deepEqual(googleLifecycleDiagnostics.signal_types, ['subscription_lifecycle']);
assert.equal(googleLifecycleDiagnostics.coverage.subscription_lifecycle, true);
assert.equal(googleLifecycleDiagnostics.issues[0].code, 'subscription_lifecycle_only');

const unsupportedDiagnostics = diagnosePlatformEvent('unknown-platform', {});
assert.equal(unsupportedDiagnostics.ok, false);
assert.equal(unsupportedDiagnostics.supported, false);
assert.equal(unsupportedDiagnostics.actionable, false);
assert.equal(unsupportedDiagnostics.issues[0].code, 'unsupported_platform');

const normalizedLark = normalizePlatformEvent('feishu', larkStartEvent);
assert.equal(normalizedLark.platform, 'lark');
assert.equal(normalizedLark.source, 'lark_webhook');
assert.equal(normalizedLark.signals[0].meeting.meeting_id, 'lark-meeting-001');
assert.equal(normalizedLark.signals[0].meeting.minute_token, 'minute-token-001');

const rawLarkProviderReplay = buildMeetingPlatformProviderReplayReport('lark', [larkStartEvent], {
  baseReceivedAtMs: startMs,
  requiredCoverage: ['meeting_start'],
});
assert.equal(rawLarkProviderReplay.accepted, true);
assert.equal(rawLarkProviderReplay.rows[0].signal_types.includes('meeting_started'), true);
assert.equal(rawLarkProviderReplay.rows[0].meetings[0].minute_token, 'minute-token-001');

const calls = [];
const client = {
  async startMeeting(input) {
    calls.push({ method: 'startMeeting', input });
    return { ok: true, input };
  },
  async endMeeting(input) {
    calls.push({ method: 'endMeeting', input });
    return { ok: true, input };
  },
  async insertMark(input) {
    calls.push({ method: 'insertMark', input });
    return { ok: true, input };
  },
};

const startResult = await ingestPlatformEvent(client, {
  platform: 'google-meet',
  body: googleStartEvent,
});
assert.equal(startResult.platform, 'google_meet');
assert.equal(startResult.results.length, 1);
assert.equal(startResult.results[0].action, 'startMeeting');
assert.equal(calls.at(-1).method, 'startMeeting');
assert.equal(calls.at(-1).input.start_time_ms, startMs);

const localDetectorStartResult = await ingestPlatformEvent(client, 'local-detector', localDetectorStartEvent);
assert.equal(localDetectorStartResult.platform, 'local_detector');
assert.equal(localDetectorStartResult.results[0].action, 'startMeeting');
assert.equal(calls.at(-1).input.platform, 'google_meet');
assert.equal(calls.at(-1).input.detector_source, 'google_meet_local_detector');
assert.equal(calls.at(-1).input.start_time_ms, startMs);

const localDetectorSpeakerResult = await ingestPlatformEvent(
  client,
  'local-detector',
  {
    type: 'active_speaker',
    detected_platform: 'google_meet',
    meeting_id: 'local-google-meet-001',
    occurred_at_ms: startMs + 15_000,
    speaker: {
      id: 'speaker-ada',
      name: 'Ada',
    },
  },
  { speakerAsAnnotation: true },
);
assert.equal(localDetectorSpeakerResult.platform, 'local_detector');
assert.equal(localDetectorSpeakerResult.signals[0].type, 'speaker_started');
assert.equal(localDetectorSpeakerResult.results[0].action, 'insertSpeakerMark');
assert.equal(calls.at(-1).method, 'insertMark');
assert.equal(calls.at(-1).input.kind, 'speaker_started');
assert.equal(calls.at(-1).input.label, 'Ada speaking');
assert.equal(calls.at(-1).input.captured_at_ms, startMs + 15_000);

const larkStartResult = await ingestPlatformEvent(client, 'lark-suite', larkStartEvent);
assert.equal(larkStartResult.platform, 'lark');
assert.equal(larkStartResult.results[0].action, 'startMeeting');
assert.equal(calls.at(-1).input.platform, 'lark');
assert.equal(calls.at(-1).input.minute_token, 'minute-token-001');

const zoomParticipantResult = await ingestPlatformEvent(
  client,
  'zoom',
  {
    event: 'meeting.participant_joined',
    event_ts: startMs + 30_000,
    payload: {
      object: {
        uuid: 'zoom-uuid-001',
        id: 987654321,
        participant: {
          user_id: 'zoom-user-1',
          user_name: 'Lin',
          join_time: new Date(startMs + 30_000).toISOString(),
        },
      },
    },
  },
  { participantAsAnnotation: true },
);
assert.equal(zoomParticipantResult.platform, 'zoom');
assert.equal(zoomParticipantResult.results[0].action, 'insertParticipantMark');
assert.equal(calls.at(-1).method, 'insertMark');
assert.equal(calls.at(-1).input.kind, 'participant_joined');
assert.equal(calls.at(-1).input.label, 'Lin joined');

const webexTranscriptResult = normalizePlatformEvent({
  platform: 'cisco-webex',
  body: {
    id: 'webex-hook-1',
    resource: 'meetingTranscripts',
    event: 'created',
    data: {
      meetingId: 'webex-meeting-001',
      id: 'transcript-1',
      txtDownloadLink: 'https://webex.example/transcript-1.vtt',
    },
  },
  receivedAtMs: startMs + 120_000,
});
assert.equal(webexTranscriptResult.platform, 'webex');
assert.equal(webexTranscriptResult.signals[0].type, 'artifact_ready');
assert.equal(webexTranscriptResult.signals[0].occurred_at_ms, startMs + 120_000);

const googleProviderSamples = sampleMeetingPlatformProviderEvents('google-meet', {
  baseReceivedAtMs: startMs,
});
assert.equal(googleProviderSamples.length, 4);
assert.equal(googleProviderSamples[0].kind, 'meeting_start');
const googleProviderReplay = buildMeetingPlatformProviderReplayReport('google-meet', googleProviderSamples, {
  baseReceivedAtMs: startMs,
});
assert.equal(googleProviderReplay.schema, 'meeting_platform_provider_replay_report');
assert.equal(googleProviderReplay.accepted, true);
assert.equal(googleProviderReplay.platform, 'google_meet');
assert.equal(googleProviderReplay.runtime_contract.runtime_action, 'provider_event');
assert.equal(googleProviderReplay.runtime_contract.provider_events_block_realtime, false);
assert.equal(googleProviderReplay.coverage.meeting_start, true);
assert.equal(googleProviderReplay.coverage.meeting_end, true);
assert.equal(googleProviderReplay.coverage.participant_track, true);
assert.equal(googleProviderReplay.coverage.artifact_ready, true);
assert.equal(googleProviderReplay.rows.every((row) => row.runtime_event_action === 'provider_event'), true);
assert.equal(googleProviderReplay.rows.find((row) => row.kind === 'meeting_start').signal_types.includes('meeting_started'), true);
const googleProviderReplayAcceptance = buildMeetingPlatformProviderReplayAcceptanceReport(googleProviderReplay);
assert.equal(googleProviderReplayAcceptance.schema, 'meeting_platform_provider_replay_acceptance_report');
assert.equal(googleProviderReplayAcceptance.accepted, true);
assert.equal(assertMeetingPlatformProviderReplayReport(googleProviderReplay).accepted, true);

const providerReplayMatrix = buildMeetingPlatformProviderReplayMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  baseReceivedAtMs: startMs,
});
assert.equal(providerReplayMatrix.schema, 'meeting_platform_provider_replay_matrix');
assert.equal(providerReplayMatrix.accepted, true);
assert.equal(providerReplayMatrix.platform_count, 5);
assert.equal(providerReplayMatrix.accepted_count, 5);
assert.equal(providerReplayMatrix.rows.every((row) => row.coverage.meeting_start === true), true);
assert.equal(providerReplayMatrix.rows.every((row) => row.coverage.meeting_end === true), true);
assert.equal(providerReplayMatrix.rows.every((row) => row.provider_events_block_realtime === false), true);
assert.equal(providerReplayMatrix.rows.find((row) => row.platform === 'microsoft_teams').signal_types.includes('participant_joined'), true);
assert.equal(assertMeetingPlatformProviderReplayMatrix(providerReplayMatrix).accepted, true);

const missingEndReplay = buildMeetingPlatformProviderReplayReport('zoom', [
  sampleMeetingPlatformProviderEvents('zoom', { baseReceivedAtMs: startMs })[0],
], {
  baseReceivedAtMs: startMs,
});
assert.equal(missingEndReplay.accepted, false);
assert.equal(missingEndReplay.issues.includes('missing_coverage:meeting_end'), true);
assert.throws(
  () => assertMeetingPlatformProviderReplayReport(missingEndReplay),
  /provider replay report is not accepted/,
);

await assert.rejects(
  () => ingestPlatformEvent(client, 'unknown-platform', {}),
  /Unsupported meeting platform/,
);

console.log('ok meeting timeline SDK platform ingest');
