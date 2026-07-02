import assert from 'node:assert/strict';

import {
  ingestPlatformEvent,
  normalizePlatformEvent,
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

const normalizedLark = normalizePlatformEvent('feishu', larkStartEvent);
assert.equal(normalizedLark.platform, 'lark');
assert.equal(normalizedLark.source, 'lark_webhook');
assert.equal(normalizedLark.signals[0].meeting.meeting_id, 'lark-meeting-001');
assert.equal(normalizedLark.signals[0].meeting.minute_token, 'minute-token-001');

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

await assert.rejects(
  () => ingestPlatformEvent(client, 'unknown-platform', {}),
  /Unsupported meeting platform/,
);

console.log('ok meeting timeline SDK platform ingest');
