import assert from 'node:assert/strict';

import {
  ingestPlatformEvent,
  normalizePlatformEvent,
} from '../packages/meeting-timeline-sdk/adapters/platform-ingest.mjs';

const startMs = 1_782_442_800_000;
const startIso = new Date(startMs).toISOString();

const googleStartEvent = {
  id: 'google-start-1',
  type: 'google.workspace.meet.conference.v2.started',
  time: startIso,
  data: {
    conferenceRecord: { name: 'conferenceRecords/google-record-001' },
    title: 'Google product review',
  },
};

const normalized = normalizePlatformEvent('meet', googleStartEvent);
assert.equal(normalized.platform, 'google_meet');
assert.equal(normalized.source, 'google_meet_webhook');
assert.equal(normalized.signals.length, 1);
assert.equal(normalized.signals[0].type, 'meeting_started');
assert.equal(normalized.signals[0].meeting.meeting_id, 'google-record-001');

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
