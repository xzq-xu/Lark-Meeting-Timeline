import assert from 'node:assert/strict';

import { createMeetingTimelineBridge } from '../packages/meeting-timeline-sdk/adapters/timeline-bridge.mjs';

const startMs = 1_782_442_800_000;
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
  async insertMarks(input) {
    calls.push({ method: 'insertMarks', input });
    return { ok: true, input };
  },
  async importTranscript(input) {
    calls.push({ method: 'importTranscript', input });
    return { ok: true, input };
  },
  async getState() {
    calls.push({ method: 'getState' });
    return { ok: true };
  },
};

const bridge = createMeetingTimelineBridge(client, {
  source: 'desktop_observer',
  applyOptions: {
    speakerAsAnnotation: true,
  },
});

const localStart = await bridge.observe({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Google Meet',
  observedAtMs: startMs,
});
assert.equal(localStart.results[0].action, 'startMeeting');
assert.equal(calls.at(-1).method, 'startMeeting');
assert.equal(calls.at(-1).input.platform, 'google_meet');
assert.equal(calls.at(-1).input.detector_source, 'google_meet_local_detector');

const localDuplicate = await bridge.observe({
  url: 'https://meet.google.com/abc-defg-hij',
  observedAtMs: startMs + 1_000,
});
assert.equal(localDuplicate.results.length, 0);

const providerStart = await bridge.ingest('google-meet', {
  id: 'google-start-bridge-1',
  type: 'google.workspace.meet.conference.v2.started',
  time: new Date(startMs + 2_000).toISOString(),
  data: {
    conferenceRecord: { name: 'conferenceRecords/google-record-bridge-1' },
    meetingUri: 'https://meet.google.com/abc-defg-hij',
    title: 'Official Google Meet',
  },
});
assert.equal(providerStart.results.length, 1);
assert.equal(providerStart.reconciliation.decisions[0].reason, 'provider_reconciles_local_axis');
assert.equal(calls.at(-1).method, 'startMeeting');
assert.equal(calls.at(-1).input.meeting_id, 'google-record-bridge-1');

await bridge.insertMark({
  id: 'bridge-mark-1',
  capturedAtMs: startMs + 10_000,
  label: 'why?',
});
assert.equal(calls.at(-1).method, 'insertMark');
assert.equal(calls.at(-1).input.id, 'bridge-mark-1');

await bridge.importTranscript({
  platform: 'google_meet',
  meeting: {
    platform: 'google_meet',
    meetingId: 'google-record-bridge-1',
    startTimeMs: startMs,
  },
  raw: {
    entries: [{
      id: 'entry-1',
      startTime: new Date(startMs + 1_000).toISOString(),
      endTime: new Date(startMs + 2_000).toISOString(),
      text: 'Hello bridge',
      participant: { displayName: 'Ada' },
    }],
  },
});
assert.equal(calls.at(-1).method, 'importTranscript');
assert.equal(calls.at(-1).input.meeting.meeting_id, 'google-record-bridge-1');
assert.equal(calls.at(-1).input.transcript.length, 1);
assert.equal(calls.at(-1).input.transcript[0].text, 'Hello bridge');

const bridgeState = bridge.getBridgeState();
assert.equal(bridgeState.local_observer.activeMeeting.meeting_id, 'abc-defg-hij');
assert.equal(bridgeState.signal_reconciler.active_meetings[0].meeting_id, 'google-record-bridge-1');

const resetState = bridge.reset();
assert.equal(resetState.local_observer.activeMeeting, null);
assert.equal(resetState.signal_reconciler.active_meetings.length, 0);

console.log('ok meeting timeline bridge');
