import assert from 'node:assert/strict';

import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';
import { createMeetingSourceAggregator } from '../packages/meeting-timeline-sdk/adapters/meeting-source.mjs';

const startMs = 1_782_873_600_000;
const googleUrl = 'https://meet.google.com/abc-defg-hij';

const calls = [];
const client = {
  async startMeeting(payload) {
    calls.push({ method: 'startMeeting', payload });
    return { ok: true, payload };
  },
  async endMeeting(payload) {
    calls.push({ method: 'endMeeting', payload });
    return { ok: true, payload };
  },
  async insertMark(input, options) {
    calls.push({ method: 'insertMark', input, options });
    return { ok: true, input };
  },
};

const sources = createMeetingSourceAggregator(client, {
  applyOptions: { speakerAsAnnotation: true },
  speakerOptions: { minStableMs: 0, endIdleMs: 500 },
  reconcileOptions: { duplicateWindowMs: 60_000 },
});

const browserStart = await sources.observeBrowser({
  windows: [{
    id: 'browser-1',
    focused: true,
    tabs: [{
      id: 'meet-tab',
      active: true,
      audible: true,
      url: googleUrl,
      title: 'Design review - Google Meet',
      page: {
        inMeeting: true,
        activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
      },
    }],
  }],
  observedAtMs: startMs,
}, {
  observedAtMs: startMs,
});
assert.equal(browserStart.source, 'browser');
assert.deepEqual(browserStart.rawSignals.map((item) => item.type), ['meeting_started', 'speaker_started']);
assert.deepEqual(browserStart.results.map((item) => item.action), ['startMeeting', 'insertSpeakerMark']);
assert.equal(calls[0].payload.platform, 'google_meet');
assert.equal(calls[0].payload.meeting_id, 'abc-defg-hij');
assert.equal(calls[1].input.kind, 'speaker_started');

const providerStart = await sources.ingestProvider('google-meet', buildPlatformFixtureEvent('google-meet', 'meeting_start', {
  startMs,
  googleUrl,
  googleRecordId: 'conference-record-001',
}));
assert.equal(providerStart.source, 'provider');
assert.equal(providerStart.rawSignals[0].type, 'meeting_started');
assert.equal(providerStart.reconciliation.decisions[0].reason, 'provider_reconciles_local_axis');
assert.deepEqual(providerStart.results.map((item) => item.action), ['startMeeting']);
assert.equal(calls[2].payload.meeting_id, 'conference-record-001');

const duplicateLocal = await sources.observeLocal({
  url: googleUrl,
  title: 'Design review - Google Meet',
  observedAtMs: startMs + 2_000,
}, {
  observedAtMs: startMs + 2_000,
});
assert.equal(duplicateLocal.rawSignals.length, 1);
assert.equal(duplicateLocal.signals.length, 0);
assert.equal(duplicateLocal.results.length, 0);
assert.equal(duplicateLocal.reconciliation.skipped[0].reason, 'provider_axis_already_active');

const duplicateSpeaker = await sources.ingestSignals({
  type: 'speaker_started',
  meeting: {
    platform: 'google_meet',
    meeting_id: 'conference-record-001',
    meeting_url: googleUrl,
  },
  occurred_at_ms: startMs + 1_000,
  source: 'browser_meeting_observer',
  speaker_id: 'ada',
  speaker_name: 'Ada',
}, {
  speakerRepeatWindowMs: 5_000,
});
assert.equal(duplicateSpeaker.results.length, 0);
assert.equal(duplicateSpeaker.reconciliation.skipped[0].reason, 'duplicate_speaker_signal');

const nativeTeams = await sources.observeNative({
  application: { name: 'Microsoft Teams', bundleId: 'com.microsoft.teams2' },
  window: {
    id: 'teams-call-1',
    title: 'Roadmap review | Microsoft Teams',
    focused: true,
  },
  accessibility: {
    inMeeting: true,
    meetingId: 'teams-roadmap-001',
    activeSpeaker: { id: 'grace', name: 'Grace', speaking: true },
  },
  observedAtMs: startMs + 20_000,
}, {
  observedAtMs: startMs + 20_000,
});
assert.equal(nativeTeams.source, 'native');
assert.deepEqual(nativeTeams.rawSignals.map((item) => item.type), ['meeting_started', 'speaker_started']);
assert.deepEqual(nativeTeams.results.map((item) => item.action), ['startMeeting', 'insertSpeakerMark']);
assert.equal(calls.at(-2).payload.platform, 'microsoft_teams');
assert.equal(calls.at(-2).payload.meeting_id, 'teams-roadmap-001');

const state = sources.getState();
assert.equal(state.browser.sessionState.activeMeeting.platform, 'google_meet');
assert.equal(state.native.sessionState.activeMeeting.platform, 'microsoft_teams');
assert.equal(state.reconciler.active_meetings.length >= 2, true);

const reset = sources.reset();
assert.equal(reset.browser.sessionState.activeMeeting, null);
assert.equal(reset.native.sessionState.activeMeeting, null);
assert.deepEqual(reset.reconciler.active_meetings, []);

console.log('ok meeting source aggregator');
