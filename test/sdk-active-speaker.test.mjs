import assert from 'node:assert/strict';

import {
  createActiveSpeakerObserver,
  createActiveSpeakerTimelineObserver,
  normalizeActiveSpeakerSample,
  observeActiveSpeakerSample,
  observeActiveSpeakerSamples,
} from '../packages/meeting-timeline-sdk/adapters/active-speaker.mjs';

const startMs = 1_782_614_400_000;
const meeting = {
  platform: 'google_meet',
  meeting_id: 'abc-defg-hij',
  meeting_url: 'https://meet.google.com/abc-defg-hij',
  title: 'Design review',
};

const ada = (observedAtMs, extra = {}) => ({
  meeting,
  activeSpeaker: {
    id: 'ada',
    name: 'Ada',
    speaking: true,
  },
  observedAtMs,
  ...extra,
});

const grace = (observedAtMs, extra = {}) => ({
  meeting,
  active_speaker: {
    id: 'grace',
    displayName: 'Grace',
    active: true,
  },
  observed_at_ms: observedAtMs,
  ...extra,
});

let normalized = normalizeActiveSpeakerSample({
  url: 'https://meet.google.com/abc-defg-hij',
  activeSpeaker: { name: 'Ada', speaking: true },
  observedAtMs: startMs,
});
assert.equal(normalized.meeting.platform, 'google_meet');
assert.equal(normalized.meeting.meeting_id, 'abc-defg-hij');
assert.equal(normalized.speaker_name, 'Ada');
assert.equal(normalized.speaking, true);

let result = observeActiveSpeakerSample(null, ada(startMs), {
  minStableMs: 300,
  source: 'browser_dom_observer',
});
assert.equal(result.signals.length, 0);
assert.equal(result.state.candidateSpeaker.speaker_id, 'ada');

result = observeActiveSpeakerSample(result.state, ada(startMs + 250), {
  minStableMs: 300,
  source: 'browser_dom_observer',
});
assert.equal(result.signals.length, 0);

result = observeActiveSpeakerSample(result.state, ada(startMs + 320), {
  minStableMs: 300,
  source: 'browser_dom_observer',
});
assert.equal(result.signals.length, 1);
assert.equal(result.signals[0].type, 'speaker_started');
assert.equal(result.signals[0].occurred_at_ms, startMs);
assert.equal(result.signals[0].source, 'browser_dom_observer');
assert.equal(result.signals[0].speaker_id, 'ada');
assert.equal(result.state.activeSpeaker.speaker_name, 'Ada');

result = observeActiveSpeakerSample(result.state, ada(startMs + 700), {
  minStableMs: 300,
});
assert.equal(result.signals.length, 0);

let switchResult = observeActiveSpeakerSample(result.state, grace(startMs + 900), {
  switchStableMs: 400,
});
assert.equal(switchResult.signals.length, 0);
assert.equal(switchResult.state.activeSpeaker.speaker_id, 'ada');
assert.equal(switchResult.state.candidateSpeaker.speaker_id, 'grace');

switchResult = observeActiveSpeakerSample(switchResult.state, ada(startMs + 1_000), {
  switchStableMs: 400,
});
assert.equal(switchResult.signals.length, 0);
assert.equal(switchResult.state.activeSpeaker.speaker_id, 'ada');
assert.equal(switchResult.state.candidateSpeaker, null);

switchResult = observeActiveSpeakerSample(switchResult.state, grace(startMs + 1_500), {
  switchStableMs: 400,
  source: 'browser_dom_observer',
});
assert.equal(switchResult.signals.length, 0);
switchResult = observeActiveSpeakerSample(switchResult.state, grace(startMs + 1_950), {
  switchStableMs: 400,
  source: 'browser_dom_observer',
});
assert.deepEqual(switchResult.signals.map((item) => item.type), ['speaker_ended', 'speaker_started']);
assert.equal(switchResult.signals[0].speaker_id, 'ada');
assert.equal(switchResult.signals[0].occurred_at_ms, startMs + 1_500);
assert.equal(switchResult.signals[1].speaker_id, 'grace');
assert.equal(switchResult.signals[1].occurred_at_ms, startMs + 1_500);
assert.equal(switchResult.state.activeSpeaker.speaker_id, 'grace');

let silence = observeActiveSpeakerSample(switchResult.state, {
  meeting,
  activeSpeaker: { id: 'grace', name: 'Grace', speaking: false },
  observedAtMs: startMs + 2_400,
}, {
  endIdleMs: 1_000,
});
assert.equal(silence.signals.length, 0);
assert.equal(silence.state.silenceSinceMs, startMs + 2_400);

silence = observeActiveSpeakerSample(silence.state, {
  meeting,
  observedAtMs: startMs + 3_450,
}, {
  endIdleMs: 1_000,
});
assert.equal(silence.signals.length, 1);
assert.equal(silence.signals[0].type, 'speaker_ended');
assert.equal(silence.signals[0].speaker_id, 'grace');
assert.equal(silence.signals[0].occurred_at_ms, startMs + 2_400);
assert.equal(silence.state.activeSpeaker, null);

const many = observeActiveSpeakerSamples(null, [
  ada(startMs + 4_000),
  ada(startMs + 4_350),
  { meeting, observedAtMs: startMs + 5_000 },
  { meeting, observedAtMs: startMs + 6_600 },
], {
  minStableMs: 300,
  endIdleMs: 1_500,
});
assert.equal(many.observations.length, 4);
assert.deepEqual(many.signals.map((item) => item.type), ['speaker_started', 'speaker_ended']);

const observer = createActiveSpeakerObserver({
  minStableMs: 0,
  emitSpeakerEnd: false,
});
const instant = observer.observe({
  meeting,
  speaker: { id: 'lin', displayName: 'Lin' },
  observedAtMs: startMs + 7_000,
});
assert.equal(instant.signals.length, 1);
assert.equal(instant.signals[0].type, 'speaker_started');
assert.equal(instant.signals[0].speaker_name, 'Lin');

const calls = [];
const timeline = createActiveSpeakerTimelineObserver({
  async insertMark(input, options) {
    calls.push({ input, options });
    return { ok: true, input };
  },
}, {
  minStableMs: 0,
  applyOptions: { speakerAsAnnotation: true },
});
const applied = await timeline.observe({
  meeting: {
    platform: 'zoom',
    meeting_id: '987654321',
    title: 'Zoom Meeting',
  },
  activeSpeaker: { name: 'Ada', speaking: true },
  observedAtMs: startMs + 8_000,
});
assert.equal(applied.signals.length, 1);
assert.equal(applied.results[0].action, 'insertSpeakerMark');
assert.equal(calls[0].input.kind, 'speaker_started');
assert.equal(calls[0].input.captured_at_ms, startMs + 8_000);
assert.equal(calls[0].input.payload.meeting.platform, 'zoom');

console.log('ok active speaker adapter');
