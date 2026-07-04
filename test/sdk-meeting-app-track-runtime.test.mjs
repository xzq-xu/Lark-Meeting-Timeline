import assert from 'node:assert/strict';

import {
  createMeetingAppTrackRuntime,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-track-runtime.mjs';

const startMs = 1_783_356_000_000;

function snapshot(observedAtMs, {
  activeSpeaker = null,
  participants = [],
} = {}) {
  return {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    title: 'Design review - Google Meet',
    observedAtMs,
    inMeeting: true,
    activeSpeaker,
    participants,
  };
}

const ada = { id: 'ada', name: 'Ada Lovelace', speaking: true };
const bob = { id: 'bob', name: 'Bob Stone', speaking: false };
const grace = { id: 'grace', name: 'Grace Hopper', speaking: true };

const snapshots = [
  snapshot(startMs, {
    activeSpeaker: ada,
    participants: [ada, bob],
  }),
  snapshot(startMs + 300, {
    activeSpeaker: ada,
    participants: [ada, bob],
  }),
  snapshot(startMs + 1_400, {
    participants: [{ ...ada, speaking: false }, bob],
  }),
  snapshot(startMs + 2_000, {
    activeSpeaker: grace,
    participants: [{ ...ada, speaking: false }, bob, grace],
  }),
  snapshot(startMs + 2_300, {
    activeSpeaker: grace,
    participants: [{ ...ada, speaking: false }, bob, grace],
  }),
  snapshot(startMs + 3_300, {
    participants: [{ ...ada, speaking: false }, { ...grace, speaking: false }],
  }),
  snapshot(startMs + 4_000, {
    participants: [{ ...ada, speaking: false }, { ...grace, speaking: false }],
  }),
];

const calls = [];
const client = {
  async insertMark(input, options) {
    calls.push({ method: 'insertMark', input, options });
    return { ok: true, input };
  },
  async insertMarks(inputs, options) {
    calls.push({ method: 'insertMarks', inputs, options });
    return { ok: true, inputs };
  },
};

const runtime = createMeetingAppTrackRuntime(client, {
  speakerTrackOptions: {
    minStableMs: 250,
    switchStableMs: 250,
    endIdleMs: 500,
    minSegmentMs: 800,
  },
  participantTrackOptions: {
    leaveStableMs: 500,
    closePendingLeavesAtMs: startMs + 4_500,
  },
});

let observed = await runtime.observe(snapshots.slice(0, 2));
assert.equal(observed.schema, 'meeting_app_track_runtime');
assert.equal(observed.new_mark_count, 0);
assert.equal(calls.length, 0);

const preview = runtime.preview(snapshots.slice(2));
assert.equal(preview.new_mark_count, 4);
assert.equal(runtime.getState().seen_mark_count, 0);

observed = await runtime.observe(snapshots.slice(2));
assert.equal(observed.inserted, true);
assert.equal(observed.new_mark_count, 4);
assert.equal(observed.duplicate_mark_count, 0);
assert.equal(observed.result_count, 1);
assert.equal(calls.length, 1);
assert.equal(calls[0].method, 'insertMarks');
assert.deepEqual(calls[0].inputs.map((mark) => mark.intent), [
  'speaker_track',
  'participant_track',
  'speaker_track',
  'participant_track',
]);
assert.equal(runtime.getState().inserted_mark_count, 4);
assert.equal(runtime.getState().seen_mark_count, 4);

observed = await runtime.observe([]);
assert.equal(observed.new_mark_count, 0);
assert.equal(observed.duplicate_mark_count, 4);
assert.equal(calls.length, 1);

runtime.reset();
observed = await runtime.observe(snapshots, { dryRun: true });
assert.equal(observed.inserted, false);
assert.equal(observed.new_mark_count, 4);
assert.equal(runtime.getState().seen_mark_count, 0);
assert.equal(calls.length, 1);

observed = await runtime.observe([], { insertOptions: { source: 'runtime-test' } });
assert.equal(observed.inserted, true);
assert.equal(observed.new_mark_count, 4);
assert.equal(calls.length, 2);
assert.equal(calls[1].options.source, 'runtime-test');

const singleCalls = [];
const singleRuntime = createMeetingAppTrackRuntime({
  insertMark: async (input, options) => {
    singleCalls.push({ input, options });
    return { ok: true, input };
  },
}, {
  speakerTrackOptions: {
    minStableMs: 250,
    minSegmentMs: 0,
    closeOpenSegmentsAtMs: startMs + 1_000,
  },
});
const singleObserved = await singleRuntime.observe([
  snapshot(startMs, {
    activeSpeaker: ada,
    participants: [ada],
  }),
  snapshot(startMs + 400, {
    activeSpeaker: ada,
    participants: [ada],
  }),
]);
assert.equal(singleObserved.new_mark_count, 1);
assert.equal(singleCalls.length, 1);
assert.equal(singleCalls[0].input.intent, 'speaker_track');

console.log('ok meeting app track runtime');
