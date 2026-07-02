import assert from 'node:assert/strict';

import { createMeetingAppTimelineRuntime } from '../packages/meeting-timeline-sdk/adapters/meeting-app-runtime.mjs';

const startMs = 1_783_269_600_000;

function node(tagName, attrs = {}, text = '') {
  return {
    tagName: tagName.toUpperCase(),
    attributes: attrs,
    dataset: Object.fromEntries(Object.entries(attrs)
      .filter(([key]) => key.startsWith('data-'))
      .map(([key, value]) => [
        key.slice(5).replace(/-([a-z])/g, (_, char) => char.toUpperCase()),
        value,
      ])),
    innerText: text,
    textContent: text,
    getAttribute(name) {
      return attrs[name] ?? null;
    },
  };
}

function meetDocument() {
  const nodes = [
    node('button', { 'aria-label': 'Turn off microphone' }),
    node('button', { 'aria-label': 'Leave call' }),
    node('div', {
      'data-participant-id': 'ada',
      'aria-label': 'Ada Lovelace is speaking',
      'data-audio-level': '0.80',
    }),
  ];
  return {
    nodeType: 9,
    title: 'Design review - Google Meet',
    hidden: false,
    location: { href: 'https://meet.google.com/abc-defg-hij' },
    querySelectorAll(selector) {
      const text = String(selector);
      if (text === 'button') return nodes.filter((item) => item.tagName === 'BUTTON');
      if (text.includes('data-participant-id')) return nodes.filter((item) => item.attributes['data-participant-id']);
      if (text.includes('speaking')) return nodes.filter((item) => /speaking/i.test(item.attributes['aria-label'] ?? ''));
      return [];
    },
  };
}

let clock = startMs;
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
  async insertMark(input, options) {
    calls.push({ method: 'insertMark', input, options });
    return { ok: true, input };
  },
  async insertMarks(inputs, options) {
    calls.push({ method: 'insertMarks', inputs, options });
    return { ok: true, inputs };
  },
};

const runtime = createMeetingAppTimelineRuntime(client, {
  now: () => clock,
  applyOptions: { speakerAsAnnotation: true },
  speakerOptions: { minStableMs: 0 },
  captureOptions: { browserName: 'Chrome' },
  unchangedObserveEveryMs: 300,
});

const sample = await runtime.sample({ document: meetDocument() });
assert.equal(sample.emitted, true);
assert.equal(sample.result.source, 'meeting_app');
assert.deepEqual(sample.result.signals.map((item) => item.type), ['meeting_started', 'speaker_started']);
assert.deepEqual(calls.map((item) => item.method), ['startMeeting', 'insertMark']);
assert.equal(calls[0].input.platform, 'google_meet');
assert.equal(calls[0].input.meeting_id, 'abc-defg-hij');
assert.equal(calls[1].input.kind, 'speaker_started');
assert.equal(calls[1].input.payload.speaker_name, 'Ada Lovelace');

await runtime.insertMark({
  id: 'mark-runtime-1',
  capturedAtMs: startMs + 1_000,
  kind: 'handwriting_trigger',
  label: 'why?',
});
assert.equal(calls.at(-1).method, 'insertMark');
assert.equal(calls.at(-1).input.id, 'mark-runtime-1');

clock += 2_000;
await runtime.ingestSignals({
  type: 'meeting_ended',
  meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
  },
  occurred_at_ms: clock,
  source: 'browser_dom_monitor',
});
assert.equal(calls.at(-1).method, 'endMeeting');
assert.equal(calls.at(-1).input.meeting_id, 'abc-defg-hij');

const running = runtime.start(() => ({ document: meetDocument() }), { immediate: false, sampleIntervalMs: 10_000 });
assert.equal(running.running, true);
assert.equal(runtime.stop().running, false);

const state = runtime.getState();
assert.equal(state.monitor.running, false);
assert.equal(state.sources.reconciler.active?.meeting_id, undefined);

console.log('ok meeting app timeline runtime');
