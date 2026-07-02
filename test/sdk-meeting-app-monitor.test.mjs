import assert from 'node:assert/strict';

import { createMeetingAppDomMonitor, meetingAppDomSnapshotSignature } from '../packages/meeting-timeline-sdk/adapters/meeting-app-monitor.mjs';
import { captureMeetingAppDomSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-capture.mjs';
import { createMeetingAppObserver } from '../packages/meeting-timeline-sdk/adapters/meeting-apps.mjs';

const startMs = 1_783_183_200_000;

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

function fakeDocument({ speakerId = 'ada', speakerLabel = 'Ada Lovelace is speaking' } = {}) {
  const nodes = [
    node('button', { 'aria-label': 'Turn off microphone' }),
    node('button', { 'aria-label': 'Leave call' }),
    node('div', {
      'data-participant-id': speakerId,
      'aria-label': speakerLabel,
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

const firstCapture = captureMeetingAppDomSnapshot({ document: fakeDocument() }, { observedAtMs: startMs });
const sameCapture = captureMeetingAppDomSnapshot({ document: fakeDocument() }, { observedAtMs: startMs + 10_000 });
const changedCapture = captureMeetingAppDomSnapshot({
  document: fakeDocument({ speakerId: 'grace', speakerLabel: 'Grace Hopper is speaking' }),
}, { observedAtMs: startMs + 10_000 });
assert.equal(meetingAppDomSnapshotSignature(firstCapture), meetingAppDomSnapshotSignature(sameCapture));
assert.notEqual(meetingAppDomSnapshotSignature(firstCapture), meetingAppDomSnapshotSignature(changedCapture));

let clock = startMs;
const observer = createMeetingAppObserver({
  source: 'browser_dom_monitor',
  speakerOptions: {
    minStableMs: 300,
    switchStableMs: 300,
    endIdleMs: 1_000,
  },
});
const monitor = createMeetingAppDomMonitor(observer, {
  now: () => clock,
  minObserveIntervalMs: 100,
  unchangedObserveEveryMs: 300,
});

let result = await monitor.sample({ document: fakeDocument() });
assert.equal(result.emitted, true);
assert.equal(result.reason, 'first_sample');
assert.deepEqual(result.result.signals.map((item) => item.type), ['meeting_started']);
assert.equal(monitor.getState().emitCount, 1);

clock += 100;
result = await monitor.sample({ document: fakeDocument() });
assert.equal(result.emitted, false);
assert.equal(result.reason, 'unchanged_throttled');
assert.equal(monitor.getState().skippedCount, 1);

clock += 250;
result = await monitor.sample({ document: fakeDocument() });
assert.equal(result.emitted, true);
assert.equal(result.reason, 'keep_alive');
assert.deepEqual(result.result.signals.map((item) => item.type), ['speaker_started']);
assert.equal(result.result.signals[0].speaker_id, 'ada');

clock += 100;
result = await monitor.sample({
  document: fakeDocument({ speakerId: 'grace', speakerLabel: 'Grace Hopper is speaking' }),
});
assert.equal(result.emitted, true);
assert.equal(result.reason, 'changed');
assert.equal(result.result.signals.length, 0);

clock += 50;
result = await monitor.sample({
  document: fakeDocument({ speakerId: 'grace', speakerLabel: 'Grace Hopper is speaking' }),
});
assert.equal(result.emitted, false);
assert.equal(result.reason, 'unchanged_throttled');

monitor.reset();
assert.equal(monitor.getState().emitCount, 0);
assert.equal(monitor.getState().running, false);

console.log('ok meeting app DOM monitor');
