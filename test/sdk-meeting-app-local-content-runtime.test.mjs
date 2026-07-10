import assert from 'node:assert/strict';

import {
  createMeetingAppLocalContentRuntime,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-local-content-runtime.mjs';

const startMs = 1_783_651_000_000;

function node(tagName, attributes = {}, text = '') {
  return {
    tagName: tagName.toUpperCase(),
    attributes,
    dataset: Object.fromEntries(Object.entries(attributes)
      .filter(([key]) => key.startsWith('data-'))
      .map(([key, value]) => [key.slice(5).replace(/-([a-z])/g, (_, char) => char.toUpperCase()), value])),
    innerText: text,
    textContent: text,
    getAttribute(name) {
      return attributes[name] ?? null;
    },
  };
}

function selectorMatches(item, selector) {
  const text = String(selector);
  if (text === '*') return true;
  if (text === 'button') return item.tagName === 'BUTTON';
  if (text.includes('role="button"') && item.attributes.role === 'button') return true;
  if (text === '[aria-label]' && item.attributes['aria-label']) return true;
  if (text.includes('[title]') && item.attributes.title) return true;
  if (text.includes('data-participant-id') && item.attributes['data-participant-id']) return true;
  if (text.includes('speaking') && /speaking|active speaker/i.test(item.attributes['aria-label'] ?? '')) return true;
  if (text.includes('aria-live') && item.attributes['aria-live']) return true;
  if (text.includes('role="status"') && item.attributes.role === 'status') return true;
  const attributeSelectors = [...text.matchAll(/\[([a-zA-Z0-9_-]+)([*]?=)?(?:"([^"]*)"|'([^']*)'|([^\]\s]+))?(?:\s+i)?\]/g)];
  if (attributeSelectors.length > 0) {
    return attributeSelectors.every((match) => {
      const [, name, operator, doubleQuoted, singleQuoted, bare] = match;
      const actual = item.attributes[name];
      if (operator == null) return actual != null;
      if (actual == null) return false;
      const expected = doubleQuoted ?? singleQuoted ?? bare ?? '';
      if (operator === '*=') return String(actual).toLowerCase().includes(String(expected).toLowerCase());
      return String(actual) === String(expected);
    });
  }
  return false;
}

const nodes = [
  node('button', { 'aria-label': 'Turn off microphone' }),
  node('button', { 'aria-label': 'Leave call' }),
  node('div', { 'data-participant-id': 'ada', 'aria-label': 'Ada Lovelace is speaking' }),
];
const listeners = new Map();
const document = {
  nodeType: 9,
  title: 'P0 runtime test - Google Meet',
  location: { href: 'https://meet.google.com/abc-defg-hij' },
  body: {},
  querySelectorAll(selector) {
    return nodes.filter((item) => selectorMatches(item, selector));
  },
};
const window = {
  document,
  location: document.location,
  addEventListener(type, listener) {
    listeners.set(type, listener);
  },
  removeEventListener(type) {
    listeners.delete(type);
  },
};

const calls = [];
const client = {
  async startMeeting(input) {
    calls.push({ method: 'startMeeting', input });
    return { ok: true };
  },
  async endMeeting(input) {
    calls.push({ method: 'endMeeting', input });
    return { ok: true };
  },
  async insertMark(input) {
    calls.push({ method: 'insertMark', input });
    return { ok: true };
  },
  async insertMarks(input) {
    calls.push({ method: 'insertMarks', input });
    return { ok: true };
  },
};

const runtime = createMeetingAppLocalContentRuntime(client, {
  window,
  document,
  location: document.location,
  platform: 'google_meet',
  speakerOptions: { minStableMs: 0, switchStableMs: 0, endIdleMs: 0 },
  windowMessaging: true,
});

const preflight = runtime.preflight();
assert.equal(preflight.accepted, true);
assert.equal(preflight.platform, 'google_meet');
assert.equal(preflight.meeting_id, 'abc-defg-hij');

const sample = await runtime.sample({ observedAtMs: startMs, trigger: 'test' });
assert.equal(sample.error, undefined);
assert.equal(calls.some((call) => call.method === 'startMeeting'), true);

const inserted = await runtime.handleMessage({
  type: 'meeting_timeline.insert_mark',
  payload: { id: 'mark-1', captured_at_ms: startMs + 100 },
});
assert.equal(inserted.handled, true);
assert.equal(calls.at(-1).method, 'insertMark');
assert.equal(calls.at(-1).input.id, 'mark-1');

const state = runtime.start();
assert.equal(state.running, true);
assert.equal(state.mutation_observer_installed, false);
assert.equal(listeners.has('message'), true);
runtime.stop();
assert.equal(runtime.getState().running, false);
assert.equal(listeners.size, 0);

console.log('ok meeting app local content runtime');
