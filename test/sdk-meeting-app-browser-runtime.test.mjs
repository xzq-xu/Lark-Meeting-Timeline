import assert from 'node:assert/strict';

import {
  createMeetingAppBrowserRuntime,
  meetingAppBrowserInput,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-browser-runtime.mjs';

const startMs = 1_783_442_400_000;

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

function fakeDocument() {
  const nodes = [
    node('button', { 'aria-label': 'Turn off microphone' }),
    node('button', { 'aria-label': 'Leave call' }),
    node('div', {
      'data-participant-id': 'ada',
      'data-display-name': 'Ada Lovelace',
      'aria-label': 'Ada Lovelace is speaking',
      'data-audio-level': '0.82',
    }),
  ];
  return {
    nodeType: 9,
    title: 'Browser runtime - Google Meet',
    hidden: false,
    location: { href: 'https://meet.google.com/abc-defg-hij' },
    querySelectorAll(selector) {
      const text = String(selector);
      if (text === 'button') return nodes.filter((item) => item.tagName === 'BUTTON');
      if (text.includes('data-participant-id')) return nodes.filter((item) => item.attributes['data-participant-id']);
      if (text.includes('data-display-name')) return nodes.filter((item) => item.attributes['data-display-name']);
      if (text.includes('speaking')) return nodes.filter((item) => /speaking/i.test(item.attributes['aria-label'] ?? ''));
      return [];
    },
  };
}

function fakePrejoinDocument() {
  const nodes = [
    node('button', { 'aria-label': 'Join now' }),
    node('div', { role: 'status' }, 'Ready to join'),
  ];
  return {
    nodeType: 9,
    title: 'Ready to join - Google Meet',
    hidden: false,
    location: { href: 'https://meet.google.com/abc-defg-hij' },
    querySelectorAll(selector) {
      const text = String(selector);
      if (text === 'button') return nodes.filter((item) => item.tagName === 'BUTTON');
      if (text.includes('role="status"')) return nodes.filter((item) => item.attributes.role === 'status');
      return [];
    },
  };
}

class FakeMutationObserver {
  static instances = [];

  constructor(callback) {
    this.callback = callback;
    this.connected = false;
    FakeMutationObserver.instances.push(this);
  }

  observe(root, options) {
    this.root = root;
    this.options = options;
    this.connected = true;
  }

  disconnect() {
    this.connected = false;
  }

  trigger(records = [{ type: 'childList' }]) {
    this.callback(records, this);
  }
}

function wait(ms = 0) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function fakeWindow(document, extra = {}) {
  const listeners = new Map();
  const win = {
    document,
    location: document.location,
    navigator: { userAgent: 'Chrome fixture' },
    ...extra,
    setDocument(nextDocument) {
      win.document = nextDocument;
      win.location = nextDocument.location;
    },
    addEventListener(eventName, handler) {
      const handlers = listeners.get(eventName) ?? [];
      handlers.push(handler);
      listeners.set(eventName, handlers);
    },
    removeEventListener(eventName, handler) {
      listeners.set(eventName, (listeners.get(eventName) ?? []).filter((item) => item !== handler));
    },
    dispatch(eventName) {
      for (const handler of listeners.get(eventName) ?? []) handler({ type: eventName });
    },
    listenerCount(eventName) {
      return (listeners.get(eventName) ?? []).length;
    },
  };
  return win;
}

let clock = startMs;
const document = fakeDocument();
const window = fakeWindow(document);
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

const input = meetingAppBrowserInput({ window });
assert.equal(input.document, document);
assert.equal(input.window, window);
assert.equal(input.url, 'https://meet.google.com/abc-defg-hij');
assert.equal(input.title, 'Browser runtime - Google Meet');

const runtime = createMeetingAppBrowserRuntime(client, {
  window,
  now: () => clock,
  applyOptions: { speakerAsAnnotation: true },
  speakerOptions: { minStableMs: 0 },
  sampleIntervalMs: 10_000,
});

const sample = await runtime.sample();
assert.equal(sample.emitted, true);
assert.deepEqual(sample.result.signals.map((item) => item.type), ['meeting_started', 'speaker_started']);
assert.deepEqual(calls.map((item) => item.method), ['startMeeting', 'insertMark']);
assert.equal(calls[0].input.platform, 'google_meet');
assert.equal(calls[0].input.meeting_id, 'abc-defg-hij');
assert.equal(calls[1].input.payload.speaker_name, 'Ada Lovelace');

window.setDocument(fakePrejoinDocument());
const prejoinSample = await runtime.sample({ force: true });
assert.equal(prejoinSample.emitted, true);
assert.equal(prejoinSample.result.signals.at(-1).type, 'meeting_ended');
assert.equal(calls.at(-1).method, 'endMeeting');
assert.equal(calls.at(-1).input.meeting_id, 'abc-defg-hij');

const running = runtime.start({ immediate: false, sampleIntervalMs: 10_000 });
assert.equal(running.running, true);
assert.equal(window.listenerCount('pagehide'), 1);
assert.equal(runtime.getState().browser_runtime.lifecycle_installed, true);

window.dispatch('pagehide');
assert.equal(runtime.getState().monitor.running, false);

const inserted = await runtime.handleMessage({
  type: 'meeting_timeline.insert_mark',
  payload: {
    mark: {
      id: 'browser-runtime-mark-1',
      capturedAtMs: startMs + 1_000,
      kind: 'handwriting_trigger',
      label: 'why?',
    },
  },
});
assert.equal(inserted.handled, true);
assert.equal(inserted.action, 'insertMark');
assert.equal(calls.at(-1).input.id, 'browser-runtime-mark-1');

const insertedBatch = await runtime.handleMessage({
  type: 'meeting_timeline.insert_marks',
  payload: {
    marks: [{
      id: 'browser-runtime-mark-2',
      capturedAtMs: startMs + 2_000,
      label: 'follow up',
    }],
  },
});
assert.equal(insertedBatch.handled, true);
assert.equal(insertedBatch.action, 'insertMarks');
assert.equal(calls.at(-1).inputs[0].id, 'browser-runtime-mark-2');

const unsupported = await runtime.handleMessage({ type: 'unknown' });
assert.equal(unsupported.handled, false);
assert.equal(unsupported.reason, 'unsupported_message_type');

clock += 2_000;
runtime.start({ immediate: false, sampleIntervalMs: 10_000 });
assert.equal(runtime.getState().monitor.running, true);
runtime.dispose();
assert.equal(runtime.getState().monitor.running, false);
assert.equal(window.listenerCount('pagehide'), 0);

FakeMutationObserver.instances = [];
const mutationCalls = [];
const mutationClient = {
  async startMeeting(input) {
    mutationCalls.push({ method: 'startMeeting', input });
    return { ok: true, input };
  },
  async endMeeting(input) {
    mutationCalls.push({ method: 'endMeeting', input });
    return { ok: true, input };
  },
  async insertMark(input, options) {
    mutationCalls.push({ method: 'insertMark', input, options });
    return { ok: true, input };
  },
};
const mutationDocument = fakeDocument();
const mutationWindow = fakeWindow(mutationDocument, { MutationObserver: FakeMutationObserver });
const mutationRuntime = createMeetingAppBrowserRuntime(mutationClient, {
  window: mutationWindow,
  now: () => clock,
  applyOptions: { speakerAsAnnotation: true },
  speakerOptions: { minStableMs: 0 },
  sampleIntervalMs: 10_000,
});
const mutationInstall = mutationRuntime.installMutationObserver({ mutationDebounceMs: 0 });
assert.equal(mutationInstall.installed, true);
assert.equal(FakeMutationObserver.instances.length, 1);
assert.equal(FakeMutationObserver.instances[0].connected, true);
assert.equal(FakeMutationObserver.instances[0].options.subtree, true);

FakeMutationObserver.instances[0].trigger([{ type: 'childList' }, { type: 'attributes' }]);
assert.equal(mutationRuntime.getState().browser_runtime.mutation_observer.pending_count, 2);
const mutationFlush = await mutationRuntime.flushMutationObserver();
assert.equal(mutationFlush.flushed, true);
assert.deepEqual(mutationCalls.map((item) => item.method), ['startMeeting', 'insertMark']);
assert.equal(mutationFlush.result.result.signals[0].type, 'meeting_started');
assert.equal(mutationRuntime.getState().browser_runtime.mutation_observer.sample_count, 1);

const flushMessage = await mutationRuntime.handleMessage({ type: 'meeting_timeline.flush_mutations' });
assert.equal(flushMessage.handled, true);
assert.equal(flushMessage.action, 'flushMutationObserver');
assert.equal(flushMessage.result.flushed, false);

mutationRuntime.stop();
assert.equal(FakeMutationObserver.instances[0].connected, false);

FakeMutationObserver.instances = [];
const filteredWindow = fakeWindow(fakeDocument(), { MutationObserver: FakeMutationObserver });
const filteredRuntime = createMeetingAppBrowserRuntime(mutationClient, {
  window: filteredWindow,
  now: () => clock,
  applyOptions: { speakerAsAnnotation: true },
  speakerOptions: { minStableMs: 0 },
  mutationIgnoreSelectors: ['.caption-line'],
  mutationTrackSelectors: ['[data-participant-id]', '[aria-label*="Leave call" i]'],
});
filteredRuntime.installMutationObserver({ mutationDebounceMs: 0 });
FakeMutationObserver.instances[0].trigger([
  { type: 'characterData', target: node('span', { class: 'caption-line' }, 'live caption noise') },
  { type: 'attributes', target: node('div', { 'data-participant-id': 'ada' }) },
]);
const filteredState = filteredRuntime.getState().browser_runtime.mutation_observer;
assert.equal(filteredState.pending_count, 1);
assert.equal(filteredState.ignored_count, 1);
await filteredRuntime.flushMutationObserver();
assert.equal(filteredRuntime.getState().browser_runtime.mutation_observer.sample_count, 1);
filteredRuntime.dispose();

FakeMutationObserver.instances = [];
let followupClock = startMs;
const followupCalls = [];
const followupClient = {
  async startMeeting(input) {
    followupCalls.push({ method: 'startMeeting', input });
    return { ok: true, input };
  },
  async endMeeting(input) {
    followupCalls.push({ method: 'endMeeting', input });
    return { ok: true, input };
  },
  async insertMark(input, options) {
    followupCalls.push({ method: 'insertMark', input, options });
    return { ok: true, input };
  },
};
const followupWindow = fakeWindow(fakeDocument(), { MutationObserver: FakeMutationObserver });
const followupRuntime = createMeetingAppBrowserRuntime(followupClient, {
  window: followupWindow,
  now: () => followupClock,
  applyOptions: { speakerAsAnnotation: true },
  speakerOptions: { minStableMs: 300 },
  speakerStableFollowupMs: 0,
  sampleIntervalMs: 10_000,
});
followupRuntime.installMutationObserver({ mutationDebounceMs: 0 });
FakeMutationObserver.instances[0].trigger([{ type: 'attributes' }]);
const followupFirstFlush = await followupRuntime.flushMutationObserver();
assert.equal(followupFirstFlush.flushed, true);
assert.deepEqual(followupFirstFlush.result.result.signals.map((item) => item.type), ['meeting_started']);
assert.equal(followupRuntime.getState().browser_runtime.mutation_observer.speaker_followup_scheduled, true);

followupClock += 300;
await wait(5);
assert.deepEqual(followupCalls.map((item) => item.method), ['startMeeting', 'insertMark']);
assert.equal(followupCalls.at(-1).input.kind, 'speaker_started');
assert.equal(followupRuntime.getState().browser_runtime.mutation_observer.sample_count, 2);
assert.equal(followupRuntime.getState().browser_runtime.mutation_observer.speaker_followup_scheduled, false);
followupRuntime.dispose();

console.log('ok meeting app browser runtime');
