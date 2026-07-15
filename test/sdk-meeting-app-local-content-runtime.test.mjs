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

const inactiveSurfaces = [
  {
    platform: 'google_meet',
    url: 'https://meet.google.com/xyz-abcd-efg',
    title: 'Google Meet',
    nodes: [
      node('button', { 'aria-label': 'Join now' }),
      node('button', { 'aria-label': 'Turn off microphone' }),
    ],
  },
  {
    platform: 'microsoft_teams',
    url: 'https://teams.microsoft.com/v2/',
    title: 'Microsoft Teams',
    nodes: [
      node('button', { 'aria-label': 'Retry' }),
      node('button', { 'aria-label': 'Clear cache and retry' }),
    ],
  },
  {
    platform: 'zoom',
    url: 'https://zoom.us/wc/123456789/join',
    title: 'Join Zoom Meeting',
    nodes: [
      node('button', { 'aria-label': 'Join meeting' }),
      node('button', { 'aria-label': 'Mute microphone' }),
    ],
  },
];

for (const surface of inactiveSurfaces) {
  const inactiveCalls = [];
  const inactiveDocument = {
    nodeType: 9,
    title: surface.title,
    location: { href: surface.url },
    body: {},
    querySelectorAll(selector) {
      return surface.nodes.filter((item) => selectorMatches(item, selector));
    },
  };
  const inactiveWindow = {
    document: inactiveDocument,
    location: inactiveDocument.location,
  };
  const inactiveRuntime = createMeetingAppLocalContentRuntime({
    async startMeeting(input) { inactiveCalls.push({ method: 'startMeeting', input }); },
    async endMeeting(input) { inactiveCalls.push({ method: 'endMeeting', input }); },
    async insertMark(input) { inactiveCalls.push({ method: 'insertMark', input }); },
    async insertMarks(input) { inactiveCalls.push({ method: 'insertMarks', input }); },
  }, {
    window: inactiveWindow,
    document: inactiveDocument,
    location: inactiveDocument.location,
    platform: surface.platform,
    windowMessaging: false,
  });
  const inactivePreflight = inactiveRuntime.preflight();
  assert.equal(inactivePreflight.accepted, false, `${surface.platform} inactive page must fail preflight`);
  assert.equal(inactivePreflight.reason, 'missing_in_meeting_evidence');
  await inactiveRuntime.sample({ observedAtMs: startMs + 1_000, trigger: 'inactive_surface_regression' });
  assert.equal(
    inactiveCalls.some((call) => call.method === 'startMeeting'),
    false,
    `${surface.platform} inactive page must not start a timeline`,
  );
}

const teamsLiveCalls = [];
const teamsLiveNodes = [
  node('button', { 'aria-label': '打开摄像头' }, '摄像头'),
  node('button', { 'aria-label': '将麦克风静音' }, '麦克风'),
  node('button', { 'aria-label': '共享内容' }, '共享'),
  node('button', { 'aria-label': '退出' }, '离开'),
  node('div', { role: 'status', 'aria-live': 'polite' }, '正在等待其他人加入…'),
];
const teamsLiveDocument = {
  nodeType: 9,
  title: '开会 | Microsoft Teams 会议 | Microsoft Teams',
  location: { href: 'https://teams.live.com/v2/' },
  body: {},
  querySelectorAll(selector) {
    return teamsLiveNodes.filter((item) => selectorMatches(item, selector));
  },
};
const teamsLiveWindow = {
  document: teamsLiveDocument,
  location: teamsLiveDocument.location,
};
const teamsLiveRuntime = createMeetingAppLocalContentRuntime({
  async startMeeting(input) { teamsLiveCalls.push({ method: 'startMeeting', input }); },
  async endMeeting(input) { teamsLiveCalls.push({ method: 'endMeeting', input }); },
  async insertMark(input) { teamsLiveCalls.push({ method: 'insertMark', input }); },
  async insertMarks(input) { teamsLiveCalls.push({ method: 'insertMarks', input }); },
}, {
  window: teamsLiveWindow,
  document: teamsLiveDocument,
  location: teamsLiveDocument.location,
  platform: 'microsoft_teams',
  windowMessaging: false,
});
const teamsLiveSample = await teamsLiveRuntime.sample({
  observedAtMs: startMs + 2_000,
  trigger: 'teams_live_chinese_regression',
});
assert.equal(teamsLiveSample.snapshot.inMeeting, true);
assert.equal(teamsLiveCalls.some((call) => call.method === 'startMeeting'), true);

console.log('ok meeting app local content runtime');
