import assert from 'node:assert/strict';

import {
  MEETING_APP_DOM_CAPTURE_PROFILES,
  captureMeetingAppDomSnapshot,
  meetingAppDomCaptureProfile,
  normalizeCapturedMeetingAppDomSnapshot,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-capture.mjs';
import { createMeetingAppObserver } from '../packages/meeting-timeline-sdk/adapters/meeting-apps.mjs';
import { createMeetingSourceAggregator } from '../packages/meeting-timeline-sdk/adapters/meeting-source.mjs';

const startMs = 1_783_096_800_000;

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

function selectorAttrMatches(item, selector) {
  if (selector === 'button') return item.tagName === 'BUTTON';
  if (selector.startsWith('.')) {
    return String(item.attributes.class ?? '').split(/\s+/).includes(selector.slice(1));
  }
  const attrParts = [...String(selector).matchAll(/\[([a-zA-Z0-9_-]+)([*]?=)?(?:"([^"]*)"|'([^']*)'|([^\]\s]+))?(?:\s+i)?\]/g)];
  if (!attrParts.length) return false;
  return attrParts.every((match) => {
    const [, attrName, operator, doubleQuoted, singleQuoted, bare] = match;
    const actual = item.attributes[attrName];
    if (operator == null) return actual != null;
    if (actual == null) return false;
    const expected = doubleQuoted ?? singleQuoted ?? bare ?? '';
    if (operator === '*=') return String(actual).toLowerCase().includes(String(expected).toLowerCase());
    return String(actual) === String(expected);
  });
}

function queryNodes(nodes, selector) {
  const text = String(selector);
  if (text === '*') return nodes;
  if (text === 'iframe, frame') return nodes.filter((item) => ['IFRAME', 'FRAME'].includes(item.tagName));
  const generic = nodes.filter((item) => selectorAttrMatches(item, text));
  if (generic.length) return generic;
  if (text === 'button') return nodes.filter((item) => item.tagName === 'BUTTON');
  if (text.includes('role="button"')) return nodes.filter((item) => item.attributes.role === 'button');
  if (text === '[aria-label]') return nodes.filter((item) => item.attributes['aria-label']);
  if (text.includes('[title]')) return nodes.filter((item) => item.attributes.title);
  if (text.includes('data-participant-id')) {
    return nodes.filter((item) => item.attributes['data-participant-id']);
  }
  if (text.includes('data-user-id')) {
    return nodes.filter((item) => item.attributes['data-user-id']);
  }
  if (text.includes('data-person-id')) {
    return nodes.filter((item) => item.attributes['data-person-id']);
  }
  if (text.includes('data-display-name')) {
    return nodes.filter((item) => item.attributes['data-display-name']);
  }
  if (text.includes('speaking')) {
    return nodes.filter((item) => /speaking|active speaker|正在发言|正在讲话|正在说话/i.test(item.attributes['aria-label'] ?? ''));
  }
  if (text.includes('aria-live')) return nodes.filter((item) => item.attributes['aria-live']);
  if (text.includes('role="status"')) return nodes.filter((item) => item.attributes.role === 'status');
  return [];
}

function fakeShadowRoot(nodes = []) {
  return {
    nodeType: 11,
    querySelectorAll(selector) {
      return queryNodes(nodes, selector);
    },
  };
}

function fakeDocument({ url, title, nodes = [], hidden = false }) {
  return {
    nodeType: 9,
    title,
    hidden,
    location: { href: url },
    querySelectorAll(selector) {
      return queryNodes(nodes, selector);
    },
  };
}

const googleDoc = fakeDocument({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Design review - Google Meet',
  nodes: [
    node('button', { 'aria-label': 'Turn off microphone' }),
    node('button', { 'aria-label': 'Leave call' }),
    node('div', {
      'data-participant-id': 'ada',
      'aria-label': 'Ada Lovelace is speaking',
      'data-audio-level': '0.84',
    }),
    node('div', {
      'data-participant-id': 'grace',
      'aria-label': 'Grace Hopper, muted',
    }),
    node('div', { role: 'status', 'aria-live': 'polite' }, 'You are presenting'),
  ],
});

assert.equal(MEETING_APP_DOM_CAPTURE_PROFILES.google_meet.platform, 'google_meet');
assert.equal(meetingAppDomCaptureProfile('google-meet').platform, 'google_meet');
assert.equal(
  meetingAppDomCaptureProfile({ url: 'https://teams.microsoft.com/l/meetup-join/abc' }).platform,
  'microsoft_teams',
);

const captured = captureMeetingAppDomSnapshot({ document: googleDoc }, {
  observedAtMs: startMs,
  browserName: 'Chrome',
});
assert.equal(captured.schema, 'meeting_app_dom_capture');
assert.equal(captured.url, 'https://meet.google.com/abc-defg-hij');
assert.equal(captured.page.buttons.length, 2);
assert.equal(captured.page.tiles.length, 2);
assert.equal(captured.page.tiles[0].id, 'ada');
assert.equal(captured.page.tiles[0].audioLevel, 0.84);
assert.equal(captured.capture.participant_count, 2);
assert.equal(captured.capture.profile, 'google_meet');
assert.equal(captured.inMeeting, true);
assert.equal(captured.page.inMeeting, true);
assert.equal(captured.page.interaction.in_call, true);
assert.equal(captured.page.interaction.can_leave, true);
assert.equal(captured.page.interaction.microphone_control_available, true);
assert.equal(captured.page.interaction.screen_share_active, true);
assert.equal(captured.page.interaction.active_speaker_candidate.id, 'ada');
assert.equal(captured.semanticSignalTypes.includes('meeting_leave_available'), true);
assert.equal(captured.controlSignalSummary.leave_available, true);
assert.equal(captured.controlSignalSummary.microphone_available, true);
assert.equal(captured.controlSignalSummary.screen_share_active, true);
assert.equal(captured.capture.control_signal_summary.active_speaker_observed, true);
assert.ok(captured.page.semanticSignals.some((signal) => signal.type === 'meeting_leave_available'));
assert.ok(captured.page.semanticSignals.some((signal) => signal.type === 'screen_share_active'));

const profileOnlyGoogleCapture = captureMeetingAppDomSnapshot({
  document: fakeDocument({
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Profile-only capture - Google Meet',
    nodes: [
      node('div', { 'data-tooltip': 'Leave call' }),
      node('div', {
        'data-avatar-tooltip': 'Ada Lovelace',
        'data-is-speaking': 'true',
        'data-tile-id': 'ada-tile',
      }),
    ],
  }),
}, {
  observedAtMs: startMs + 500,
  captureProfile: 'google_meet',
});
assert.equal(profileOnlyGoogleCapture.capture.profile, 'google_meet');
assert.equal(profileOnlyGoogleCapture.page.buttons[0].label, 'Leave call');
assert.equal(profileOnlyGoogleCapture.page.tiles[0].id, 'ada-tile');
assert.equal(profileOnlyGoogleCapture.page.tiles[0].name, 'Ada Lovelace');
assert.equal(profileOnlyGoogleCapture.page.tiles[0].speaking, true);
assert.equal(profileOnlyGoogleCapture.page.interaction.active_speaker_candidate.name, 'Ada Lovelace');

const shadowHost = node('meet-shell');
shadowHost.shadowRoot = fakeShadowRoot([
  node('div', { 'data-tooltip': 'Leave call' }),
  node('div', {
    'data-avatar-tooltip': 'Shadow Ada',
    'data-is-speaking': 'true',
    'data-tile-id': 'shadow-ada',
  }),
]);
const shadowDoc = fakeDocument({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Shadow capture - Google Meet',
  nodes: [shadowHost],
});
const shallowShadowCapture = captureMeetingAppDomSnapshot({ document: shadowDoc }, {
  observedAtMs: startMs + 600,
  captureProfile: 'google_meet',
});
assert.equal(shallowShadowCapture.capture.shadow_root_count, undefined);
assert.equal(shallowShadowCapture.page.buttons.length, 0);
assert.equal(shallowShadowCapture.page.tiles.length, 0);
const deepShadowCapture = captureMeetingAppDomSnapshot({ document: shadowDoc }, {
  observedAtMs: startMs + 700,
  captureProfile: 'google_meet',
  includeShadowDom: true,
});
assert.equal(deepShadowCapture.capture.shadow_root_count, 1);
assert.equal(deepShadowCapture.page.buttons[0].label, 'Leave call');
assert.equal(deepShadowCapture.page.tiles[0].id, 'shadow-ada');
assert.equal(deepShadowCapture.page.tiles[0].name, 'Shadow Ada');

const normalized = normalizeCapturedMeetingAppDomSnapshot({ document: googleDoc }, {
  observedAtMs: startMs,
});
assert.equal(normalized.platform, 'google_meet');
assert.equal(normalized.meeting_id, 'abc-defg-hij');
assert.equal(normalized.inMeeting, true);
assert.equal(normalized.activeSpeaker.name, 'Ada Lovelace');

let observed = createMeetingAppObserver({
  source: 'browser_dom_capture',
  speakerOptions: { minStableMs: 0 },
}).observe(captured, {
  observedAtMs: startMs,
});
assert.deepEqual(observed.signals.map((item) => item.type), ['meeting_started', 'speaker_started']);
assert.equal(observed.signals[1].speaker_id, 'ada');

const teamsPrejoin = normalizeCapturedMeetingAppDomSnapshot({
  document: fakeDocument({
    url: 'https://teams.microsoft.com/l/meetup-join/19%3Ameeting_abc%40thread.v2/0',
    title: 'Weekly sync | Microsoft Teams',
    nodes: [
      node('button', { 'aria-label': 'Join now' }),
      node('div', { role: 'status' }, 'Choose your audio and video settings'),
    ],
  }),
}, {
  observedAtMs: startMs + 1_000,
});
assert.equal(teamsPrejoin.platform, 'microsoft_teams');
assert.equal(teamsPrejoin.inMeeting, false);
assert.equal(teamsPrejoin.page.interaction.pre_join, true);
assert.equal(teamsPrejoin.page.interaction.can_join, true);
assert.equal(teamsPrejoin.page.controlSignalSummary.join_available, true);
assert.ok(teamsPrejoin.page.semanticSignals.some((signal) => signal.type === 'meeting_join_available'));

const teamsChineseLive = normalizeCapturedMeetingAppDomSnapshot({
  document: fakeDocument({
    url: 'https://teams.live.com/v2/',
    title: '开会 | Microsoft Teams 会议 | Microsoft Teams',
    nodes: [
      node('button', { 'aria-label': '打开摄像头' }, '摄像头'),
      node('button', { 'aria-label': '将麦克风静音' }, '麦克风'),
      node('button', { 'aria-label': '共享内容' }, '共享'),
      node('button', { 'aria-label': '退出' }, '离开'),
      node('div', { role: 'status', 'aria-live': 'polite' }, '正在等待其他人加入…'),
    ],
  }),
}, {
  observedAtMs: startMs + 1_500,
  captureProfile: 'microsoft_teams',
  requireInMeetingEvidence: true,
});
assert.equal(teamsChineseLive.platform, 'microsoft_teams');
assert.equal(teamsChineseLive.inMeeting, true);
assert.equal(teamsChineseLive.page.interaction.in_call, true);
assert.equal(teamsChineseLive.page.interaction.can_leave, true);
assert.equal(teamsChineseLive.page.controlSignalSummary.leave_available, true);
assert.equal(teamsChineseLive.page.controlSignalSummary.microphone_available, true);
assert.equal(teamsChineseLive.page.controlSignalSummary.camera_available, true);
assert.equal(teamsChineseLive.page.controlSignalSummary.screen_share_available, true);

const zoomCaptured = captureMeetingAppDomSnapshot({
  document: fakeDocument({
    url: 'https://us06web.zoom.us/wc/987654321/start',
    title: 'Zoom Meeting',
    nodes: [
      node('button', { 'aria-label': 'Leave Meeting' }),
      node('button', { 'aria-label': 'Participants' }),
      node('div', { 'data-user-id': 'mira', 'aria-label': 'Mira Patel is speaking' }),
    ],
  }),
}, {
  observedAtMs: startMs + 2_000,
});
observed = createMeetingAppObserver({
  source: 'browser_dom_capture',
  speakerOptions: { minStableMs: 0 },
}).observe(zoomCaptured, {
  observedAtMs: startMs + 2_000,
});
assert.deepEqual(observed.signals.map((item) => item.type), ['meeting_started', 'speaker_started']);
assert.equal(observed.signals[0].meeting.platform, 'zoom');
assert.equal(observed.signals[1].speaker_name, 'Mira Patel');
assert.equal(zoomCaptured.page.interaction.can_leave, true);
assert.equal(zoomCaptured.page.interaction.participant_roster_observed, true);
assert.equal(zoomCaptured.page.controlSignalSummary.participants_available, true);
assert.ok(zoomCaptured.page.semanticSignals.some((signal) => signal.type === 'participants_control'));

const zoomChineseCaptured = captureMeetingAppDomSnapshot({
  document: fakeDocument({
    url: 'https://app.zoom.us/wc/987654321/join',
    title: 'Test Zoom Meeting',
    nodes: [
      node('button', { 'aria-label': '离开' }, '离开'),
      node('button', { 'aria-label': 'open the participants list pane' }, '1 参会者'),
      node('button', { 'aria-label': 'mute my microphone' }, '静音'),
    ],
  }),
}, {
  observedAtMs: startMs + 2_500,
  captureProfile: 'zoom',
});
assert.equal(zoomChineseCaptured.inMeeting, true);
assert.equal(zoomChineseCaptured.page.interaction.can_leave, true);
assert.equal(zoomChineseCaptured.page.controlSignalSummary.participants_available, true);
assert.equal(zoomChineseCaptured.page.interaction.microphone_control_available, true);

const zoomMeetingFrame = node('iframe');
zoomMeetingFrame.contentDocument = fakeDocument({
  url: 'https://app.zoom.us/wc/987654321/join?from=pwa',
  title: 'Web Zoom Meeting',
  nodes: [
    node('button', { 'aria-label': '离开' }, '离开'),
    node('button', { 'aria-label': 'open the participants list pane' }, '1 参会者'),
    node('div', { 'data-user-id': 'self', 'aria-label': 'Timeline Adapter Acceptance is speaking' }),
  ],
});
const zoomFrameCaptured = captureMeetingAppDomSnapshot({
  document: fakeDocument({
    url: 'https://app.zoom.us/wc/987654321/join?fromPWA=1',
    title: 'Test Zoom Meeting',
    nodes: [zoomMeetingFrame],
  }),
}, {
  observedAtMs: startMs + 2_700,
  captureProfile: 'zoom',
  includeSameOriginFrames: true,
});
assert.equal(zoomFrameCaptured.inMeeting, true);
assert.equal(zoomFrameCaptured.capture.frame_document_count, 1);
assert.equal(zoomFrameCaptured.page.interaction.active_speaker_candidate.id, 'self');

const larkCaptured = captureMeetingAppDomSnapshot({
  document: fakeDocument({
    url: 'https://vc.feishu.cn/j/123456789',
    title: '会议进展实时可视化 - 飞书',
    nodes: [
      node('button', { 'aria-label': '挂断' }),
      node('button', { 'aria-label': 'AI 视图' }),
      node('div', {
        'data-user-id': 'xzq',
        'data-display-name': '徐智强',
        'aria-label': '徐智强 正在发言',
      }),
    ],
  }),
}, {
  observedAtMs: startMs + 3_000,
});
assert.equal(larkCaptured.page.tiles[0].id, 'xzq');
assert.equal(larkCaptured.page.tiles[0].name, '徐智强');
observed = createMeetingAppObserver({
  source: 'browser_dom_capture',
  speakerOptions: { minStableMs: 0 },
}).observe(larkCaptured, {
  observedAtMs: startMs + 3_000,
});
assert.deepEqual(observed.signals.map((item) => item.type), ['meeting_started', 'speaker_started']);
assert.equal(observed.signals[0].meeting.platform, 'lark');
assert.equal(observed.signals[1].speaker_name, '徐智强');
assert.equal(larkCaptured.page.interaction.ai_summary_available, true);
assert.equal(larkCaptured.page.controlSignalSummary.ai_summary_available, true);

const webexCaptured = normalizeCapturedMeetingAppDomSnapshot({
  document: fakeDocument({
    url: 'https://example.webex.com/meet/product-review',
    title: 'Product review - Webex',
    nodes: [
      node('button', { 'aria-label': 'Leave meeting' }),
      node('button', { 'aria-label': 'Unmute' }),
      node('div', {
        'data-person-id': 'maya',
        'data-display-name': 'Maya Chen',
        'aria-label': 'Maya Chen, active speaker',
      }),
    ],
  }),
}, {
  observedAtMs: startMs + 4_000,
});
assert.equal(webexCaptured.platform, 'webex');
assert.equal(webexCaptured.meeting_id, 'meet-product-review');
assert.equal(webexCaptured.inMeeting, true);
assert.equal(webexCaptured.activeSpeaker.id, 'maya');
assert.equal(webexCaptured.activeSpeaker.name, 'Maya Chen');
assert.equal(webexCaptured.page.interaction.can_leave, true);
assert.equal(webexCaptured.page.controlSignalSummary.microphone_available, true);

const calls = [];
const source = createMeetingSourceAggregator({
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
}, {
  applyOptions: { speakerAsAnnotation: true },
  speakerOptions: { minStableMs: 0 },
});

const applied = await source.observeMeetingApp(captured, { observedAtMs: startMs });
assert.deepEqual(applied.signals.map((item) => item.type), ['meeting_started', 'speaker_started']);
assert.deepEqual(calls.map((item) => item.method), ['startMeeting', 'insertMark']);

console.log('ok meeting app DOM capture');
