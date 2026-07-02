import assert from 'node:assert/strict';

import {
  captureMeetingAppDomSnapshot,
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

function fakeDocument({ url, title, nodes = [], hidden = false }) {
  return {
    nodeType: 9,
    title,
    hidden,
    location: { href: url },
    querySelectorAll(selector) {
      const text = String(selector);
      if (text === 'button') return nodes.filter((item) => item.tagName === 'BUTTON');
      if (text.includes('role="button"')) return nodes.filter((item) => item.attributes.role === 'button');
      if (text.includes('aria-label') && !text.includes('*=')) {
        return nodes.filter((item) => item.attributes['aria-label']);
      }
      if (text.includes('[title]')) return nodes.filter((item) => item.attributes.title);
      if (text.includes('data-participant-id')) {
        return nodes.filter((item) => item.attributes['data-participant-id']);
      }
      if (text.includes('data-user-id')) {
        return nodes.filter((item) => item.attributes['data-user-id']);
      }
      if (text.includes('speaking')) {
        return nodes.filter((item) => /speaking|正在发言/i.test(item.attributes['aria-label'] ?? ''));
      }
      if (text.includes('aria-live')) return nodes.filter((item) => item.attributes['aria-live']);
      if (text.includes('role="status"')) return nodes.filter((item) => item.attributes.role === 'status');
      return [];
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
