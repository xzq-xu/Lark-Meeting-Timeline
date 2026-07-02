import assert from 'node:assert/strict';

import {
  createBrowserMeetingObserver,
  createBrowserMeetingTimelineObserver,
  normalizeBrowserActiveSpeakerSample,
  normalizeBrowserMeetingCandidate,
  normalizeBrowserMeetingCandidates,
  observeBrowserMeetingSample,
  selectBrowserMeetingCandidate,
} from '../packages/meeting-timeline-sdk/adapters/browser-meeting.mjs';

const startMs = 1_782_700_800_000;

function meetSnapshot(observedAtMs, activeSpeaker = null, overrides = {}) {
  return {
    browser: { name: 'Chrome' },
    windows: [{
      id: 'win-1',
      focused: true,
      tabs: [{
        id: 'tab-mail',
        active: false,
        url: 'https://mail.google.com',
        title: 'Inbox',
      }, {
        id: 'tab-meet',
        active: true,
        audible: true,
        url: 'https://meet.google.com/abc-defg-hij',
        title: 'Design review - Google Meet',
        page: {
          inMeeting: true,
          documentVisible: true,
          activeSpeaker,
        },
      }],
    }],
    observedAtMs,
    ...overrides,
  };
}

let candidates = normalizeBrowserMeetingCandidates(meetSnapshot(startMs, { id: 'ada', name: 'Ada', speaking: true }), {
  observedAtMs: startMs,
});
assert.equal(candidates.length, 1);
assert.equal(candidates[0].platform, 'google_meet');
assert.equal(candidates[0].meeting_id, 'abc-defg-hij');
assert.equal(candidates[0].tab.id, 'tab-meet');
assert.equal(candidates[0].in_meeting, true);

let selected = selectBrowserMeetingCandidate(meetSnapshot(startMs), { observedAtMs: startMs });
assert.equal(selected.detectedMeeting.platform, 'google_meet');
assert.equal(selected.detectedMeeting.meeting_id, 'abc-defg-hij');

const candidate = normalizeBrowserMeetingCandidate({
  tab: {
    id: 'teams-tab',
    active: true,
    url: 'https://teams.microsoft.com/l/meetup-join/19%3Ameeting_XYZ%40thread.v2/0',
    title: 'Weekly sync | Microsoft Teams',
  },
  page: { hasJoined: true },
  observedAtMs: startMs,
});
assert.equal(candidate.platform, 'microsoft_teams');
assert.equal(candidate.meeting_id, '19:meeting_XYZ@thread.v2');
assert.equal(candidate.in_meeting, true);

const speakerSample = normalizeBrowserActiveSpeakerSample({
  tab: {
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Design review - Google Meet',
  },
  page: {
    inMeeting: true,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
  },
  observedAtMs: startMs + 100,
});
assert.equal(speakerSample.meeting.platform, 'google_meet');
assert.equal(speakerSample.meeting.meeting_id, 'abc-defg-hij');
assert.equal(speakerSample.speaker_id, 'ada');
assert.equal(speakerSample.speaking, true);

let observed = observeBrowserMeetingSample(null, meetSnapshot(startMs, { id: 'ada', name: 'Ada', speaking: true }), {
  minStableMs: 300,
  source: 'browser_extension',
  observedAtMs: startMs,
});
assert.deepEqual(observed.signals.map((item) => item.type), ['meeting_started']);
assert.equal(observed.session.signals[0].meeting.platform, 'google_meet');
assert.equal(observed.speaker.signals.length, 0);

observed = observeBrowserMeetingSample(observed.state, meetSnapshot(startMs + 350, { id: 'ada', name: 'Ada', speaking: true }), {
  minStableMs: 300,
  source: 'browser_extension',
  observedAtMs: startMs + 350,
});
assert.deepEqual(observed.signals.map((item) => item.type), ['speaker_started']);
assert.equal(observed.speaker.signals[0].speaker_id, 'ada');
assert.equal(observed.speaker.signals[0].occurred_at_ms, startMs);

observed = observeBrowserMeetingSample(observed.state, meetSnapshot(startMs + 600, { id: 'grace', name: 'Grace', speaking: true }), {
  switchStableMs: 500,
  source: 'browser_extension',
  observedAtMs: startMs + 600,
});
assert.equal(observed.speaker.signals.length, 0);

observed = observeBrowserMeetingSample(observed.state, meetSnapshot(startMs + 750, { id: 'ada', name: 'Ada', speaking: true }), {
  switchStableMs: 500,
  source: 'browser_extension',
  observedAtMs: startMs + 750,
});
assert.equal(observed.speaker.signals.length, 0);

observed = observeBrowserMeetingSample(observed.state, meetSnapshot(startMs + 1_500, { id: 'grace', name: 'Grace', speaking: true }), {
  switchStableMs: 400,
  source: 'browser_extension',
  observedAtMs: startMs + 1_500,
});
assert.equal(observed.speaker.signals.length, 0);
observed = observeBrowserMeetingSample(observed.state, meetSnapshot(startMs + 1_950, { id: 'grace', name: 'Grace', speaking: true }), {
  switchStableMs: 400,
  source: 'browser_extension',
  observedAtMs: startMs + 1_950,
});
assert.deepEqual(observed.speaker.signals.map((item) => item.type), ['speaker_ended', 'speaker_started']);
assert.equal(observed.speaker.signals[0].speaker_id, 'ada');
assert.equal(observed.speaker.signals[1].speaker_id, 'grace');

const observer = createBrowserMeetingObserver({
  source: 'browser_extension',
  speakerOptions: {
    minStableMs: 0,
    endIdleMs: 500,
  },
});
const start = observer.observe(meetSnapshot(startMs + 3_000, { id: 'lin', name: 'Lin', speaking: true }), {
  observedAtMs: startMs + 3_000,
});
assert.deepEqual(start.signals.map((item) => item.type), ['meeting_started', 'speaker_started']);
assert.equal(observer.getState().sessionState.activeMeeting.meeting_id, 'abc-defg-hij');

let ending = observer.observe({
  windows: [{
    id: 'win-1',
    focused: true,
    tabs: [{ id: 'tab-mail', active: true, url: 'https://mail.google.com', title: 'Inbox' }],
  }],
  observedAtMs: startMs + 4_000,
}, {
  observedAtMs: startMs + 4_000,
});
assert.deepEqual(ending.signals.map((item) => item.type), ['meeting_ended']);

ending = observer.observe({
  windows: [{
    id: 'win-1',
    focused: true,
    tabs: [{ id: 'tab-mail', active: true, url: 'https://mail.google.com', title: 'Inbox' }],
  }],
  observedAtMs: startMs + 4_600,
}, {
  observedAtMs: startMs + 4_600,
});
assert.deepEqual(ending.signals.map((item) => item.type), ['speaker_ended']);

const calls = [];
const timeline = createBrowserMeetingTimelineObserver({
  async startMeeting(payload) {
    calls.push({ method: 'startMeeting', payload });
    return { ok: true, payload };
  },
  async insertMark(input, options) {
    calls.push({ method: 'insertMark', input, options });
    return { ok: true, input };
  },
}, {
  source: 'browser_extension',
  speakerOptions: { minStableMs: 0 },
  applyOptions: { speakerAsAnnotation: true },
});

const applied = await timeline.observe({
  tabs: [{
    id: 'zoom-web',
    active: true,
    audible: true,
    url: 'https://us06web.zoom.us/wc/987654321/start',
    title: 'Zoom Meeting',
    page: {
      inMeeting: true,
      activeSpeaker: { name: 'Ada', speaking: true },
    },
  }],
  observedAtMs: startMs + 5_000,
}, {
  observedAtMs: startMs + 5_000,
});
assert.deepEqual(applied.signals.map((item) => item.type), ['meeting_started', 'speaker_started']);
assert.deepEqual(applied.results.map((item) => item.action), ['startMeeting', 'insertSpeakerMark']);
assert.equal(calls[0].payload.platform, 'zoom');
assert.equal(calls[0].payload.meeting_id, '987654321');
assert.equal(calls[1].input.kind, 'speaker_started');
assert.equal(calls[1].input.captured_at_ms, startMs + 5_000);

console.log('ok browser meeting adapter');
