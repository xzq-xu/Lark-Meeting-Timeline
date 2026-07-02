import assert from 'node:assert/strict';

import {
  createLocalMeetingObserver,
  createLocalMeetingTimelineObserver,
  observeMeetingSnapshot,
  selectMeetingSnapshot,
} from '../packages/meeting-timeline-sdk/adapters/local-observer.mjs';

const startMs = 1_782_442_800_000;

let result = observeMeetingSnapshot(null, {
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Google Meet',
  observed_at_ms: startMs,
});
assert.equal(result.signals.length, 1);
assert.equal(result.signals[0].type, 'meeting_started');
assert.equal(result.signals[0].meeting.platform, 'google_meet');
assert.equal(result.signals[0].meeting.meeting_id, 'abc-defg-hij');
assert.equal(result.signals[0].occurred_at_ms, startMs);
assert.equal(result.state.activeMeeting.platform, 'google_meet');
assert.equal(result.state.activeSinceMs, startMs);

result = observeMeetingSnapshot(result.state, {
  url: 'https://meet.google.com/abc-defg-hij',
  observed_at_ms: startMs + 5_000,
});
assert.equal(result.signals.length, 0);
assert.equal(result.state.activeMeeting.meeting_id, 'abc-defg-hij');
assert.equal(result.state.activeSinceMs, startMs);

result = observeMeetingSnapshot(result.state, {
  url: 'https://teams.microsoft.com/l/meetup-join/19%3Ameeting_XYZ%40thread.v2/0',
  observed_at_ms: startMs + 10_000,
});
assert.deepEqual(result.signals.map((item) => item.type), ['meeting_ended', 'meeting_started']);
assert.equal(result.signals[0].meeting.platform, 'google_meet');
assert.equal(result.signals[1].meeting.platform, 'microsoft_teams');
assert.equal(result.signals[1].meeting.meeting_id, '19:meeting_XYZ@thread.v2');
assert.equal(result.state.activeMeeting.platform, 'microsoft_teams');

result = observeMeetingSnapshot(result.state, {
  visible: false,
  observed_at_ms: startMs + 20_000,
});
assert.equal(result.signals.length, 1);
assert.equal(result.signals[0].type, 'meeting_ended');
assert.equal(result.signals[0].meeting.platform, 'microsoft_teams');
assert.equal(result.state.activeMeeting, null);

const observer = createLocalMeetingObserver({ source: 'desktop_observer' });
const first = observer.observe({
  window: {
    url: 'https://us06web.zoom.us/j/987654321',
    title: 'Zoom Meeting',
  },
  observedAtMs: startMs + 30_000,
});
assert.equal(first.signals[0].source, 'desktop_observer');
assert.equal(first.signals[0].meeting.platform, 'zoom');
assert.equal(first.signals[0].meeting.meeting_id, '987654321');

const second = observer.observe({
  window: {
    url: 'https://us06web.zoom.us/j/987654321',
  },
  observedAtMs: startMs + 31_000,
});
assert.equal(second.signals.length, 0);
assert.equal(observer.getState().activeMeeting.platform, 'zoom');

const ended = observer.observe({
  url: 'https://example.com/not-a-meeting',
  observedAtMs: startMs + 40_000,
});
assert.equal(ended.signals[0].type, 'meeting_ended');
assert.equal(observer.getState().activeMeeting, null);

const windowSelection = selectMeetingSnapshot({
  windows: [{
    id: 'win-1',
    focused: true,
    lastFocusedAtMs: startMs + 45_000,
    tabs: [
      {
        id: 'tab-docs',
        active: true,
        url: 'https://docs.example.com/notes',
        title: 'Notes',
      },
      {
        id: 'tab-meet',
        active: false,
        url: 'https://meet.google.com/qwe-rtyu-iop',
        title: 'Project review - Google Meet',
      },
    ],
  }],
}, {
  recencyBaseMs: startMs,
});
assert.equal(windowSelection.detectedMeeting.platform, 'google_meet');
assert.equal(windowSelection.detectedMeeting.meeting_id, 'qwe-rtyu-iop');
assert.equal(windowSelection.candidates.length, 1);
assert.equal(windowSelection.selectedSnapshot.active, undefined);

const candidateObserver = createLocalMeetingObserver({ source: 'browser_extension' });
const candidateStart = candidateObserver.observeCandidates({
  windows: [{
    focused: true,
    tabs: [
      { active: true, url: 'https://mail.example.com', title: 'Mail' },
      { active: false, url: 'https://meet.google.com/qwe-rtyu-iop', title: 'Project review - Google Meet' },
    ],
  }],
}, {
  observedAtMs: startMs + 46_000,
});
assert.equal(candidateStart.signals.length, 1);
assert.equal(candidateStart.signals[0].type, 'meeting_started');
assert.equal(candidateStart.signals[0].meeting.platform, 'google_meet');
assert.equal(candidateStart.selection.candidates.length, 1);

const candidateEnd = candidateObserver.observeCandidates({
  tabs: [
    { active: true, url: 'https://mail.example.com', title: 'Mail' },
    { active: false, url: 'https://docs.example.com/notes', title: 'Notes' },
  ],
}, {
  observedAtMs: startMs + 47_000,
});
assert.equal(candidateEnd.signals.length, 1);
assert.equal(candidateEnd.signals[0].type, 'meeting_ended');
assert.equal(candidateObserver.getState().activeMeeting, null);

const clientCalls = [];
const timelineClient = {
  async startMeeting(payload) {
    clientCalls.push({ method: 'startMeeting', payload });
    return { ok: true, payload };
  },
  async endMeeting(payload) {
    clientCalls.push({ method: 'endMeeting', payload });
    return { ok: true, payload };
  },
};

const timelineObserver = createLocalMeetingTimelineObserver(timelineClient, {
  source: 'desktop_observer',
});

const timelineStart = await timelineObserver.observe({
  tab: {
    url: 'https://meet.google.com/xyz-abcd-uvw',
    title: 'Weekly sync - Google Meet',
  },
  observedAtMs: startMs + 50_000,
});
assert.equal(timelineStart.signals.length, 1);
assert.equal(timelineStart.results.length, 1);
assert.equal(timelineStart.results[0].action, 'startMeeting');
assert.equal(clientCalls[0].method, 'startMeeting');
assert.equal(clientCalls[0].payload.platform, 'google_meet');
assert.equal(clientCalls[0].payload.meeting_id, 'xyz-abcd-uvw');
assert.equal(clientCalls[0].payload.detector_source, 'google_meet_desktop_observer');

const timelineRepeat = await timelineObserver.observe({
  url: 'https://meet.google.com/xyz-abcd-uvw',
  observedAtMs: startMs + 51_000,
});
assert.equal(timelineRepeat.signals.length, 0);
assert.equal(timelineRepeat.results.length, 0);
assert.equal(clientCalls.length, 1);

const timelineEnd = await timelineObserver.observe({
  active: false,
  observedAtMs: startMs + 60_000,
});
assert.equal(timelineEnd.signals.length, 1);
assert.equal(timelineEnd.results.length, 1);
assert.equal(timelineEnd.results[0].action, 'endMeeting');
assert.equal(clientCalls[1].method, 'endMeeting');
assert.equal(clientCalls[1].payload.meeting_id, 'xyz-abcd-uvw');
assert.equal(clientCalls[1].payload.end_time_ms, startMs + 60_000);

const timelineCandidateStart = await timelineObserver.observeCandidates({
  tabs: [
    {
      active: true,
      url: 'https://teams.microsoft.com/l/meetup-join/19%3Ameeting_CANDIDATE%40thread.v2/0',
      title: 'Candidate Teams meeting',
    },
  ],
}, {
  observedAtMs: startMs + 70_000,
});
assert.equal(timelineCandidateStart.signals[0].meeting.platform, 'microsoft_teams');
assert.equal(timelineCandidateStart.results[0].action, 'startMeeting');
assert.equal(clientCalls[2].method, 'startMeeting');

console.log('ok local meeting observer state machine');
