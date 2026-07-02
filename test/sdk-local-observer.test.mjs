import assert from 'node:assert/strict';

import {
  createLocalMeetingObserver,
  observeMeetingSnapshot,
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

console.log('ok local meeting observer state machine');
