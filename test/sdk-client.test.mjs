import assert from 'node:assert/strict';
import {
  MeetingTimelineApiError,
  MeetingTimelineSdkError,
  buildMeetingEndPayload,
  buildMeetingStartPayload,
  buildTimelineMark,
  createMeetingTimelineClient,
  normalizeAbsoluteMs,
} from '../packages/meeting-timeline-sdk/index.mjs';

const startMs = 1_782_442_800_000;

assert.equal(normalizeAbsoluteMs(1_782_442_800), startMs);
assert.equal(normalizeAbsoluteMs(startMs), startMs);
assert.equal(normalizeAbsoluteMs(new Date(startMs)), startMs);

const startPayload = buildMeetingStartPayload({
  meetingId: 'sdk-meeting-001',
  meetingNo: '123456789',
  title: 'SDK meeting',
  meetingUrl: 'https://vc.feishu.cn/j/sdk-meeting-001',
  startTimeMs: startMs,
  detectorSource: 'sdk-test',
});
assert.deepEqual(startPayload, {
  platform: 'lark',
  meeting_id: 'sdk-meeting-001',
  external_meeting_id: '123456789',
  meeting_url: 'https://vc.feishu.cn/j/sdk-meeting-001',
  title: 'SDK meeting',
  start_time_ms: startMs,
  detector_source: 'sdk-test',
});

const markPayload = buildTimelineMark({
  id: 'sdk-mark-001',
  capturedAtMs: startMs + 12_000,
  label: 'why?',
  textCandidates: ['why?', 'why'],
  intent: 'question',
  strokes: [[{ x: 0.1, y: 0.2, t: startMs + 12_000 }]],
}, {
  source: 'hanwang_epaper',
  deviceId: 'hanwang-001',
});
assert.equal(markPayload.source, 'hanwang_epaper');
assert.equal(markPayload.device_id, 'hanwang-001');
assert.equal(markPayload.captured_at_ms, startMs + 12_000);
assert.deepEqual(markPayload.text_candidates, ['why?', 'why']);

assert.throws(() => buildTimelineMark({ id: 'missing-time' }), MeetingTimelineSdkError);
assert.equal(buildTimelineMark({ id: 'server-time-ok' }, { requireCapturedAt: false }).captured_at_ms, undefined);

const inlinePayload = buildTimelineMark({
  id: 'sdk-inline-001',
  capturedAtMs: startMs + 30_000,
  label: 'follow up',
  meetingSession: {
    meetingId: 'sdk-inline-meeting',
    title: 'Inline meeting',
    startTimeMs: startMs,
  },
});
assert.equal(inlinePayload.meeting_session.meeting_id, 'sdk-inline-meeting');
assert.equal(inlinePayload.meeting_session.start_time_ms, startMs);

const endPayload = buildMeetingEndPayload({
  endTimeMs: startMs + 60_000,
  detectorSource: 'sdk-test',
});
assert.deepEqual(endPayload, {
  end_time_ms: startMs + 60_000,
  detector_source: 'sdk-test',
});

const calls = [];
const client = createMeetingTimelineClient({
  baseUrl: 'http://localhost:8787/',
  source: 'hanwang_epaper',
  deviceId: 'hanwang-001',
  detectorSource: 'sdk-test',
  fetch: async (url, init = {}) => {
    const body = init.body ? JSON.parse(init.body) : null;
    calls.push({ url, init, body });
    return new Response(JSON.stringify({ ok: true, url, body }), {
      status: 200,
      headers: { 'content-type': 'application/json' },
    });
  },
});

await client.startMeeting({
  meetingId: 'sdk-meeting-001',
  title: 'SDK meeting',
  startTimeMs: startMs,
});
assert.equal(calls.at(-1).url, 'http://localhost:8787/api/meeting-session/start');
assert.equal(calls.at(-1).body.detector_source, 'sdk-test');

await client.insertMark({
  id: 'sdk-mark-002',
  capturedAtMs: startMs + 20_000,
  textCandidates: ['why?'],
});
assert.equal(calls.at(-1).url, 'http://localhost:8787/api/annotations');
assert.equal(calls.at(-1).body.source, 'hanwang_epaper');
assert.equal(calls.at(-1).body.device_id, 'hanwang-001');

await client.insertMarks([
  { id: 'sdk-mark-003', capturedAtMs: startMs + 21_000, label: 'A' },
  { id: 'sdk-mark-004', capturedAtMs: startMs + 22_000, label: 'B' },
]);
assert.equal(calls.at(-1).url, 'http://localhost:8787/api/annotations/batch');
assert.equal(calls.at(-1).body.annotations.length, 2);

await client.getAnnotationStatus('sdk-mark-002');
assert.equal(calls.at(-1).url, 'http://localhost:8787/api/annotations/status?id=sdk-mark-002');

const failing = createMeetingTimelineClient({
  baseUrl: 'http://localhost:8787',
  fetch: async () => new Response(JSON.stringify({ error: 'bad request' }), {
    status: 400,
    headers: { 'content-type': 'application/json' },
  }),
});
await assert.rejects(
  () => failing.getState(),
  (error) => error instanceof MeetingTimelineApiError && error.status === 400 && error.body.error === 'bad request',
);

console.log('ok meeting timeline SDK client');
