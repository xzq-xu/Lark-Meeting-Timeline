import assert from 'node:assert/strict';
import {
  MeetingTimelineApiError,
  MeetingTimelineSdkError,
  createMeetingTimelineAnnotationProducer,
} from '../packages/meeting-timeline-sdk/index.mjs';
import {
  MeetingTimelineAnnotationProducer,
} from '../packages/meeting-timeline-sdk/producer.mjs';

const startMs = 1_783_900_000_000;
let now = startMs;

const captureProducer = createMeetingTimelineAnnotationProducer({
  client: { insertMark: async () => ({ accepted: true }) },
  producerId: 'reference-ui',
  clock: () => now,
  idFactory: ({ producerId, sequence }) => `${producerId}:${sequence}`,
});

assert.equal(captureProducer instanceof MeetingTimelineAnnotationProducer, true);
const captured = captureProducer.capture({ label: 'why?', intent: 'question' });
assert.equal(captured.id, 'reference-ui:1');
assert.equal(captured.source, 'sdk_annotation_producer');
assert.equal(captured.captured_at_ms, startMs);
assert.equal(captured.realtime, true);
assert.equal(captureProducer.pendingCount, 0);

const inlineCaptured = captureProducer.capture({
  label: 'inline',
  meetingSession: {
    platform: 'google_meet',
    meetingId: 'inline-meeting',
    startTimeMs: startMs,
    meetingUrl: 'https://meet.google.com/abc-defg-hij',
  },
});
assert.equal(inlineCaptured.meeting_session.meeting_id, 'inline-meeting');
assert.equal(inlineCaptured.meeting_session.start_time_ms, startMs);
assert.equal(inlineCaptured.meeting_session.meeting_url, 'https://meet.google.com/abc-defg-hij');

const retryCalls = [];
const retryDelays = [];
const retryProducer = createMeetingTimelineAnnotationProducer({
  client: {
    async insertMark(annotation, options) {
      retryCalls.push({ annotation, options });
      if (retryCalls.length === 1) {
        throw new MeetingTimelineApiError('temporary', { status: 503 });
      }
      now += 12;
      return {
        accepted: true,
        annotation_evidence: { row: { visible_at_ms: now } },
      };
    },
  },
  clock: () => now,
  retryDelayMs: 25,
  sleep: async (delayMs) => { retryDelays.push(delayMs); },
});

const retried = await retryProducer.publish({
  id: 'stable-on-retry',
  capturedAtMs: startMs,
  label: 'what?',
});
assert.equal(retried.accepted, true);
assert.equal(retried.attempts, 2);
assert.equal(retried.delivery_latency_ms, 12);
assert.equal(retryCalls.length, 2);
assert.equal(retryCalls[0].annotation.id, retryCalls[1].annotation.id);
assert.deepEqual(retryCalls.map((call) => call.options), [{ raw: true }, { raw: true }]);
assert.deepEqual(retryDelays, [25]);
assert.equal(retryProducer.pendingCount, 0);

let rejectDelivery = true;
const retainedProducer = createMeetingTimelineAnnotationProducer({
  client: {
    async insertMark(annotation) {
      if (rejectDelivery) throw new MeetingTimelineApiError('bad request', { status: 400 });
      return { accepted: true, annotation };
    },
  },
  clock: () => now,
  maxAttempts: 5,
});

await assert.rejects(
  retainedProducer.publish({ id: 'retained', capturedAtMs: startMs + 1, label: 'retry later' }),
  (error) => error instanceof MeetingTimelineSdkError
    && error.details.annotation_id === 'retained'
    && error.details.attempts === 1,
);
assert.equal(retainedProducer.pendingCount, 1);
assert.equal(retainedProducer.pendingAnnotations()[0].id, 'retained');

rejectDelivery = false;
const flushed = await retainedProducer.flush({ maxAttempts: 1 });
assert.equal(flushed.accepted, true);
assert.equal(flushed.delivered_count, 1);
assert.equal(flushed.deliveries[0].attempts, 2);
assert.equal(retainedProducer.pendingCount, 0);

const batchIds = [];
const batchProducer = createMeetingTimelineAnnotationProducer({
  client: {
    async insertMark(annotation) {
      batchIds.push(annotation.id);
      return { accepted: true };
    },
  },
  clock: () => now,
  idFactory: ({ sequence }) => `batch-${sequence}`,
});
const batch = await batchProducer.publishBatch([
  { capturedAtMs: startMs + 2, label: 'one' },
  { capturedAtMs: startMs + 3, label: 'two' },
]);
assert.equal(batch.accepted, true);
assert.equal(batch.requested_count, 2);
assert.deepEqual(batch.requested_ids, ['batch-1', 'batch-2']);
assert.deepEqual(batchIds, ['batch-1', 'batch-2']);

let fetchThisValue = Symbol('not-called');
const httpProducer = createMeetingTimelineAnnotationProducer({
  baseUrl: 'https://timeline.example',
  fetch: async function fetchWithoutReceiver(url, init) {
    fetchThisValue = this;
    assert.equal(url, 'https://timeline.example/api/annotations');
    assert.equal(init.method, 'POST');
    return {
      ok: true,
      status: 200,
      async text() {
        return JSON.stringify({ accepted: true });
      },
    };
  },
  clock: () => now,
});
await httpProducer.publish({ id: 'browser-fetch-binding', label: 'browser-safe fetch' });
assert.equal(fetchThisValue, undefined);

assert.throws(
  () => createMeetingTimelineAnnotationProducer({
    client: { insertMark: async () => ({}) },
    clock: () => startMs,
    idFactory: () => '',
  }).capture({ label: 'missing id' }),
  MeetingTimelineSdkError,
);

console.log('ok sdk annotation producer');
