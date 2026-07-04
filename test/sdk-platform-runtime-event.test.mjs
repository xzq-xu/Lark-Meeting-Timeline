import assert from 'node:assert/strict';

import {
  MEETING_PLATFORM_RUNTIME_EVENT_ENDPOINT,
  MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA,
  assertMeetingPlatformRuntimeEvent,
  buildMeetingPlatformAnnotationRuntimeEvent,
  buildMeetingPlatformObserveRuntimeEvent,
  buildMeetingPlatformProviderRuntimeEvent,
  createMeetingPlatformRuntimeEventClient,
  meetingPlatformRuntimeEventEndpoint,
  normalizeMeetingPlatformRuntimeEventAction,
} from '../packages/meeting-timeline-sdk/adapters/platform-runtime-event.mjs';
import {
  createMeetingPlatformIntegrationRuntime,
} from '../packages/meeting-timeline-sdk/adapters/platform-integration-runtime.mjs';
import {
  buildPlatformFixtureEvent,
} from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';

const baseUrl = 'https://timeline.example.com';
const now = 1_782_614_400_000;

assert.equal(normalizeMeetingPlatformRuntimeEventAction('insert-mark'), 'insert_annotation');
assert.equal(meetingPlatformRuntimeEventEndpoint({ baseUrl }), `${baseUrl}${MEETING_PLATFORM_RUNTIME_EVENT_ENDPOINT}`);

const observeEvent = buildMeetingPlatformObserveRuntimeEvent('google-meet', {
  url: 'https://meet.google.com/abc-defg-hij',
  observed_at_ms: now,
}, {
  now: () => now + 1,
  source: 'content_script',
});
assert.equal(observeEvent.schema, MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA);
assert.equal(observeEvent.action, 'observe_meeting_app');
assert.equal(observeEvent.platform, 'google_meet');
assert.equal(observeEvent.snapshot.url, 'https://meet.google.com/abc-defg-hij');
assert.equal(observeEvent.sent_at_ms, now + 1);

const annotationEvent = buildMeetingPlatformAnnotationRuntimeEvent('zoom', {
  annotation: {
    id: 'runtime-event-note-1',
    label: 'why?',
    captured_at_ms: now + 2_000,
  },
  current_meeting: {
    platform: 'zoom',
    meeting_id: '987654321',
    start_time_ms: now,
  },
}, {
  event_id: 'evt-note-1',
  now: () => now + 3,
});
assert.equal(annotationEvent.action, 'insert_annotation');
assert.equal(annotationEvent.event_id, 'evt-note-1');
assert.equal(annotationEvent.annotation.label, 'why?');
assert.equal(annotationEvent.current_meeting.meeting_id, '987654321');
assert.equal(assertMeetingPlatformRuntimeEvent(annotationEvent).platform, 'zoom');
assert.throws(
  () => buildMeetingPlatformAnnotationRuntimeEvent('', {
    annotation: { label: 'why?', captured_at_ms: now },
  }),
  /requires platform/,
);

const calls = [];
const client = {
  async startMeeting(input) {
    calls.push({ method: 'startMeeting', input });
    return { ok: true, method: 'startMeeting', input };
  },
  async endMeeting(input) {
    calls.push({ method: 'endMeeting', input });
    return { ok: true, method: 'endMeeting', input };
  },
  async insertMark(input) {
    calls.push({ method: 'insertMark', input });
    return { ok: true, method: 'insertMark', input };
  },
  async insertMarks(input) {
    calls.push({ method: 'insertMarks', input });
    return { ok: true, method: 'insertMarks', input };
  },
};

const runtime = createMeetingPlatformIntegrationRuntime(client, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
  verify: false,
});

const payloadOnlyInsert = await runtime.handleEvent({
  action: 'insert_annotation',
  platform: 'google-meet',
  payload: {
    annotation: {
      id: 'payload-only-note-1',
      label: 'follow up',
      captured_at_ms: now + 4_000,
    },
    current_meeting: {
      platform: 'google_meet',
      meeting_id: 'abc-defg-hij',
      meeting_url: 'https://meet.google.com/abc-defg-hij',
      start_time_ms: now,
    },
  },
});
assert.equal(payloadOnlyInsert.action, 'insert_annotation');
assert.equal(payloadOnlyInsert.result.status, 'ready_to_insert');
assert.equal(calls.some((call) => call.method === 'insertMark' && call.input.id === 'payload-only-note-1'), true);

const providerEvent = buildMeetingPlatformProviderRuntimeEvent(
  'google-meet',
  buildPlatformFixtureEvent('google-meet', 'meeting_end'),
  { now: () => now + 5 },
);
const providerResult = await runtime.handleEvent(providerEvent);
assert.equal(providerResult.action, 'ingest_provider');
assert.equal(providerResult.live_evidence.provider_record_count, 1);

let capturedRequest = null;
const runtimeEventClient = createMeetingPlatformRuntimeEventClient({
  baseUrl,
  now: () => now + 6,
  fetch: async (url, init) => {
    capturedRequest = {
      url,
      method: init.method,
      headers: init.headers,
      body: JSON.parse(init.body),
    };
    return new Response(JSON.stringify({ ok: true, accepted: capturedRequest.body }), {
      status: 201,
      headers: { 'content-type': 'application/json' },
    });
  },
});

const sendResult = await runtimeEventClient.insertAnnotation('zoom', {
  annotation: {
    id: 'client-note-1',
    label: 'next',
    captured_at_ms: now + 7_000,
  },
  current_meeting: {
    platform: 'zoom',
    meeting_id: '987654321',
    start_time_ms: now,
  },
});
assert.equal(capturedRequest.url, `${baseUrl}${MEETING_PLATFORM_RUNTIME_EVENT_ENDPOINT}`);
assert.equal(capturedRequest.method, 'POST');
assert.equal(capturedRequest.body.schema, MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA);
assert.equal(capturedRequest.body.action, 'insert_annotation');
assert.equal(capturedRequest.body.platform, 'zoom');
assert.equal(sendResult.accepted.annotation.id, 'client-note-1');
assert.equal((await runtimeEventClient.manifest()).accepted.action, 'manifest');

console.log('ok meeting platform runtime event envelope');
