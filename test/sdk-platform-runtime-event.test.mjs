import assert from 'node:assert/strict';

import {
  MEETING_PLATFORM_RUNTIME_EVENT_ENDPOINT,
  MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA,
  assertMeetingPlatformRuntimeEvent,
  buildMeetingPlatformAnnotationRuntimeEvent,
  buildMeetingPlatformAdapterRouteRuntimeEvent,
  buildMeetingPlatformAdapterRoutesRuntimeEvent,
  buildMeetingPlatformCandidateObservationRuntimeEvent,
  buildMeetingPlatformObserveRuntimeEvent,
  buildMeetingPlatformProviderRuntimeEvent,
  buildMeetingPlatformRunHandoffReadinessRuntimeEvent,
  buildMeetingPlatformRunManifestRuntimeEvent,
  buildMeetingPlatformRuntimeEventPlan,
  buildMeetingPlatformRuntimeEventPlanMatrix,
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
assert.equal(normalizeMeetingPlatformRuntimeEventAction('observe-candidates'), 'observe_platform_candidates');
assert.equal(normalizeMeetingPlatformRuntimeEventAction('integration-runtime-run-manifest'), 'run_manifest');
assert.equal(normalizeMeetingPlatformRuntimeEventAction('handoff-readiness-run'), 'run_handoff_readiness');
assert.equal(normalizeMeetingPlatformRuntimeEventAction('adapter-route-matrix'), 'adapter_routes');
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

const candidateObservationEvent = buildMeetingPlatformCandidateObservationRuntimeEvent({
  windows: [{
    id: 'browser-window-1',
    tabs: [{
      id: 'meet-tab',
      active: true,
      url: 'https://meet.google.com/abc-defg-hij',
      title: 'Design review - Google Meet',
    }],
  }],
}, {
  now: () => now + 2,
  source: 'native_host',
});
assert.equal(candidateObservationEvent.action, 'observe_platform_candidates');
assert.equal(candidateObservationEvent.platform, undefined);
assert.equal(candidateObservationEvent.windows[0].tabs[0].active, true);
assert.equal(candidateObservationEvent.environment.windows[0].id, 'browser-window-1');

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

const runManifestEvent = buildMeetingPlatformRunManifestRuntimeEvent({
  platforms: ['google-meet'],
  requireHandoffReady: true,
  target: 'production',
}, {
  now: () => now + 4,
});
assert.equal(runManifestEvent.action, 'run_manifest');
assert.equal(runManifestEvent.requireHandoffReady, true);
assert.deepEqual(runManifestEvent.platforms, ['google-meet']);

const runHandoffReadinessEvent = buildMeetingPlatformRunHandoffReadinessRuntimeEvent({
  platforms: ['zoom'],
  target: 'pilot',
}, {
  now: () => now + 5,
});
assert.equal(runHandoffReadinessEvent.action, 'run_handoff_readiness');
assert.deepEqual(runHandoffReadinessEvent.platforms, ['zoom']);

const adapterRouteEvent = buildMeetingPlatformAdapterRouteRuntimeEvent('google-meet', {
  purpose: 'inspect_single_route',
}, {
  now: () => now + 6,
});
assert.equal(adapterRouteEvent.action, 'adapter_route');
assert.equal(adapterRouteEvent.platform, 'google_meet');

const adapterRoutesEvent = buildMeetingPlatformAdapterRoutesRuntimeEvent({
  platforms: ['google-meet', 'zoom'],
}, {
  now: () => now + 7,
});
assert.equal(adapterRoutesEvent.action, 'adapter_routes');
assert.equal(adapterRoutesEvent.platform, undefined);
assert.deepEqual(adapterRoutesEvent.platforms, ['google-meet', 'zoom']);

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

const candidateObservationResult = await runtime.handleEvent(candidateObservationEvent, undefined, {
  observedAtMs: now - 1_000,
});
assert.equal(candidateObservationResult.action, 'observe_platform_candidates');
assert.equal(candidateObservationResult.platform, 'google_meet');
assert.equal(candidateObservationResult.signals[0].type, 'meeting_started');
assert.equal(calls.some((call) => call.method === 'startMeeting' && call.input.meeting_id === 'abc-defg-hij'), true);

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

const runtimeRunManifestResult = await runtime.handleEvent({
  action: 'run_manifest',
  platforms: ['google-meet'],
  requireHandoffReady: true,
  target: 'production',
});
assert.equal(runtimeRunManifestResult.require_handoff_ready, true);
assert.equal(runtimeRunManifestResult.platform_count, 1);
assert.equal(runtimeRunManifestResult.host_integration_ready, false);
assert.equal(runtimeRunManifestResult.issues.some((issue) => issue.code === 'runtime_host_replay_not_ready'), true);

const runtimeRunHandoffReadinessResult = await runtime.handleEvent({
  action: 'run_handoff_readiness',
  platforms: ['zoom'],
  target: 'production',
});
assert.equal(runtimeRunHandoffReadinessResult.platform_count, 1);
assert.equal(runtimeRunHandoffReadinessResult.runtime_host_replay_ready_count, 0);
assert.equal((await runtime.handleEvent(adapterRouteEvent)).routes[0].route, 'local_observer_axis');
assert.equal((await runtime.handleEvent(adapterRoutesEvent)).platform_count, 2);

const googlePlan = buildMeetingPlatformRuntimeEventPlan('google-meet', {
  baseUrl,
  now: () => now + 6,
});
assert.equal(googlePlan.schema, 'meeting_platform_runtime_event_plan');
assert.equal(googlePlan.platform, 'google_meet');
assert.equal(googlePlan.endpoint, `${baseUrl}${MEETING_PLATFORM_RUNTIME_EVENT_ENDPOINT}`);
assert.equal(googlePlan.client_factory, 'createMeetingPlatformRuntimeEventClient');
assert.equal(googlePlan.realtime_contract.primary_clock_field, 'captured_at_ms');
assert.equal(googlePlan.realtime_contract.provider_events_required_for_realtime, false);
assert.equal(googlePlan.realtime_contract.transcript_required_for_realtime, false);
assert.equal(googlePlan.provider_start_event_example, 'google.workspace.meet.conference.v2.started');
assert.equal(googlePlan.provider_end_event_example, 'google.workspace.meet.conference.v2.ended');
assert.equal(googlePlan.actions.find((row) => row.action === 'insert_annotation').client_method, 'insertAnnotation');
assert.equal(googlePlan.actions.find((row) => row.action === 'observe_platform_candidates').client_method, 'observePlatformCandidates');
assert.equal(googlePlan.actions.find((row) => row.action === 'adapter_route').client_method, 'adapterRoute');
assert.equal(googlePlan.actions.find((row) => row.action === 'adapter_routes').client_method, 'adapterRoutes');
assert.equal(googlePlan.actions.find((row) => row.action === 'run_manifest').client_method, 'runManifest');
assert.equal(googlePlan.actions.find((row) => row.action === 'run_handoff_readiness').client_method, 'runHandoffReadiness');
assert.equal(googlePlan.actions.find((row) => row.action === 'provider_event').realtime_role, 'reconcile_and_backfill_only');
assert.equal(googlePlan.examples.insert_annotation.action, 'insert_annotation');
assert.equal(googlePlan.examples.observe_platform_candidates.action, 'observe_platform_candidates');
assert.equal(googlePlan.examples.adapter_route.action, 'adapter_route');
assert.equal(googlePlan.examples.adapter_routes.action, 'adapter_routes');
assert.equal(googlePlan.examples.run_manifest.action, 'run_manifest');
assert.equal(googlePlan.examples.run_handoff_readiness.action, 'run_handoff_readiness');
assert.equal(googlePlan.examples.insert_annotation.annotation.captured_at_ms, now + 15_000);
assert.equal(googlePlan.examples.observe_meeting_app.snapshot.url, 'https://meet.google.com/abc-defg-hij');

const runtimePlanMatrix = buildMeetingPlatformRuntimeEventPlanMatrix({
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom'],
  now: () => now + 7,
});
assert.equal(runtimePlanMatrix.schema, 'meeting_platform_runtime_event_plan_matrix');
assert.equal(runtimePlanMatrix.platform_count, 3);
assert.equal(runtimePlanMatrix.realtime_provider_dependency_count, 0);
assert.equal(runtimePlanMatrix.transcript_realtime_dependency_count, 0);
assert.equal(runtimePlanMatrix.rows.some((row) => row.platform === 'microsoft_teams' && row.action === 'speaker_track'), true);
assert.equal(runtimePlanMatrix.rows.some((row) => row.platform === 'zoom' && row.action === 'observe_platform_candidates'), true);
assert.equal(runtimePlanMatrix.rows.some((row) => row.platform === 'google_meet' && row.action === 'adapter_routes'), true);
assert.equal(runtimePlanMatrix.rows.some((row) => row.platform === 'google_meet' && row.action === 'run_manifest'), true);
assert.equal(runtimePlanMatrix.rows.some((row) => row.platform === 'zoom' && row.action === 'run_handoff_readiness'), true);
assert.equal(runtimePlanMatrix.next_actions.includes('include_captured_at_ms_on_every_annotation'), true);

let capturedRequest = null;
const runtimeEventClient = createMeetingPlatformRuntimeEventClient({
  baseUrl,
  now: () => now + 8,
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
await runtimeEventClient.observePlatformCandidates({
  windows: [{
    tabs: [{ active: true, url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet' }],
  }],
});
assert.equal(capturedRequest.body.action, 'observe_platform_candidates');
assert.equal(capturedRequest.body.platform, undefined);
assert.equal(capturedRequest.body.windows[0].tabs[0].url, 'https://meet.google.com/abc-defg-hij');
assert.equal((await runtimeEventClient.manifest()).accepted.action, 'manifest');
await runtimeEventClient.adapterRoute('google-meet');
assert.equal(capturedRequest.body.action, 'adapter_route');
assert.equal(capturedRequest.body.platform, 'google_meet');
await runtimeEventClient.adapterRoutes({
  platforms: ['google-meet', 'zoom'],
});
assert.equal(capturedRequest.body.action, 'adapter_routes');
assert.deepEqual(capturedRequest.body.platforms, ['google-meet', 'zoom']);
await runtimeEventClient.runManifest({
  platforms: ['google-meet'],
  requireHandoffReady: true,
});
assert.equal(capturedRequest.body.action, 'run_manifest');
assert.equal(capturedRequest.body.requireHandoffReady, true);
await runtimeEventClient.runHandoffReadiness({
  platforms: ['zoom'],
  target: 'production',
});
assert.equal(capturedRequest.body.action, 'run_handoff_readiness');
assert.deepEqual(capturedRequest.body.platforms, ['zoom']);

console.log('ok meeting platform runtime event envelope');
