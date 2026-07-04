import assert from 'node:assert/strict';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';
import {
  assertMeetingPlatformLiveAdapterReadiness,
  assertMeetingPlatformLiveAdapterReadinessMatrix,
  buildMeetingPlatformLiveAdapterHandoff,
  buildMeetingPlatformLiveAdapterHandoffBundle,
  buildMeetingPlatformLiveAdapterMatrix,
  buildMeetingPlatformLiveAdapterPlan,
  buildMeetingPlatformLiveAdapterReadiness,
  buildMeetingPlatformLiveAdapterReadinessMatrix,
  createMeetingPlatformLiveAdapter,
  createMeetingPlatformLiveAdapterSuite,
} from '../packages/meeting-timeline-sdk/adapters/platform-live-adapter.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';
const startMs = 1_784_010_000_000;
const durationMs = 120_000;
const googleEnv = {
  GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
  GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL: 'meet-events@example.iam.gserviceaccount.com',
};
const calls = [];
const client = {
  async startMeeting(input) {
    calls.push({ method: 'startMeeting', input });
    return { ok: true, input };
  },
  async endMeeting(input) {
    calls.push({ method: 'endMeeting', input });
    return { ok: true, input };
  },
  async insertMark(input, options) {
    calls.push({ method: 'insertMark', input, options });
    return { ok: true, input };
  },
  async insertMarks(input, options) {
    calls.push({ method: 'insertMarks', input, options });
    return { ok: true, input };
  },
};

const adapter = createMeetingPlatformLiveAdapter('google-meet', client, {
  id: 'google-live-adapter-test',
  baseUrl,
  env: googleEnv,
  createdAtMs: startMs,
  requireMeetingEnd: false,
  speakerOptions: { minStableMs: 0 },
  reconcileOptions: { duplicateWindowMs: 60_000 },
});

assert.equal(adapter.platform, 'google_meet');
assert.equal(adapter.summary().can_insert_realtime_marks, false);

const googlePlan = buildMeetingPlatformLiveAdapterPlan('google-meet', {
  baseUrl,
  env: googleEnv,
});
assert.equal(googlePlan.schema, 'meeting_platform_live_adapter_plan');
assert.equal(googlePlan.platform, 'google_meet');
assert.equal(googlePlan.live_adapter.factory, 'createMeetingPlatformLiveAdapter');
assert.equal(googlePlan.realtime_axis.provider_events_block_realtime, false);
assert.equal(googlePlan.live_adapter.realtime_methods.includes('insertAnnotation'), true);

const blockedReadiness = buildMeetingPlatformLiveAdapterReadiness('google-meet', {
  baseUrl,
  env: googleEnv,
});
assert.equal(blockedReadiness.schema, 'meeting_platform_live_adapter_readiness');
assert.equal(blockedReadiness.status, 'blocked');
assert.equal(blockedReadiness.candidate_observation_ready, true);
assert.equal(blockedReadiness.candidate_observation.message_type, 'meeting_timeline.observe_candidates');
assert.equal(blockedReadiness.candidate_observation.runtime_event_action, 'observe_platform_candidates');
assert.equal(blockedReadiness.blocking_checks.some((item) => item.code === 'pilot_realtime_axis_ready'), true);
assert.equal(blockedReadiness.warnings.some((item) => item.code === 'evidence_package_available'), true);
assert.throws(
  () => assertMeetingPlatformLiveAdapterReadiness('google-meet', {
    baseUrl,
    env: googleEnv,
  }),
  (error) => error.name === 'MeetingTimelineSdkError'
    && error.details.platform === 'google_meet'
    && error.details.blocking_checks.some((item) => item.code === 'pilot_realtime_axis_ready'),
);

const liveMatrix = buildMeetingPlatformLiveAdapterMatrix({
  baseUrl,
  env: googleEnv,
  platforms: ['google-meet', 'zoom', 'teams'],
});
assert.equal(liveMatrix.schema, 'meeting_platform_live_adapter_matrix');
assert.deepEqual(liveMatrix.platforms, ['google_meet', 'zoom', 'microsoft_teams']);
assert.equal(liveMatrix.rows.every((row) => row.provider_blocks_realtime === false), true);
assert.equal(liveMatrix.rows.every((row) => row.transcript_blocks_realtime === false), true);

const suite = createMeetingPlatformLiveAdapterSuite(client, {
  baseUrl,
  env: googleEnv,
  platforms: ['google-meet', 'zoom'],
});
assert.deepEqual(suite.platforms, ['google_meet', 'zoom']);
assert.equal(suite.plan('google-meet').platform, 'google_meet');
assert.equal(suite.matrix().platform_count, 2);
assert.equal(suite.summary().platform_count, 2);
assert.equal(suite.adapter('google-meet'), suite.adapter('google-meet'));
assert.equal(suite.adapters().zoom.platform, 'zoom');
assert.equal(suite.readiness('google-meet').platform, 'google_meet');
assert.equal(suite.readinessMatrix().platform_count, 2);
assert.throws(
  () => suite.assertReadiness('google-meet'),
  (error) => error.name === 'MeetingTimelineSdkError'
    && error.details.platform === 'google_meet',
);

const activeSnapshot = buildMeetingAppFixtureSnapshot('google-meet', {
  state: 'active',
  observedAtMs: startMs,
});
const observed = await adapter.observeMeetingApp(activeSnapshot, {
  capturedAtMs: startMs,
});
assert.equal(observed.action, 'observe_meeting_app');
assert.equal(observed.source, 'meeting_app');
assert.equal(observed.rawSignals.some((signal) => signal.type === 'meeting_started'), true);
assert.equal(calls[0].method, 'startMeeting');
assert.equal(calls[0].input.platform, 'google_meet');
assert.equal(observed.live_evidence.can_insert_realtime_marks, true);
assert.equal(observed.live_evidence.status, 'realtime_ready_provider_pending');
assert.equal(observed.live_evidence.meeting_app_record_count, 1);

const inserted = await adapter.insertAnnotation({
  id: 'why-mark-1',
  capturedAtMs: startMs + 10_000,
  label: 'why?',
  kind: 'question',
});
assert.equal(inserted.action, 'insert_annotation');
assert.equal(inserted.result.mode, 'realtime_annotation_pipeline');
assert.equal(inserted.result.status, 'ready_to_insert');
assert.equal(inserted.result.pipeline.status, 'ready_to_insert');
assert.equal(inserted.result.pipeline.selected_meeting.meeting_id, 'abc-defg-hij');
assert.equal(calls.at(-1).method, 'insertMark');
assert.equal(calls.at(-1).input.id, 'why-mark-1');
assert.equal(calls.at(-1).input.time_ms, 10_000);
assert.equal(inserted.live_evidence.can_insert_realtime_marks, true);

const beforeStrictClockInsertCount = calls.length;
const strictClockBlocked = await adapter.insertAnnotation({
  id: 'strict-clock-note',
  capturedAtMs: startMs + 11_000,
  label: 'strict clock',
}, {
  requireClockSync: true,
});
assert.equal(strictClockBlocked.result.status, 'needs_clock_sync');
assert.equal(strictClockBlocked.result.pipeline.accepted_for_realtime, false);
assert.equal(calls.length, beforeStrictClockInsertCount);

const routedCalls = [];
const routedClient = {
  async startMeeting(input) {
    routedCalls.push({ method: 'startMeeting', input });
    return { ok: true, input };
  },
  async endMeeting(input) {
    routedCalls.push({ method: 'endMeeting', input });
    return { ok: true, input };
  },
  async insertMark(input, options) {
    routedCalls.push({ method: 'insertMark', input, options });
    return { ok: true, input };
  },
  async insertMarks(input, options) {
    routedCalls.push({ method: 'insertMarks', input, options });
    return { ok: true, input };
  },
};
const routedAdapter = createMeetingPlatformLiveAdapter('zoom', routedClient, {
  id: 'zoom-route-from-local-observer-test',
  baseUrl,
});
const routedAnnotation = await routedAdapter.insertAnnotation({
  id: 'zoom-route-note',
  capturedAtMs: startMs + 5_000,
  label: 'follow up',
  local_observer: {
    url: 'https://zoom.us/j/987654321',
    observed_at_ms: startMs,
  },
});
assert.equal(routedAnnotation.result.status, 'start_axis_then_insert');
assert.deepEqual(routedAnnotation.result.actions, ['start_meeting_session', 'insert_mark']);
assert.deepEqual(routedCalls.map((call) => call.method), ['startMeeting', 'insertMark']);
assert.equal(routedCalls[0].input.meeting_id, '987654321');
assert.equal(routedCalls[1].input.id, 'zoom-route-note');
assert.equal(routedCalls[1].input.time_ms, 5_000);

const beforePendingCount = routedCalls.length;
const pendingAnnotation = await routedAdapter.insertAnnotation({
  id: 'zoom-pending-note',
  capturedAtMs: startMs + 6_000,
  label: 'pending without meeting identity',
});
assert.equal(pendingAnnotation.result.status, 'pending_real_meeting');
assert.equal(pendingAnnotation.result.pending_payload.id, 'zoom-pending-note');
assert.equal(routedCalls.length, beforePendingCount);

const providerStart = await adapter.ingestProvider({
  method: 'POST',
  url: `${baseUrl}/api/platform-events/google-meet`,
  body: buildPlatformFixtureEvent('google-meet', 'meeting_start', {
    startMs,
    durationMs,
  }),
}, undefined, {
  capturedAtMs: startMs,
});
assert.equal(providerStart.action, 'ingest_provider');
assert.equal(providerStart.source, 'provider');
assert.equal(providerStart.live_evidence.provider_record_count, 1);
assert.equal(providerStart.live_evidence.provider_reconcile_ready, false);

const providerEnd = await adapter.ingestProvider({
  method: 'POST',
  url: `${baseUrl}/api/platform-events/google-meet`,
  body: buildPlatformFixtureEvent('google-meet', 'meeting_end', {
    startMs,
    durationMs,
  }),
}, undefined, {
  capturedAtMs: startMs + durationMs,
});
assert.equal(providerEnd.live_evidence.provider_record_count, 2);

const endedSnapshot = buildMeetingAppFixtureSnapshot('google-meet', {
  state: 'prejoin',
  observedAtMs: startMs + durationMs,
});
const ended = await adapter.observeMeetingApp(endedSnapshot, {
  phase: 'ended',
  capturedAtMs: startMs + durationMs,
});
assert.equal(ended.live_evidence.production_ready, true);
assert.equal(ended.live_evidence.correlation_status, 'matched');
assert.equal(adapter.exportPackage().rollout_plan.status, 'production_ready');
assert.equal(adapter.verify().passed, true);
assert.equal(adapter.correlation().passed, true);

const productionReadiness = buildMeetingPlatformLiveAdapterReadiness('google-meet', {
  baseUrl,
  env: googleEnv,
  target: 'production',
  adapter,
  evidencePackage: adapter.exportPackage(),
});
assert.equal(productionReadiness.status, 'ready');
assert.equal(productionReadiness.passed, true);
assert.equal(productionReadiness.production_ready, true);
assert.equal(productionReadiness.candidate_observation_ready, true);
assert.equal(productionReadiness.verification.passed, true);
assert.equal(productionReadiness.checks.some((item) => item.code === 'adapter_methods_available' && item.passed), true);
assert.equal(productionReadiness.checks.some((item) => item.code === 'candidate_observation_contract_ready' && item.passed), true);
assert.equal(assertMeetingPlatformLiveAdapterReadiness('google-meet', {
  baseUrl,
  env: googleEnv,
  target: 'production',
  adapter,
  evidencePackage: adapter.exportPackage(),
}).status, 'ready');

const readinessMatrix = buildMeetingPlatformLiveAdapterReadinessMatrix({
  baseUrl,
  env: googleEnv,
  target: 'production',
  platforms: ['google-meet'],
  evidencePackage: { google_meet: adapter.exportPackage() },
});
assert.equal(readinessMatrix.schema, 'meeting_platform_live_adapter_readiness_matrix');
assert.equal(readinessMatrix.platform_count, 1);
assert.equal(readinessMatrix.ready_count, 1);
assert.equal(readinessMatrix.candidate_observer_count, 1);
assert.equal(readinessMatrix.rows[0].status, 'ready');
assert.equal(readinessMatrix.rows[0].candidate_observer_message_type, 'meeting_timeline.observe_candidates');
assert.equal(assertMeetingPlatformLiveAdapterReadinessMatrix({
  baseUrl,
  env: googleEnv,
  target: 'production',
  platforms: ['google-meet'],
  evidencePackage: { google_meet: adapter.exportPackage() },
}).passed_count, 1);
assert.throws(
  () => assertMeetingPlatformLiveAdapterReadinessMatrix({
    baseUrl,
    env: googleEnv,
    target: 'production',
    platforms: ['google-meet', 'teams'],
    evidencePackage: { google_meet: adapter.exportPackage() },
  }),
  (error) => error.name === 'MeetingTimelineSdkError'
    && error.details.failed_platforms.includes('microsoft_teams')
    && error.details.matrix.blocked_count === 1,
);
assert.equal(suite.assertReadiness('google-meet', {
  target: 'production',
  evidencePackage: adapter.exportPackage(),
}).status, 'ready');
assert.equal(suite.assertReadinessMatrix({
  target: 'production',
  platforms: ['google-meet'],
  evidencePackage: { google_meet: adapter.exportPackage() },
}).passed_count, 1);

const brokenCandidateReadiness = buildMeetingPlatformLiveAdapterReadiness('google-meet', {
  baseUrl,
  env: googleEnv,
  target: 'production',
  adapter,
  evidencePackage: adapter.exportPackage(),
  candidateObservation: {
    runtime_event_action: 'observe_meeting_app',
  },
});
assert.equal(brokenCandidateReadiness.passed, false);
assert.equal(brokenCandidateReadiness.candidate_observation_ready, false);
assert.equal(brokenCandidateReadiness.blocking_checks.some((item) => item.code === 'candidate_observation_contract_ready'), true);

const googleHandoff = buildMeetingPlatformLiveAdapterHandoff('google-meet', {
  baseUrl,
  env: googleEnv,
  target: 'production',
  adapter,
  evidencePackage: adapter.exportPackage(),
});
assert.equal(googleHandoff.schema, 'meeting_platform_live_adapter_handoff');
assert.equal(googleHandoff.platform, 'google_meet');
assert.equal(googleHandoff.passed, true);
assert.equal(googleHandoff.host_contract.annotation_timestamp_field, 'captured_at_ms');
assert.equal(googleHandoff.host_contract.can_insert_before_provider_event, true);
assert.equal(googleHandoff.candidate_observation_ready, true);
assert.equal(googleHandoff.candidate_observation_contract.endpoint, '/api/meeting-platform/observe-candidates');
assert.equal(googleHandoff.host_contract.candidate_observation_runtime_action, 'observe_platform_candidates');
assert.equal(googleHandoff.sdk.factory, 'createMeetingPlatformLiveAdapter');
assert.equal(googleHandoff.sdk.kit_methods.includes('platformLiveAdapterHandoffBundle'), true);
assert.equal(googleHandoff.commands.validate_live_readiness, 'npm run meeting-platform:live-readiness');
assert.equal(googleHandoff.evidence_paths.evidence_package, 'data/meeting-platform-evidence-packages/google_meet.json');
assert.equal(googleHandoff.runbook.steps.some((step) => step.id === 'validate_rollout'), true);
assert.equal(googleHandoff.realtime_flow.includes('insert annotation marks with captured_at_ms onto the active axis'), true);

const handoffBundle = buildMeetingPlatformLiveAdapterHandoffBundle({
  baseUrl,
  env: googleEnv,
  target: 'production',
  platforms: ['google-meet', 'zoom'],
  evidencePackage: { google_meet: adapter.exportPackage() },
});
assert.equal(handoffBundle.schema, 'meeting_platform_live_adapter_handoff_bundle');
assert.deepEqual(handoffBundle.platforms, ['google_meet', 'zoom']);
assert.equal(handoffBundle.platform_count, 2);
assert.equal(handoffBundle.ready_count, 1);
assert.equal(handoffBundle.blocked_count, 1);
assert.equal(handoffBundle.candidate_observer_count, 2);
assert.equal(handoffBundle.host_contract.provider_events_block_realtime, false);
assert.equal(handoffBundle.host_contract.candidate_observation_message_type, 'meeting_timeline.observe_candidates');
assert.equal(handoffBundle.handoffs[0].schema, 'meeting_platform_live_adapter_handoff');
assert.equal(handoffBundle.readiness_matrix.platform_count, 2);
assert.equal(handoffBundle.live_adapter_matrix.platform_count, 2);
assert.equal(handoffBundle.commands.package_smoke, 'npm run sdk:package-smoke');

assert.equal(suite.handoff('google-meet', {
  target: 'production',
  evidencePackage: adapter.exportPackage(),
}).passed, true);
assert.equal(suite.handoffBundle({
  target: 'production',
  platforms: ['google-meet'],
  evidencePackage: { google_meet: adapter.exportPackage() },
}).ready_count, 1);

const state = adapter.getState();
assert.equal(state.platform, 'google_meet');
assert.equal(state.evidence.provider_record_count, 2);
assert.equal(state.evidence.meeting_app_record_count, 2);

const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl,
  verify: false,
});
const zoomAdapter = kit.platformLiveAdapter('zoom', {
  id: 'zoom-live-adapter-test',
});
assert.equal(zoomAdapter.platform, 'zoom');
assert.equal(zoomAdapter.summary().platform, 'zoom');
assert.equal(kit.platformLiveAdapterPlan('zoom').platform, 'zoom');
assert.equal(kit.platformLiveAdapterMatrix({ platforms: ['google-meet'] }).rows[0].platform, 'google_meet');
assert.equal(kit.platformLiveAdapterReadiness('google-meet', {
  evidencePackage: adapter.exportPackage(),
}).passed, true);
assert.equal(kit.assertPlatformLiveAdapterReadiness('google-meet', {
  evidencePackage: adapter.exportPackage(),
}).passed, true);
assert.equal(kit.platformLiveAdapterReadinessMatrix({
  platforms: ['google-meet'],
  evidencePackage: { google_meet: adapter.exportPackage() },
}).passed_count, 1);
assert.equal(kit.assertPlatformLiveAdapterReadinessMatrix({
  platforms: ['google-meet'],
  evidencePackage: { google_meet: adapter.exportPackage() },
}).passed_count, 1);
assert.equal(kit.platformLiveAdapterHandoff('google-meet', {
  evidencePackage: adapter.exportPackage(),
}).host_contract.annotation_timestamp_field, 'captured_at_ms');
assert.equal(kit.platformLiveAdapterHandoffBundle({
  platforms: ['google-meet'],
  evidencePackage: { google_meet: adapter.exportPackage() },
}).handoffs[0].sdk.kit_module, '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit');
assert.equal(kit.platformLiveAdapterSuite({ platforms: ['zoom'] }).summary().platform_count, 1);

assert.equal(adapter.reset().evidence.removed.meeting_app_records, 2);
assert.equal(adapter.summary().package_ready_for_handoff, false);

console.log('ok meeting platform live adapter');
