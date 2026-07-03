import assert from 'node:assert/strict';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';
import {
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
assert.equal(blockedReadiness.blocking_checks.some((item) => item.code === 'pilot_realtime_axis_ready'), true);
assert.equal(blockedReadiness.warnings.some((item) => item.code === 'evidence_package_available'), true);

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
assert.equal(calls.at(-1).method, 'insertMark');
assert.equal(calls.at(-1).input.id, 'why-mark-1');
assert.equal(inserted.live_evidence.can_insert_realtime_marks, true);

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
assert.equal(productionReadiness.verification.passed, true);
assert.equal(productionReadiness.checks.some((item) => item.code === 'adapter_methods_available' && item.passed), true);

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
assert.equal(readinessMatrix.rows[0].status, 'ready');

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
assert.equal(kit.platformLiveAdapterReadinessMatrix({
  platforms: ['google-meet'],
  evidencePackage: { google_meet: adapter.exportPackage() },
}).passed_count, 1);
assert.equal(kit.platformLiveAdapterSuite({ platforms: ['zoom'] }).summary().platform_count, 1);

assert.equal(adapter.reset().evidence.removed.meeting_app_records, 2);
assert.equal(adapter.summary().package_ready_for_handoff, false);

console.log('ok meeting platform live adapter');
