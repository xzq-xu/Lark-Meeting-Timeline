import assert from 'node:assert/strict';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import { buildMeetingAppSnapshotRecordSet } from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';
import { buildMeetingPlatformEvidencePackage, buildMeetingPlatformEvidencePackageSummary, createMeetingPlatformEvidencePackageBuilder, verifyMeetingPlatformEvidencePackage } from '../packages/meeting-timeline-sdk/adapters/platform-evidence-package.mjs';
import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';
import { capturePlatformWebhookEvent } from '../packages/meeting-timeline-sdk/adapters/platform-capture.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';
const capturedAtMs = 1_783_702_000_000;
const googleEnv = {
  GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
  GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL: 'meet-events@example.iam.gserviceaccount.com',
};

function providerRecords(platform) {
  return [
    capturePlatformWebhookEvent(platform, buildPlatformFixtureEvent(platform, 'meeting_start', {
      startMs: capturedAtMs,
      durationMs: 120_000,
    }), {
      capturedAtMs,
      label: `${platform} start`,
    }),
    capturePlatformWebhookEvent(platform, buildPlatformFixtureEvent(platform, 'meeting_end', {
      startMs: capturedAtMs,
      durationMs: 120_000,
    }), {
      capturedAtMs: capturedAtMs + 120_000,
      label: `${platform} end`,
    }),
  ];
}

function meetingAppRecordSet(platform) {
  return buildMeetingAppSnapshotRecordSet([
    {
      platform,
      phase: 'active',
      capturedAtMs,
      snapshot: buildMeetingAppFixtureSnapshot(platform, {
        observedAtMs: capturedAtMs,
        state: 'active',
      }),
    },
    {
      platform,
      phase: 'ended',
      capturedAtMs: capturedAtMs + 120_000,
      snapshot: buildMeetingAppFixtureSnapshot(platform, {
        observedAtMs: capturedAtMs + 120_000,
        state: 'prejoin',
      }),
    },
  ], {
    id: `${platform}-meeting-app-record-set`,
    createdAtMs: capturedAtMs + 121_000,
  });
}

const googlePackage = buildMeetingPlatformEvidencePackage('google-meet', {
  providerRecords: providerRecords('google-meet'),
  meetingAppRecordSet: meetingAppRecordSet('google-meet'),
}, {
  baseUrl,
  env: googleEnv,
  createdAtMs: capturedAtMs + 122_000,
  label: 'google meet pilot evidence',
});

assert.equal(googlePackage.schema, 'meeting_platform_evidence_package');
assert.equal(googlePackage.platform, 'google_meet');
assert.equal(googlePackage.provider_records.length, 2);
assert.equal(googlePackage.meeting_app_record_set.record_count, 2);
assert.equal(googlePackage.rollout_plan.status, 'production_ready');
assert.equal(googlePackage.rollout_plan.production_ready, true);
assert.equal(googlePackage.adapter_route.schema, 'meeting_platform_adapter_route');
assert.equal(googlePackage.adapter_route.recommended_mode, 'local_observer_first_provider_reconcile');
assert.equal(googlePackage.adapter_route.routes[0].route, 'local_observer_axis');
assert.equal(googlePackage.adapter_route.realtime_invariants.provider_events_block_realtime, false);
assert.equal(googlePackage.adapter_route.realtime_invariants.transcript_blocks_realtime, false);
assert.equal(googlePackage.evidence_correlation.passed, true);
assert.equal(googlePackage.evidence_correlation.confidence, 'high');
assert.equal(googlePackage.evidence_summary.provider.evidence_level, 'captured_events');
assert.equal(googlePackage.evidence_summary.local_dom.evidence_level, 'captured_dom');
assert.equal(googlePackage.evidence_summary.adapter_route.first_route, 'local_observer_axis');
assert.equal(googlePackage.evidence_summary.adapter_route.provider_events_block_realtime, false);
assert.equal(googlePackage.evidence_summary.correlation.passed, true);
assert.deepEqual(googlePackage.env_summary.configured_keys, [
  'GOOGLE_PUBSUB_OIDC_AUDIENCE',
  'GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL',
]);
assert.equal(googlePackage.env_summary.values, undefined);
assert.equal(googlePackage.handoff.final_gate, 'buildMeetingPlatformRolloutPlan');
assert.equal(googlePackage.handoff.adapter_route_gate, 'buildMeetingPlatformAdapterRoute');
assert.equal(googlePackage.handoff.adapter_first_route, 'local_observer_axis');
assert.equal(googlePackage.runbook.steps.some((item) => item.id === 'validate_rollout'), true);

const googleSummary = buildMeetingPlatformEvidencePackageSummary(googlePackage);
assert.equal(googleSummary.type, 'meeting_platform_evidence_package_summary');
assert.equal(googleSummary.package_id, googlePackage.id);
assert.equal(googleSummary.production_ready, true);
assert.equal(googleSummary.provider_record_count, 2);
assert.equal(googleSummary.meeting_app_record_count, 2);
assert.equal(googleSummary.correlation_passed, true);
assert.equal(googleSummary.correlation_confidence, 'high');
assert.equal(googleSummary.adapter_route_ready, true);
assert.equal(googleSummary.adapter_recommended_mode, 'local_observer_first_provider_reconcile');
assert.equal(googleSummary.adapter_first_route, 'local_observer_axis');
assert.equal(googleSummary.provider_events_block_realtime, false);
assert.equal(googleSummary.transcript_blocks_realtime, false);

const googleVerification = verifyMeetingPlatformEvidencePackage(googlePackage, {
  baseUrl,
  env: googleEnv,
});
assert.equal(googleVerification.type, 'meeting_platform_evidence_package_verification');
assert.equal(googleVerification.passed, true);
assert.equal(googleVerification.embedded_plan_matches, true);
assert.equal(googleVerification.requirement, 'production_ready');
assert.equal(googleVerification.correlation_required, true);
assert.equal(googleVerification.correlation_passed, true);
assert.equal(googleVerification.adapter_route_ready, true);
assert.equal(googleVerification.adapter_first_route, 'local_observer_axis');
assert.equal(googleVerification.provider_events_block_realtime, false);
assert.equal(googleVerification.transcript_blocks_realtime, false);

const zoomDomOnly = buildMeetingPlatformEvidencePackage({
  platform: 'zoom',
  baseUrl,
  meetingAppRecordSet: meetingAppRecordSet('zoom'),
  includeRunbook: false,
});
assert.equal(zoomDomOnly.rollout_plan.status, 'realtime_ready_provider_pending');
assert.equal(zoomDomOnly.rollout_plan.ready_for_realtime_annotations, true);
assert.equal(zoomDomOnly.rollout_plan.production_ready, false);
assert.equal(zoomDomOnly.runbook, undefined);
assert.equal(zoomDomOnly.evidence_summary.provider.record_count, 0);
assert.equal(zoomDomOnly.evidence_summary.local_dom.record_count, 2);
assert.equal(verifyMeetingPlatformEvidencePackage(zoomDomOnly, { baseUrl }).passed, false);
assert.equal(verifyMeetingPlatformEvidencePackage(zoomDomOnly, {
  baseUrl,
  requireProductionReady: false,
}).passed, true);

const builder = createMeetingPlatformEvidencePackageBuilder('webex', {
  baseUrl,
  env: {
    WEBEX_WEBHOOK_SECRET: 'secret',
  },
  createdAtMs: capturedAtMs + 200_000,
});
builder.captureProviderWebhook({
  method: 'POST',
  url: `${baseUrl}/api/platform-events/webex`,
  body: buildPlatformFixtureEvent('webex', 'meeting_start', {
    startMs: capturedAtMs,
    durationMs: 120_000,
  }),
}, undefined, {
  capturedAtMs,
});
builder.captureProviderWebhook({
  method: 'POST',
  url: `${baseUrl}/api/platform-events/webex`,
  body: buildPlatformFixtureEvent('webex', 'meeting_end', {
    startMs: capturedAtMs,
    durationMs: 120_000,
  }),
}, undefined, {
  capturedAtMs: capturedAtMs + 120_000,
});
builder.addMeetingAppRecord({
  phase: 'active',
  capturedAtMs,
  snapshot: buildMeetingAppFixtureSnapshot('webex', {
    observedAtMs: capturedAtMs,
    state: 'active',
  }),
});
builder.addMeetingAppRecord({
  phase: 'ended',
  capturedAtMs: capturedAtMs + 120_000,
  snapshot: buildMeetingAppFixtureSnapshot('webex', {
    observedAtMs: capturedAtMs + 120_000,
    state: 'prejoin',
  }),
});

assert.equal(builder.getState().provider_record_count, 2);
assert.equal(builder.getState().meeting_app_record_count, 2);
const webexPackage = builder.exportPackage({ label: 'webex field evidence' });
assert.equal(webexPackage.platform, 'webex');
assert.equal(webexPackage.rollout_plan.production_ready, true);
assert.equal(webexPackage.evidence_correlation.passed, true);
assert.equal(builder.summary().production_ready, true);

const kit = createMeetingPlatformTimelineKit({ baseUrl, env: googleEnv, verify: false });
const kitPackage = kit.platformEvidencePackage('google-meet', {
  providerRecords: providerRecords('google-meet'),
  meetingAppRecordSet: meetingAppRecordSet('google-meet'),
});
assert.equal(kitPackage.rollout_plan.production_ready, true);
assert.equal(kit.platformEvidencePackageSummary(kitPackage).provider_record_count, 2);
assert.equal(kit.verifyPlatformEvidencePackage(kitPackage).passed, true);
const kitBuilder = kit.platformEvidencePackageBuilder('google-meet');
kitBuilder.addProviderRecord(providerRecords('google-meet')[0]);
assert.equal(kitBuilder.getState().provider_record_count, 1);

const mismatchedPackage = buildMeetingPlatformEvidencePackage('google-meet', {
  providerRecords: providerRecords('google-meet'),
  meetingAppRecordSet: meetingAppRecordSet('google-meet'),
}, {
  baseUrl,
  env: googleEnv,
  createdAtMs: capturedAtMs + 5 * 60 * 60 * 1000,
  maxClockSkewMs: 1_000,
});
mismatchedPackage.meeting_app_record_set.records = mismatchedPackage.meeting_app_record_set.records.map((record) => ({
  ...record,
  capturedAtMs: record.capturedAtMs + 3 * 60 * 60 * 1000,
  captured_at_ms: record.captured_at_ms + 3 * 60 * 60 * 1000,
}));
const mismatchedVerification = verifyMeetingPlatformEvidencePackage(mismatchedPackage, {
  baseUrl,
  env: googleEnv,
  maxClockSkewMs: 1_000,
});
assert.equal(mismatchedVerification.production_ready, true);
assert.equal(mismatchedVerification.correlation_passed, false);
assert.equal(mismatchedVerification.passed, false);
assert.equal(mismatchedVerification.evidence_correlation.issues.some((item) => item.code === 'time_window_mismatch'), true);

console.log('ok meeting platform evidence package');
