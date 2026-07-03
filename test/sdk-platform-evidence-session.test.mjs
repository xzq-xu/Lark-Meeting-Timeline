import assert from 'node:assert/strict';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';
import {
  createMeetingPlatformEvidenceSession,
} from '../packages/meeting-timeline-sdk/adapters/platform-evidence-session.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';
const capturedAtMs = 1_783_961_000_000;
const googleEnv = {
  GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
  GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL: 'meet-events@example.iam.gserviceaccount.com',
};

function addDomLifecycle(session, platform, startMs = capturedAtMs, durationMs = 90_000) {
  session.captureMeetingAppSnapshot({
    phase: 'active',
    capturedAtMs: startMs,
    snapshot: buildMeetingAppFixtureSnapshot(platform, {
      observedAtMs: startMs,
      state: 'active',
    }),
  });
  session.captureMeetingAppSnapshot({
    phase: 'ended',
    capturedAtMs: startMs + durationMs,
    snapshot: buildMeetingAppFixtureSnapshot(platform, {
      observedAtMs: startMs + durationMs,
      state: 'prejoin',
    }),
  });
}

function addProviderLifecycle(session, platform, startMs = capturedAtMs, durationMs = 90_000) {
  session.captureProviderWebhook({
    method: 'POST',
    url: `${baseUrl}/api/platform-events/${platform}`,
    body: buildPlatformFixtureEvent(platform, 'meeting_start', {
      startMs,
      durationMs,
    }),
  }, undefined, {
    capturedAtMs: startMs,
  });
  session.captureProviderWebhook({
    method: 'POST',
    url: `${baseUrl}/api/platform-events/${platform}`,
    body: buildPlatformFixtureEvent(platform, 'meeting_end', {
      startMs,
      durationMs,
    }),
  }, undefined, {
    capturedAtMs: startMs + durationMs,
  });
}

const googleSession = createMeetingPlatformEvidenceSession('google-meet', {
  id: 'google-live-session',
  baseUrl,
  env: googleEnv,
  createdAtMs: capturedAtMs - 1_000,
});

assert.equal(googleSession.platform, 'google_meet');
assert.equal(googleSession.getState().provider_record_count, 0);
assert.equal(googleSession.getState().meeting_app_record_count, 0);

const emptySummary = googleSession.summary();
assert.equal(emptySummary.type, 'meeting_platform_evidence_session_summary');
assert.equal(emptySummary.platform, 'google_meet');
assert.equal(emptySummary.package_ready_for_handoff, false);
assert.equal(emptySummary.production_ready, false);
assert.equal(emptySummary.can_insert_realtime_marks, false);

addDomLifecycle(googleSession, 'google-meet');
const domOnlySummary = googleSession.summary();
assert.equal(domOnlySummary.package_ready_for_handoff, true);
assert.equal(domOnlySummary.provider_record_count, 0);
assert.equal(domOnlySummary.meeting_app_record_count, 2);
assert.equal(domOnlySummary.ready_for_realtime_annotations, true);
assert.equal(domOnlySummary.can_insert_realtime_marks, true);
assert.equal(domOnlySummary.production_ready, false);
assert.equal(domOnlySummary.provider_reconcile_ready, false);
assert.equal(domOnlySummary.correlation_status, 'single_source');
assert.equal(domOnlySummary.verification_passed, false);

addProviderLifecycle(googleSession, 'google-meet');
const productionSummary = googleSession.summary({
  includePackageSummary: true,
  includeVerification: true,
});
assert.equal(productionSummary.production_ready, true);
assert.equal(productionSummary.provider_reconcile_ready, true);
assert.equal(productionSummary.local_observer_ready, true);
assert.equal(productionSummary.verification_passed, true);
assert.equal(productionSummary.correlation_status, 'matched');
assert.equal(productionSummary.correlation_confidence, 'high');
assert.equal(productionSummary.evidence_package_summary.provider_record_count, 2);
assert.equal(productionSummary.verification.passed, true);

const googleCorrelation = googleSession.correlation();
assert.equal(googleCorrelation.passed, true);
assert.equal(googleCorrelation.coverage.provider_has_start, true);
assert.equal(googleCorrelation.coverage.local_has_active, true);

const googlePackage = googleSession.exportPackage({ label: 'google live evidence' });
assert.equal(googlePackage.schema, 'meeting_platform_evidence_package');
assert.equal(googlePackage.id, 'google-live-session-package');
assert.equal(googlePackage.source, 'meeting_platform_evidence_session');
assert.equal(googlePackage.rollout_plan.status, 'production_ready');
assert.equal(googleSession.verify().passed, true);
assert.equal(googleSession.strategy().recommendation, 'enable_pilot_with_provider_reconcile_and_monitoring');

const kit = createMeetingPlatformTimelineKit({
  baseUrl,
  verify: false,
});
const zoomSession = kit.platformEvidenceSession('zoom', {
  id: 'zoom-live-session',
  createdAtMs: capturedAtMs,
});
addDomLifecycle(zoomSession, 'zoom');
assert.equal(zoomSession.summary().status, 'realtime_ready_provider_pending');
assert.equal(zoomSession.summary().can_insert_realtime_marks, true);
assert.equal(zoomSession.verify({ requireProductionReady: false }).passed, true);

assert.deepEqual(zoomSession.reset().removed, {
  provider_records: 0,
  provider_samples: 0,
  meeting_app_records: 2,
});
assert.equal(zoomSession.getState().meeting_app_record_count, 0);
assert.equal(zoomSession.summary().package_ready_for_handoff, false);

console.log('ok meeting platform evidence session');
