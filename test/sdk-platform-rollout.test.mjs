import assert from 'node:assert/strict';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import { buildMeetingAppSnapshotRecordSet } from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';
import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';
import { capturePlatformWebhookEvent } from '../packages/meeting-timeline-sdk/adapters/platform-capture.mjs';
import {
  MEETING_PLATFORM_ROLLOUT_STATUSES,
  buildAllMeetingPlatformRolloutPlans,
  buildMeetingPlatformRolloutPlan,
  buildMeetingPlatformRolloutSummary,
} from '../packages/meeting-timeline-sdk/adapters/platform-rollout.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';
const observedAtMs = 1_783_356_000_000;
const googleEnv = {
  GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
};

function providerRecords(platform) {
  return [
    capturePlatformWebhookEvent(platform, buildPlatformFixtureEvent(platform, 'meeting_start'), {
      capturedAtMs: observedAtMs,
    }),
    capturePlatformWebhookEvent(platform, buildPlatformFixtureEvent(platform, 'meeting_end'), {
      capturedAtMs: observedAtMs + 60_000,
    }),
  ];
}

function meetingAppRecordSet(platform) {
  const key = platform.replaceAll('-', '_');
  return buildMeetingAppSnapshotRecordSet([
    {
      platform: key,
      phase: 'active',
      capturedAtMs: observedAtMs,
      snapshot: buildMeetingAppFixtureSnapshot(platform, {
        observedAtMs,
        state: 'active',
      }),
    },
    {
      platform: key,
      phase: 'ended',
      capturedAtMs: observedAtMs + 60_000,
      snapshot: buildMeetingAppFixtureSnapshot(platform, {
        observedAtMs: observedAtMs + 60_000,
        state: 'prejoin',
      }),
    },
  ], {
    id: `${key}-record-set`,
    createdAtMs: observedAtMs + 61_000,
  });
}

assert.equal(MEETING_PLATFORM_ROLLOUT_STATUSES.includes('production_ready'), true);

const googleMissingEvidence = buildMeetingPlatformRolloutPlan('google-meet', {
  baseUrl,
  env: googleEnv,
});
assert.equal(googleMissingEvidence.status, 'needs_live_dom_and_provider_evidence');
assert.equal(googleMissingEvidence.ready_for_realtime_annotations, false);
assert.equal(googleMissingEvidence.local_observer.runtime_ready, true);
assert.equal(googleMissingEvidence.local_observer.production_ready, false);
assert.equal(googleMissingEvidence.provider_events.setup_ready, true);
assert.equal(googleMissingEvidence.provider_events.production_ready, false);
assert.equal(googleMissingEvidence.next_actions.includes('capture_live_dom_snapshots_for_local_observer'), true);
assert.equal(googleMissingEvidence.next_actions.includes('capture_real_provider_start_end_events'), true);
assert.equal(googleMissingEvidence.recommended_mode, 'capture_live_dom_snapshots_first');

const googleReady = buildMeetingPlatformRolloutPlan('google-meet', {
  baseUrl,
  env: googleEnv,
  providerRecords: providerRecords('google-meet'),
  meetingAppRecordSet: meetingAppRecordSet('google-meet'),
});
assert.equal(googleReady.status, 'production_ready');
assert.equal(googleReady.production_ready, true);
assert.equal(googleReady.ready_for_realtime_annotations, true);
assert.equal(googleReady.local_observer.evidence_level, 'captured_dom');
assert.equal(googleReady.provider_events.evidence_level, 'captured_events');
assert.equal(googleReady.provider_events.event_types.includes('google.workspace.meet.conference.v2.started'), true);
assert.equal(googleReady.source_priority[0], 'local_observer');
assert.equal(googleReady.next_actions.includes('enable_pilot_rollout_with_monitoring'), true);

const zoomLocalReady = buildMeetingPlatformRolloutPlan('zoom', {
  baseUrl,
  meetingAppRecordSet: meetingAppRecordSet('zoom'),
});
assert.equal(zoomLocalReady.status, 'realtime_ready_provider_pending');
assert.equal(zoomLocalReady.ready_for_realtime_annotations, true);
assert.equal(zoomLocalReady.production_ready, false);
assert.equal(zoomLocalReady.local_observer.production_ready, true);
assert.equal(zoomLocalReady.provider_events.setup_ready, false);
assert.equal(zoomLocalReady.next_actions.includes('complete_provider_setup_readiness'), true);

const summary = buildMeetingPlatformRolloutSummary({
  baseUrl,
  env: googleEnv,
  platforms: ['google-meet', 'zoom'],
  providerRecords: providerRecords('google-meet'),
  meetingAppRecordSet: meetingAppRecordSet('google-meet'),
});
assert.equal(summary.plan_count, 2);
assert.equal(summary.production_ready_count, 1);
assert.equal(summary.realtime_ready_count, 1);
assert.deepEqual(summary.pending_platforms, ['zoom']);

const allPlans = buildAllMeetingPlatformRolloutPlans({
  baseUrl,
  platforms: ['google-meet', 'teams', 'webex'],
});
assert.deepEqual(allPlans.map((plan) => plan.platform), ['google_meet', 'microsoft_teams', 'webex']);

const allDefaultSummary = buildMeetingPlatformRolloutSummary({ baseUrl });
assert.equal(allDefaultSummary.plans.some((plan) => plan.platform === 'local_detector'), true);
assert.equal(allDefaultSummary.plans.find((plan) => plan.platform === 'local_detector').local_observer, undefined);

const kit = createMeetingPlatformTimelineKit({ baseUrl, env: googleEnv, verify: false });
assert.equal(
  kit.platformRolloutPlan('google-meet', {
    providerRecords: providerRecords('google-meet'),
    meetingAppRecordSet: meetingAppRecordSet('google-meet'),
  }).production_ready,
  true,
);
assert.equal(kit.platformRolloutSummary({ platforms: ['google-meet'] }).plans.length, 1);
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_rollout.type, 'meeting_platform_rollout_summary');

console.log('ok meeting platform rollout');
