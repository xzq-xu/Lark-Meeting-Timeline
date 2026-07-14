import assert from 'node:assert/strict';

import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';
import {
  assertPlatformLaunchGate,
  buildFixtureLaunchGateInput,
  buildMeetingPlatformLaunchGateSummary,
  buildPlatformLaunchGate,
} from '../packages/meeting-timeline-sdk/adapters/platform-gate.mjs';
import { capturePlatformWebhookEvent } from '../packages/meeting-timeline-sdk/adapters/platform-capture.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import { platformCapabilityContract } from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';

const baseUrl = 'https://timeline.example.com';
const googleEnv = {
  GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
};

assert.equal(
  platformCapabilityContract('google-meet', { baseUrl }).sdk_modules.platform_gate,
  '@ai-annotation/meeting-timeline-sdk/adapters/platform-gate',
);

const googleRecords = [
  capturePlatformWebhookEvent('google-meet', buildPlatformFixtureEvent('google-meet', 'meeting_start'), {
    capturedAtMs: 1_782_442_800_000,
  }),
  capturePlatformWebhookEvent('google-meet', buildPlatformFixtureEvent('google-meet', 'meeting_end'), {
    capturedAtMs: 1_782_444_600_000,
  }),
];

const passingGate = buildPlatformLaunchGate('google-meet', {
  baseUrl,
  env: googleEnv,
  records: googleRecords,
  requireEndEvent: true,
});
assert.equal(passingGate.status, 'passed');
assert.equal(passingGate.passed, true);
assert.equal(passingGate.production_ready, true);
assert.equal(passingGate.evidence_level, 'captured_events');
assert.deepEqual(passingGate.missing_required_coverage, []);

const assertedGate = assertPlatformLaunchGate('google-meet', {
  baseUrl,
  env: googleEnv,
  records: googleRecords,
});
assert.equal(assertedGate.production_ready, true);

const missingEvidenceGate = buildPlatformLaunchGate('google-meet', {
  baseUrl,
  env: googleEnv,
});
assert.equal(missingEvidenceGate.status, 'failed');
assert.equal(missingEvidenceGate.blocking_issues.some((item) => item.code === 'missing_real_event_evidence'), true);

const fixtureOnlyGate = buildPlatformLaunchGate('google-meet', {
  ...buildFixtureLaunchGateInput({ baseUrl, env: googleEnv }),
});
assert.equal(fixtureOnlyGate.evidence_level, 'fixture_events');
assert.equal(fixtureOnlyGate.status, 'failed');
assert.equal(fixtureOnlyGate.production_ready, false);

const fixtureAllowedForNonProduction = assertPlatformLaunchGate('google-meet', {
  ...buildFixtureLaunchGateInput({ baseUrl, env: googleEnv }),
  allowFixtureProduction: true,
  requireProductionReady: false,
});
assert.equal(fixtureAllowedForNonProduction.status, 'warning');
assert.equal(fixtureAllowedForNonProduction.passed, true);

const missingArtifactGate = buildPlatformLaunchGate('google-meet', {
  baseUrl,
  env: googleEnv,
  records: googleRecords,
  requireArtifact: true,
});
assert.equal(missingArtifactGate.status, 'failed');
assert.equal(missingArtifactGate.missing_required_coverage.includes('artifact_ready'), true);
assert.throws(
  () => assertPlatformLaunchGate('google-meet', {
    baseUrl,
    env: googleEnv,
    records: googleRecords,
    requireArtifact: true,
  }),
  /launch gate failed/i,
);

const brokenRuntimeContractGate = buildPlatformLaunchGate('google-meet', {
  baseUrl,
  env: googleEnv,
  records: googleRecords,
  candidateObservation: {
    ready: false,
    message_type: 'wrong.message',
  },
});
assert.equal(brokenRuntimeContractGate.status, 'failed');
assert.equal(brokenRuntimeContractGate.onboarding_status, 'blocked_by_runtime_contract');
assert.equal(brokenRuntimeContractGate.blocking_issues.some((item) => item.code === 'runtime_contract_not_ready'), true);
assert.equal(brokenRuntimeContractGate.next_actions.includes('fix_candidate_observation:candidate_observation_not_ready'), true);

const summary = buildMeetingPlatformLaunchGateSummary({
  ...buildFixtureLaunchGateInput({
    baseUrl,
    env: {
      GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
      MICROSOFT_GRAPH_CLIENT_STATE: 'graph-state',
      ZOOM_WEBHOOK_SECRET_TOKEN: 'zoom-secret',
      WEBEX_WEBHOOK_SECRET: 'webex-secret',
    },
  }),
  allowFixtureProduction: true,
});
assert.equal(summary.gates.length, 6);
assert.equal(summary.ok, true);
assert.equal(summary.production_ready, false);

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
  async insertMark(input) {
    calls.push({ method: 'insertMark', input });
    return { ok: true, input };
  },
};
const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl,
  env: googleEnv,
  verify: false,
});
const kitGate = kit.launchGate('google-meet', { records: googleRecords });
assert.equal(kitGate.production_ready, true);
assert.equal(kit.assertLaunchGate('google-meet', { records: googleRecords }).passed, true);

console.log('ok meeting platform launch gate');
