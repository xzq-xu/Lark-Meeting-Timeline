import assert from 'node:assert/strict';

import {
  assertAllMeetingAppLaunchGates,
  assertMeetingAppLaunchGate,
  buildMeetingAppLaunchGate,
  buildMeetingAppLaunchGateSummary,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-gate.mjs';
import {
  MEETING_APP_FIXTURE_PLATFORMS,
  buildMeetingAppFixtureSnapshot,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';

const observedAtMs = 1_783_356_000_000;

const fixtureOnly = buildMeetingAppLaunchGate('google-meet', { observedAtMs });
assert.equal(fixtureOnly.type, 'meeting_app_launch_gate');
assert.equal(fixtureOnly.platform, 'google_meet');
assert.equal(fixtureOnly.evidence_level, 'fixture_dom');
assert.equal(fixtureOnly.status, 'failed');
assert.equal(fixtureOnly.passed, false);
assert.equal(fixtureOnly.production_ready, false);
assert.equal(fixtureOnly.runtime_ready, true);
assert.equal(fixtureOnly.coverage.runtime_preset, true);
assert.equal(fixtureOnly.coverage.capture_profile, true);
assert.equal(fixtureOnly.coverage.mutation_observer, true);
assert.equal(fixtureOnly.coverage.mutation_track_selectors, true);
assert.equal(fixtureOnly.coverage.mutation_ignore_selectors, true);
assert.equal(fixtureOnly.coverage.meeting_started, true);
assert.equal(fixtureOnly.coverage.speaker_started, true);
assert.equal(fixtureOnly.coverage.meeting_ended, true);
assert.equal(fixtureOnly.blocking_issues.some((item) => item.code === 'fixture_dom_only'), true);
assert.equal(fixtureOnly.recommended_runtime.runtimePreset, 'google_meet');
assert.equal(fixtureOnly.recommended_runtime.mutation_track_selector_count > 0, true);

const fixtureAllowed = assertMeetingAppLaunchGate('google-meet', {
  observedAtMs,
  allowFixtureProduction: true,
  requireProductionReady: false,
});
assert.equal(fixtureAllowed.status, 'warning');
assert.equal(fixtureAllowed.passed, true);
assert.equal(fixtureAllowed.production_ready, false);
assert.equal(fixtureAllowed.warnings.some((item) => item.code === 'fixture_dom_only'), true);

const activeSnapshot = buildMeetingAppFixtureSnapshot('google-meet', {
  observedAtMs,
  state: 'active',
});
const endedSnapshot = buildMeetingAppFixtureSnapshot('google-meet', {
  observedAtMs: observedAtMs + 1_000,
  state: 'prejoin',
});
const capturedGate = buildMeetingAppLaunchGate('google-meet', {
  snapshots: {
    google_meet: [activeSnapshot, endedSnapshot],
  },
});
assert.equal(capturedGate.evidence_level, 'captured_dom');
assert.equal(capturedGate.status, 'passed');
assert.equal(capturedGate.passed, true);
assert.equal(capturedGate.production_ready, true);
assert.deepEqual(capturedGate.missing_required_coverage, []);
assert.equal(capturedGate.coverage.ended_in_meeting_false, true);
assert.equal(capturedGate.reports.captured.signal_types.includes('meeting_ended'), true);

const runtimeDisabled = buildMeetingAppLaunchGate('google-meet', {
  observedAtMs,
  runtimePreset: false,
  allowFixtureProduction: true,
  requireProductionReady: false,
});
assert.equal(runtimeDisabled.status, 'failed');
assert.equal(runtimeDisabled.runtime_ready, false);
assert.equal(runtimeDisabled.missing_required_coverage.includes('runtime_preset'), true);
assert.equal(runtimeDisabled.missing_required_coverage.includes('mutation_observer'), true);

const noEvidence = buildMeetingAppLaunchGate('zoom', {
  allowFixtureEvidence: false,
  requireProductionReady: false,
});
assert.equal(noEvidence.evidence_level, 'none');
assert.equal(noEvidence.status, 'failed');
assert.equal(noEvidence.coverage.platform_detected, false);
assert.equal(noEvidence.coverage.meeting_started, false);
assert.equal(noEvidence.blocking_issues.some((item) => item.code === 'missing_meeting_app_evidence'), true);

const summary = buildMeetingAppLaunchGateSummary({
  platforms: MEETING_APP_FIXTURE_PLATFORMS,
  allowFixtureProduction: true,
  requireProductionReady: false,
});
assert.equal(summary.type, 'meeting_app_launch_gate_summary');
assert.equal(summary.ok, true);
assert.equal(summary.production_ready, false);
assert.equal(summary.passed_count, MEETING_APP_FIXTURE_PLATFORMS.length);
assert.equal(summary.warning_count, MEETING_APP_FIXTURE_PLATFORMS.length);

const assertedSummary = assertAllMeetingAppLaunchGates({
  platforms: ['google_meet', 'webex'],
  allowFixtureProduction: true,
  requireProductionReady: false,
});
assert.equal(assertedSummary.ok, true);
assert.equal(assertedSummary.gates.length, 2);

assert.throws(
  () => assertMeetingAppLaunchGate('google-meet', { observedAtMs }),
  /Meeting app launch gate failed/,
);

assert.throws(
  () => buildMeetingAppLaunchGate('local-detector'),
  /Unsupported meeting app launch gate platform/,
);

console.log('ok meeting app launch gate');
