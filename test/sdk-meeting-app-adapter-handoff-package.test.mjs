import assert from 'node:assert/strict';

import {
  MEETING_APP_ADAPTER_HANDOFF_PACKAGE_MATRIX_SCHEMA,
  MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA,
  MEETING_APP_ADAPTER_VERIFICATION_PLAN_SCHEMA,
  MEETING_APP_ADAPTER_VERIFICATION_REPORT_MATRIX_SCHEMA,
  MEETING_APP_ADAPTER_VERIFICATION_REPORT_SCHEMA,
  assertMeetingAppAdapterHandoffPackage,
  assertMeetingAppAdapterHandoffPackageMatrix,
  assertMeetingAppAdapterVerificationReport,
  assertMeetingAppAdapterVerificationReportMatrix,
  buildMeetingAppAdapterHandoffPackage,
  buildMeetingAppAdapterHandoffPackageMatrix,
  buildMeetingAppAdapterVerificationReport,
  buildMeetingAppAdapterVerificationReportMatrix,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-adapter-handoff-package.mjs';
import { buildMeetingAppAdapterSpec } from '../packages/meeting-timeline-sdk/adapters/meeting-app-adapter-spec.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const wherebySpec = {
  adapter_key: 'whereby',
  display_name: 'Whereby',
  matches: ['https://whereby.com/*'],
  host_permissions: ['https://whereby.com/*'],
  control_selectors: ['[aria-label*="Leave" i]', '[data-testid*="toolbar" i]'],
  participant_selectors: ['[data-participant-id]', '[aria-label*="speaking" i]'],
  text_selectors: ['[role="status"]', '[aria-live]'],
  mutation_track_selectors: ['[data-participant-id]', '[role="status"]'],
};

const whereby = buildMeetingAppAdapterHandoffPackage(wherebySpec, {
  windowMessaging: true,
  allowedOrigins: ['https://whereby.com'],
});
assert.equal(whereby.schema, MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA);
assert.equal(whereby.accepted, true);
assert.equal(whereby.adapter_key, 'whereby');
assert.equal(whereby.adapter_spec.accepted, true);
assert.equal(whereby.runtime_config.accepted, true);
assert.equal(whereby.adapter_manifest, undefined);
assert.equal(whereby.file_paths.includes('adapter-spec.json'), true);
assert.equal(whereby.file_paths.includes('runtime-config.json'), true);
assert.equal(whereby.file_paths.includes('extension-manifest-fragment.json'), true);
assert.equal(whereby.file_paths.includes('verification-plan.json'), true);
assert.equal(whereby.file_paths.includes('integration-readme.md'), true);
assert.equal(whereby.contracts.timestamp_field, 'captured_at_ms');
assert.equal(whereby.contracts.provider_events_block_realtime, false);
assert.equal(whereby.validation.required_live_evidence.includes('annotation_insert_current_axis'), true);
assert.equal(whereby.verification_plan.schema, MEETING_APP_ADAPTER_VERIFICATION_PLAN_SCHEMA);
assert.equal(whereby.verification_plan.acceptance_policy.timestamp_field, 'captured_at_ms');
assert.equal(whereby.verification_plan.required_evidence.some((item) => item.id === 'speaker_track'), true);
assert.equal(whereby.verification_plan.acceptance_checks.some((item) => item.id === 'verify_annotation_insert_current_axis'), true);
assert.equal(
  String(whereby.files.find((file) => file.path === 'integration-readme.md').content).includes('installMeetingAppContentScriptBridge'),
  true,
);
assert.equal(
  whereby.files.find((file) => file.path === 'verification-plan.json').content.required_evidence.length,
  5,
);
assert.equal(assertMeetingAppAdapterHandoffPackage(whereby).adapter_key, 'whereby');

const completeEvidence = {
  live_dom_snapshot_count: 1,
  candidate_observation_count: 1,
  speaker_segments: [{ captured_at_ms: 1_782_614_401_000, speaker_label: 'Ada' }],
  participant_segments: [{ captured_at_ms: 1_782_614_401_000, participant_label: 'Ada' }],
  annotations: [{
    captured_at_ms: 1_782_614_402_000,
    meeting_id: 'whereby-demo',
    axis_id: 'axis-1',
  }],
};
const missingEvidenceReport = buildMeetingAppAdapterVerificationReport(whereby);
assert.equal(missingEvidenceReport.schema, MEETING_APP_ADAPTER_VERIFICATION_REPORT_SCHEMA);
assert.equal(missingEvidenceReport.accepted, false);
assert.equal(missingEvidenceReport.static_ready, true);
assert.equal(missingEvidenceReport.live_evidence_ready, false);
assert.equal(missingEvidenceReport.missing_evidence_count, 5);
assert.equal(missingEvidenceReport.next_actions.includes('verify_live_dom_snapshot'), true);
assert.throws(() => assertMeetingAppAdapterVerificationReport(missingEvidenceReport), /Meeting app adapter verification failed/);

const verifiedWhereby = buildMeetingAppAdapterVerificationReport(whereby, {
  evidence: completeEvidence,
});
assert.equal(verifiedWhereby.accepted, true);
assert.equal(verifiedWhereby.live_evidence_ready, true);
assert.equal(verifiedWhereby.pilot_ready, true);
assert.equal(verifiedWhereby.production_ready, true);
assert.equal(verifiedWhereby.evidence_summary.annotation_insert_current_axis, 1);
assert.equal(assertMeetingAppAdapterVerificationReport(verifiedWhereby).adapter_key, 'whereby');

const google = buildMeetingAppAdapterHandoffPackage('google-meet');
assert.equal(google.accepted, true);
assert.equal(google.source, 'built_in_manifest');
assert.equal(google.adapter_manifest.schema, 'meeting_app_adapter_manifest');
assert.equal(google.file_paths.includes('adapter-manifest.json'), true);
assert.equal(google.extension_manifest_fragment.matches.includes('https://meet.google.com/*'), true);

const brokenSpec = buildMeetingAppAdapterSpec({
  adapter_key: 'broken_app',
  matches: ['https://broken.example/*'],
});
const broken = buildMeetingAppAdapterHandoffPackage(brokenSpec);
assert.equal(broken.accepted, false);
assert.equal(broken.issues.some((item) => item.code === 'adapter_spec_not_accepted'), true);
assert.throws(() => assertMeetingAppAdapterHandoffPackage(broken), /Meeting app adapter handoff package acceptance failed/);

const matrix = buildMeetingAppAdapterHandoffPackageMatrix({
  platforms: ['google-meet', 'zoom'],
  adapters: [wherebySpec],
});
assert.equal(matrix.schema, MEETING_APP_ADAPTER_HANDOFF_PACKAGE_MATRIX_SCHEMA);
assert.equal(matrix.accepted, true);
assert.equal(matrix.package_count, 3);
assert.equal(matrix.accepted_count, 3);
assert.equal(matrix.built_in_count, 2);
assert.equal(matrix.custom_count, 1);
assert.equal(matrix.runtime_config_ready_count, 3);
assert.equal(matrix.content_script_ready_count, 3);
assert.equal(matrix.live_evidence_required_count, 3);
assert.equal(assertMeetingAppAdapterHandoffPackageMatrix(matrix).accepted, true);

const verificationMatrix = buildMeetingAppAdapterVerificationReportMatrix({
  platforms: ['google-meet'],
  adapters: [wherebySpec],
  evidenceByAdapter: {
    google_meet: completeEvidence,
    whereby: completeEvidence,
  },
});
assert.equal(verificationMatrix.schema, MEETING_APP_ADAPTER_VERIFICATION_REPORT_MATRIX_SCHEMA);
assert.equal(verificationMatrix.accepted, true);
assert.equal(verificationMatrix.report_count, 2);
assert.equal(verificationMatrix.production_ready_count, 2);
assert.equal(assertMeetingAppAdapterVerificationReportMatrix(verificationMatrix).accepted_count, 2);

const kit = createMeetingPlatformTimelineKit({ baseUrl: 'http://localhost:8787' });
assert.equal(kit.meetingAppAdapterHandoffPackage(wherebySpec).adapter_key, 'whereby');
assert.equal(kit.meetingAppAdapterHandoffPackageMatrix({ platforms: ['google-meet'], adapters: [wherebySpec] }).package_count, 2);
assert.equal(kit.assertMeetingAppAdapterHandoffPackage(wherebySpec).accepted, true);
assert.equal(kit.assertMeetingAppAdapterHandoffPackageMatrix({ platforms: ['zoom'], adapters: [wherebySpec] }).accepted, true);
assert.equal(kit.report({ platforms: ['google-meet'] }).meeting_app_adapter_handoff_package_matrix.package_count, 1);
assert.equal(kit.meetingAppAdapterVerificationReport(wherebySpec, { evidence: completeEvidence }).accepted, true);
assert.equal(kit.meetingAppAdapterVerificationReportMatrix({
  platforms: ['google-meet'],
  evidenceByAdapter: { google_meet: completeEvidence },
}).schema, MEETING_APP_ADAPTER_VERIFICATION_REPORT_MATRIX_SCHEMA);
assert.equal(kit.assertMeetingAppAdapterVerificationReport(wherebySpec, { evidence: completeEvidence }).accepted, true);
assert.equal(kit.assertMeetingAppAdapterVerificationReportMatrix({
  platforms: ['google-meet'],
  evidenceByAdapter: { google_meet: completeEvidence },
}).accepted, true);
assert.equal(kit.report({
  platforms: ['google-meet'],
  evidenceByAdapter: { google_meet: completeEvidence },
}).meeting_app_adapter_verification_report_matrix.production_ready_count, 1);

console.log('ok meeting app adapter handoff package');
