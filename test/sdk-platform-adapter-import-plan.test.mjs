import assert from 'node:assert/strict';

import {
  buildMeetingPlatformAdapterExportPackage,
  buildMeetingPlatformAdapterExportPackageMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-export-package.mjs';
import {
  assertMeetingPlatformAdapterImportPlan,
  buildMeetingPlatformAdapterImportPlan,
  buildMeetingPlatformAdapterImportPlanMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-import-plan.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import {
  createMeetingAppTimelineSdk,
} from '../packages/meeting-timeline-sdk/index.mjs';

const baseUrl = 'https://timeline.example.com';
const googleExport = buildMeetingPlatformAdapterExportPackage('google-meet', {}, {
  baseUrl,
  target: 'static',
});
const googleFiles = googleExport.host_files.map((file) => file.path);

const googlePlan = buildMeetingPlatformAdapterImportPlan(googleExport, {
  availableFiles: googleFiles,
  target: 'static',
});
assert.equal(googlePlan.schema, 'meeting_platform_adapter_import_plan');
assert.equal(googlePlan.platform, 'google_meet');
assert.equal(googlePlan.accepted, true);
assert.equal(googlePlan.selected_surface, 'browser_extension');
assert.equal(googlePlan.runtime_contract.timestamp_field, 'captured_at_ms');
assert.equal(googlePlan.runtime_contract.local_axis_first, true);
assert.equal(googlePlan.runtime_contract.provider_events_block_realtime, false);
assert.equal(googlePlan.runtime_contract.transcript_blocks_realtime, false);
assert.equal(googlePlan.runtime_contract.adapter_preflight_required_before_realtime_insert, true);
assert.equal(googlePlan.runtime_contract.adapter_preflight_url_only_status, 'needs_live_page_evidence');
assert.equal(googlePlan.adapter_blueprint.available, true);
assert.equal(googlePlan.adapter_blueprint.path, 'google_meet/adapter-blueprint.json');
assert.equal(googlePlan.adapter_blueprint.schema, 'meeting_platform_adapter_blueprint');
assert.equal(googlePlan.adapter_blueprint.command.includes('meeting-platform:adapter-blueprint'), true);
assert.equal(googlePlan.adapter_preflight.status, 'needs_live_page_evidence');
assert.equal(googlePlan.adapter_preflight.selected_surface, 'browser_extension');
assert.equal(googlePlan.adapter_preflight.startup_ready, true);
assert.equal(googlePlan.adapter_preflight.live_evidence_required, true);
assert.equal(googlePlan.adapter_preflight.realtime_annotation_ready, false);
assert.equal(googlePlan.adapter_preflight.command.includes('meeting-platform:adapter-preflight'), true);
assert.equal(googlePlan.adapter_preflight.bridge_messages.includes('meeting_timeline.preflight_candidates'), true);
assert.equal(googlePlan.host_file_coverage.status, 'complete');
assert.equal(googlePlan.readiness.hard_contract_ready, true);
assert.equal(googlePlan.readiness.adapter_preflight_contract_ready, true);
assert.equal(googlePlan.readiness.selected_surface_ready, true);
assert.equal(googlePlan.install_steps.find((step) => step.id === 'run_adapter_preflight').url_only_status, 'needs_live_page_evidence');
assert.equal(googlePlan.install_steps.find((step) => step.id === 'bind_current_axis').sdk_method, 'observePlatformCandidates');
assert.equal(googlePlan.install_steps.find((step) => step.id === 'insert_realtime_marks').action.includes("insertAnnotation('google_meet'"), true);
assert.equal(googlePlan.sdk_imports.adapter_blueprint, '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-blueprint');
assert.equal(googlePlan.sdk_imports.adapter_preflight, '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-preflight');
assert.equal(googlePlan.sdk_imports.adapter_import_plan, '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-import-plan');
assert.equal(assertMeetingPlatformAdapterImportPlan(googlePlan), googlePlan);

const googleExportWithArtifacts = buildMeetingPlatformAdapterExportPackage('google-meet', {}, {
  baseUrl,
  target: 'static',
  includeArtifacts: true,
});
const googlePlanWithBlueprint = buildMeetingPlatformAdapterImportPlan(googleExportWithArtifacts, {
  availableFiles: googleExportWithArtifacts.host_files.map((file) => file.path),
  target: 'static',
});
assert.equal(googlePlanWithBlueprint.adapter_blueprint.primary_surface, 'browser_extension');
assert.equal(googlePlanWithBlueprint.adapter_blueprint.provider_blocks_realtime, false);
assert.equal(googlePlanWithBlueprint.adapter_blueprint.first_acceptance_gate, 'local_candidate_preflight_accepts_active_meeting');

const missingFilesPlan = buildMeetingPlatformAdapterImportPlan(googleExport, {
  availableFiles: ['google_meet/adapter-export-package.json'],
  target: 'static',
});
assert.equal(missingFilesPlan.accepted, false);
assert.equal(missingFilesPlan.host_file_coverage.status, 'missing_required_files');
assert.equal(missingFilesPlan.readiness.issues.some((item) => item.code === 'missing_required_host_files'), true);

const webviewPlan = buildMeetingPlatformAdapterImportPlan(googleExport, {
  availableFiles: googleFiles,
  surface: 'webview-preload',
});
assert.equal(webviewPlan.selected_surface, 'webview_preload');
assert.equal(webviewPlan.accepted, true);

const customExport = buildMeetingPlatformAdapterExportPackage('Acme Rooms', {}, {
  baseUrl,
  target: 'static',
  displayName: 'Acme Rooms',
  browserMatches: ['https://meet.acme.example/*'],
  providerPath: 'acme_rooms_webhooks',
});
const customBlocked = buildMeetingPlatformAdapterImportPlan(customExport, {
  availableFiles: customExport.host_files.map((file) => file.path),
});
assert.equal(customBlocked.accepted, false);
assert.equal(customBlocked.readiness.issues.some((item) => item.code === 'export_package_not_accepted'), true);
const customAllowed = buildMeetingPlatformAdapterImportPlan(customExport, {
  availableFiles: customExport.host_files.map((file) => file.path),
  allowCustomAuthoring: true,
});
assert.equal(customAllowed.accepted, true);
assert.equal(customAllowed.built_in, false);

const matrixExport = buildMeetingPlatformAdapterExportPackageMatrix({
  platforms: ['google-meet', 'zoom'],
}, {
  baseUrl,
  target: 'static',
});
const matrixFiles = matrixExport.packages.flatMap((pkg) => pkg.host_files.map((file) => file.path));
const matrix = buildMeetingPlatformAdapterImportPlanMatrix(matrixExport.packages, {
  availableFiles: matrixFiles,
  target: 'static',
});
assert.equal(matrix.schema, 'meeting_platform_adapter_import_plan_matrix');
assert.equal(matrix.package_count, 2);
assert.equal(matrix.accepted_count, 2);
assert.equal(matrix.missing_file_count, 0);
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').selected_surface, 'native_detector');
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').adapter_blueprint_available, true);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').adapter_preflight_status, 'needs_live_page_evidence');
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').adapter_preflight_selected_surface, 'browser_extension');
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').adapter_preflight_selected_surface, 'native_detector');
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').adapter_preflight_startup_ready, true);
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').adapter_preflight_realtime_ready, false);

const client = {
  async startMeeting(input) {
    return { ok: true, input };
  },
  async endMeeting(input) {
    return { ok: true, input };
  },
  async insertMark(input) {
    return { ok: true, input };
  },
};
const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(kit.platformAdapterImportPlan(googleExport, { availableFiles: googleFiles }).accepted, true);
assert.equal(kit.platformAdapterImportPlanMatrix(matrixExport.packages, { availableFiles: matrixFiles }).accepted_count, 2);
assert.equal(kit.assertPlatformAdapterImportPlan(googlePlan), googlePlan);
assert.equal(kit.report().platform_adapter_import_plan_matrix.package_count, 2);

const sdk = createMeetingAppTimelineSdk({
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(sdk.platformAdapterImportPlan(googleExport, { availableFiles: googleFiles }).accepted, true);
assert.equal(sdk.adapterImportPlan(googleExport, { availableFiles: googleFiles }).accepted, true);
assert.equal(sdk.platformAdapterImportPlanMatrix(matrixExport.packages, { availableFiles: matrixFiles }).accepted_count, 2);
assert.equal(sdk.adapterImportPlanMatrix(matrixExport.packages, { availableFiles: matrixFiles }).blocked_count, 0);

console.log('ok meeting platform adapter import plan');
