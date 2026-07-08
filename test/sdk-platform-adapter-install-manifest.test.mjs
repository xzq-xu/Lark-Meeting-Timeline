import assert from 'node:assert/strict';

import {
  buildMeetingPlatformAdapterExportPackage,
  buildMeetingPlatformAdapterExportPackageMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-export-package.mjs';
import {
  buildMeetingPlatformAdapterImportPlan,
  buildMeetingPlatformAdapterImportPlanMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-import-plan.mjs';
import {
  assertMeetingPlatformAdapterInstallManifest,
  buildMeetingPlatformAdapterInstallManifest,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-install-manifest.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import {
  createMeetingAppTimelineSdk,
} from '../packages/meeting-timeline-sdk/index.mjs';

const baseUrl = 'https://timeline.example.com';
const exportMatrix = buildMeetingPlatformAdapterExportPackageMatrix({
  platforms: ['google-meet', 'zoom'],
}, {
  baseUrl,
  target: 'static',
  includeArtifacts: true,
});
const availableFiles = exportMatrix.packages.flatMap((pkg) => pkg.host_files.map((file) => file.path));
const importMatrix = buildMeetingPlatformAdapterImportPlanMatrix(exportMatrix.packages, {
  availableFiles,
  target: 'static',
});

const manifest = buildMeetingPlatformAdapterInstallManifest(importMatrix.plans, {
  baseUrl,
  installTarget: 'desktop_host',
});
assert.equal(manifest.schema, 'meeting_platform_adapter_install_manifest');
assert.equal(manifest.accepted, true);
assert.equal(manifest.platform_count, 2);
assert.deepEqual(manifest.platforms, ['google_meet', 'zoom']);
assert.deepEqual(manifest.selected_surfaces, ['browser_extension', 'native_detector']);
assert.equal(manifest.sdk_imports.adapter_blueprint, '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-blueprint');
assert.equal(manifest.sdk_imports.adapter_preflight, '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-preflight');
assert.equal(manifest.sdk_imports.raw_signal, '@ai-annotation/meeting-timeline-sdk/adapters/platform-raw-signal');
assert.equal(manifest.runtime_contract.timestamp_field, 'captured_at_ms');
assert.equal(manifest.runtime_contract.raw_signal_validation_required_before_preflight, true);
assert.equal(manifest.runtime_contract.adapter_preflight_required_before_realtime_insert, true);
assert.equal(manifest.runtime_contract.adapter_preflight_url_only_status, 'needs_live_page_evidence');
assert.equal(manifest.browser_extension.enabled, true);
assert.equal(manifest.browser_extension.content_scripts.length, 1);
assert.equal(manifest.native_detector.enabled, true);
assert.equal(manifest.native_detector.platform_count, 1);
assert.equal(manifest.browser_extension.host_permissions.includes('https://meet.google.com/*'), true);
assert.equal(manifest.browser_extension.message_types.includes('meeting_timeline.preflight_current_window'), true);
assert.equal(manifest.browser_extension.message_types.includes('meeting_timeline.preflight_candidates'), true);
assert.equal(manifest.platform_registry.find((row) => row.platform === 'google_meet').mark_insert_method, 'insertAnnotation');
assert.equal(manifest.platform_registry.find((row) => row.platform === 'google_meet').adapter_blueprint_path, 'google_meet/adapter-blueprint.json');
assert.equal(manifest.platform_registry.find((row) => row.platform === 'google_meet').adapter_blueprint_primary_surface, 'browser_extension');
assert.equal(manifest.platform_registry.find((row) => row.platform === 'google_meet').raw_signal_validation_path, 'google_meet/raw-signal-validation.json');
assert.equal(manifest.platform_registry.find((row) => row.platform === 'google_meet').raw_signal_validation_status, 'ready');
assert.equal(manifest.platform_registry.find((row) => row.platform === 'google_meet').raw_signal_validation_runtime_actions.includes('insert_annotation'), true);
assert.equal(manifest.platform_registry.find((row) => row.platform === 'google_meet').adapter_preflight_status, 'needs_live_page_evidence');
assert.equal(manifest.platform_registry.find((row) => row.platform === 'google_meet').adapter_preflight_selected_surface, 'browser_extension');
assert.equal(manifest.platform_registry.find((row) => row.platform === 'google_meet').adapter_preflight_startup_ready, true);
assert.equal(manifest.platform_registry.find((row) => row.platform === 'google_meet').adapter_preflight_realtime_ready, false);
assert.equal(manifest.adapter_blueprints.enabled, true);
assert.equal(manifest.adapter_blueprints.platform_count, 2);
assert.equal(manifest.adapter_blueprints.rows.find((row) => row.platform === 'google_meet').primary_surface, 'browser_extension');
assert.equal(manifest.raw_signal_validation.enabled, true);
assert.equal(manifest.raw_signal_validation.platform_count, 2);
assert.equal(manifest.raw_signal_validation.ready_count, 2);
assert.equal(manifest.raw_signal_validation.rows.find((row) => row.platform === 'google_meet').status, 'ready');
assert.equal(manifest.adapter_preflight.enabled, true);
assert.equal(manifest.adapter_preflight.platform_count, 2);
assert.equal(manifest.adapter_preflight.startup_ready_count, 2);
assert.equal(manifest.adapter_preflight.realtime_ready_count, 0);
assert.equal(manifest.adapter_preflight.rows.find((row) => row.platform === 'zoom').selected_surface, 'native_detector');
assert.equal(manifest.install_sequence.find((step) => step.id === 'load_adapter_blueprints').required, false);
assert.equal(manifest.install_sequence.find((step) => step.id === 'run_raw_signal_validation_before_preflight').sdk_method, 'platformRawSignalBatch');
assert.equal(manifest.install_sequence.find((step) => step.id === 'run_adapter_preflight_before_realtime_session').url_only_status, 'needs_live_page_evidence');
assert.equal(manifest.install_sequence.find((step) => step.id === 'bind_axis_before_marks').sdk_method, 'observePlatformCandidates');
assert.equal(manifest.provider_reconcile.platform_count, 2);
assert.equal(manifest.provider_reconcile.rows.find((row) => row.platform === 'google_meet').provider_path, 'google_workspace_events_pubsub');
assert.equal(manifest.provider_reconcile.realtime_blocking_count, 0);
assert.equal(assertMeetingPlatformAdapterInstallManifest(manifest), manifest);

const fromPackages = buildMeetingPlatformAdapterInstallManifest({
  packages: exportMatrix.packages,
}, {
  availableFiles,
  baseUrl,
});
assert.equal(fromPackages.accepted, true);
assert.equal(fromPackages.platform_count, 2);

const blockedPlan = buildMeetingPlatformAdapterImportPlan(exportMatrix.packages[0], {
  availableFiles: ['google_meet/adapter-export-package.json'],
});
const blockedManifest = buildMeetingPlatformAdapterInstallManifest([blockedPlan]);
assert.equal(blockedManifest.accepted, false);
assert.equal(blockedManifest.readiness.issues.some((issue) => issue.code === 'blocked_import_plans'), true);
assert.throws(() => assertMeetingPlatformAdapterInstallManifest(blockedManifest), /install manifest is not ready/);

const duplicateManifest = buildMeetingPlatformAdapterInstallManifest([
  importMatrix.plans[0],
  importMatrix.plans[0],
]);
assert.equal(duplicateManifest.accepted, false);
assert.equal(duplicateManifest.readiness.issues.some((issue) => issue.code === 'duplicate_platforms'), true);

const webviewExport = buildMeetingPlatformAdapterExportPackage('google-meet', {}, {
  baseUrl,
  target: 'static',
});
const webviewPlan = buildMeetingPlatformAdapterImportPlan(webviewExport, {
  availableFiles: webviewExport.host_files.map((file) => file.path),
  surface: 'webview-preload',
});
const webviewManifest = buildMeetingPlatformAdapterInstallManifest([webviewPlan]);
assert.equal(webviewManifest.accepted, true);
assert.equal(webviewManifest.browser_extension.enabled, false);
assert.equal(webviewManifest.webview_preload.enabled, true);
assert.equal(webviewManifest.webview_preload.platform_count, 1);

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
assert.equal(kit.platformAdapterInstallManifest(importMatrix.plans).accepted, true);
assert.equal(kit.assertPlatformAdapterInstallManifest(manifest), manifest);
assert.equal(kit.report().platform_adapter_install_manifest.platform_count, 2);

const sdk = createMeetingAppTimelineSdk({
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(sdk.platformAdapterInstallManifest(importMatrix.plans).accepted, true);
assert.equal(sdk.adapterInstallManifest(importMatrix.plans).browser_extension.platform_count, 1);
assert.equal(sdk.adapterInstallManifest(importMatrix.plans).native_detector.platform_count, 1);
assert.equal(sdk.assertPlatformAdapterInstallManifest(manifest), manifest);

console.log('ok meeting platform adapter install manifest');
