import assert from 'node:assert/strict';

import {
  buildMeetingPlatformAdapterExportPackage,
  buildMeetingPlatformAdapterExportPackageMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-export-package.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import {
  createMeetingAppTimelineSdk,
} from '../packages/meeting-timeline-sdk/index.mjs';

const baseUrl = 'https://timeline.example.com';

const google = buildMeetingPlatformAdapterExportPackage('google-meet', {}, {
  baseUrl,
  target: 'static',
  includeArtifacts: true,
});
assert.equal(google.schema, 'meeting_platform_adapter_export_package');
assert.equal(google.platform, 'google_meet');
assert.equal(google.display_name, 'Google Meet');
assert.equal(google.built_in, true);
assert.equal(google.accepted, true);
assert.equal(google.export_ready, true);
assert.equal(google.timestamp_field, 'captured_at_ms');
assert.equal(google.provider_events_block_realtime, false);
assert.equal(google.transcript_blocks_realtime, false);
assert.equal(google.import_paths.platform_kit, '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit');
assert.equal(google.import_paths.adapter_export_package, '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-export-package');
assert.equal(google.import_paths.adapter_blueprint, '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-blueprint');
assert.equal(google.import_paths.adapter_preflight, '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-preflight');
assert.equal(google.commands.adapter_preflight.includes('meeting-platform:adapter-preflight'), true);
assert.equal(google.adapter_preflight.status, 'needs_live_page_evidence');
assert.equal(google.adapter_preflight.selected_surface, 'browser_extension');
assert.equal(google.adapter_preflight.startup_ready, true);
assert.equal(google.adapter_preflight.live_evidence_required, true);
assert.equal(google.adapter_preflight.realtime_annotation_ready, false);
assert.equal(google.adapter_preflight.bridge_messages.includes('meeting_timeline.preflight_current_window'), true);
assert.equal(google.surface_entrypoints.browser_extension.ready, true);
assert.equal(google.surface_entrypoints.browser_extension.first_message_type, 'meeting_timeline.observe_candidates');
assert.equal(google.surface_entrypoints.browser_extension.preflight_required_before_insert, true);
assert.equal(google.surface_entrypoints.browser_extension.preflight_message_types.includes('meeting_timeline.preflight_candidates'), true);
assert.equal(google.surface_entrypoints.browser_extension.mark_insert_method, 'insertAnnotation');
assert.equal(google.surface_entrypoints.provider_reconcile.ready, true);
assert.equal(google.surface_entrypoints.provider_reconcile.required_for_realtime, false);
assert.equal(google.setup_order.find((step) => step.id === 'run_adapter_preflight').url_only_status, 'needs_live_page_evidence');
assert.equal(google.setup_order.find((step) => step.id === 'insert_realtime_marks').action.includes("insertAnnotation('google_meet'"), true);
assert.equal(google.host_files.find((file) => file.source === 'runtime_bundle').path, 'google_meet/runtime-bundle.json');
assert.equal(google.host_files.find((file) => file.source === 'adapter_blueprint').path, 'google_meet/adapter-blueprint.json');
assert.equal(google.host_files.find((file) => file.source === 'provider_connection').required_from, 'production');
assert.equal(google.artifact_refs.runtime_bundle.path, 'google_meet/runtime-bundle.json');
assert.equal(google.artifact_refs.adapter_blueprint.path, 'google_meet/adapter-blueprint.json');
assert.equal(google.artifacts.runtime_bundle.schema, 'meeting_platform_runtime_bundle');
assert.equal(google.artifacts.adapter_blueprint.schema, 'meeting_platform_adapter_blueprint');
assert.equal(google.artifacts.adapter_blueprint.primary_surface, 'browser_extension');
assert.equal(google.artifacts.provider_connection.event_mapping.length > 0, true);
assert.equal(google.commands.export_package.includes('meeting-platform:adapter-export-package'), true);
assert.equal(google.commands.adapter_blueprint.includes('meeting-platform:adapter-blueprint'), true);
assert.equal(google.readiness.static_accepted, true);
assert.equal(google.readiness.adapter_preflight_startup_ready, true);
assert.equal(google.readiness.adapter_preflight_realtime_ready, false);
assert.deepEqual(google.readiness.failed_required_ids, []);

const custom = buildMeetingPlatformAdapterExportPackage('Acme Rooms', {}, {
  baseUrl,
  target: 'static',
  displayName: 'Acme Rooms',
  browserMatches: ['https://meet.acme.example/*'],
  providerPath: 'acme_rooms_webhooks',
});
assert.equal(custom.platform, 'acme_rooms');
assert.equal(custom.built_in, false);
assert.equal(custom.accepted, false);
assert.equal(custom.export_ready, false);
assert.equal(custom.host_files.some((file) => file.source === 'runtime_bundle'), false);
assert.equal(custom.host_files.find((file) => file.source === 'authoring_plan').path, 'acme_rooms/authoring-plan.json');
assert.equal(custom.next_actions.includes('author_missing_platform_adapter_before_runtime_install'), true);

const matrix = buildMeetingPlatformAdapterExportPackageMatrix({
  platforms: ['google-meet', 'teams', 'Acme Rooms'],
}, {
  baseUrl,
  target: 'static',
});
assert.equal(matrix.schema, 'meeting_platform_adapter_export_package_matrix');
assert.equal(matrix.platform_count, 3);
assert.equal(matrix.accepted_count, 2);
assert.equal(matrix.export_ready_count, 2);
assert.equal(matrix.built_in_count, 2);
assert.equal(matrix.custom_authoring_count, 1);
assert.equal(matrix.browser_extension_ready_count, 2);
assert.equal(matrix.provider_reconcile_count, 3);
assert.equal(matrix.adapter_preflight_startup_ready_count, 2);
assert.equal(matrix.adapter_preflight_realtime_ready_count, 0);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').host_file_count >= 8, true);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').adapter_preflight_status, 'needs_live_page_evidence');
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').adapter_preflight_selected_surface, 'browser_extension');
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').adapter_preflight_startup_ready, true);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').adapter_preflight_realtime_ready, false);
assert.equal(matrix.rows.find((row) => row.platform === 'acme_rooms').accepted, false);
assert.equal(matrix.acceptance_checklist_matrix.accepted_count, 2);

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
assert.equal(kit.platformAdapterExportPackage('zoom', {}, { target: 'static' }).accepted, true);
assert.equal(kit.platformAdapterExportPackage('zoom', {}, { target: 'static' }).recommended_first_surface, 'native_detector');
assert.equal(kit.platformAdapterExportPackage('zoom', {}, { target: 'static' }).surface_entrypoints.native_detector.ready, true);
assert.equal(kit.platformAdapterExportPackageMatrix({}, { target: 'static' }).platform_count, 2);

const sdk = createMeetingAppTimelineSdk({
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(sdk.platformAdapterExportPackage('google-meet', {}, { target: 'static' }).accepted, true);
assert.equal(sdk.adapterExportPackage('zoom', {}, { target: 'static' }).accepted, true);
assert.equal(sdk.platformAdapterExportPackageMatrix({}, { target: 'static' }).accepted_count, 2);
assert.equal(sdk.adapterExportPackageMatrix({ platforms: ['google-meet', 'Acme Rooms'] }, { target: 'static' }).custom_authoring_count, 1);

console.log('ok meeting platform adapter export package');
