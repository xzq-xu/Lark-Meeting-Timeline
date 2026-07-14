import assert from 'node:assert/strict';

import {
  assertMeetingPlatformAdapterRuntimeManifest,
  assertMeetingPlatformAdapterRuntimeRecipe,
  assertMeetingPlatformAdapterRuntimeRecipeMatrix,
  assertMeetingPlatformAdapterRuntimeTarget,
  buildMeetingPlatformAdapterRuntimeManifest,
  buildMeetingPlatformAdapterRuntimeRecipe,
  buildMeetingPlatformAdapterRuntimeRecipeMatrix,
  buildMeetingPlatformAdapterRuntimeTarget,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-runtime-recipe.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import {
  createMeetingAppTimelineSdk,
} from '../packages/meeting-timeline-sdk/index.mjs';

const baseUrl = 'https://timeline.example.com';

const google = buildMeetingPlatformAdapterRuntimeRecipe({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Google Meet',
}, {
  baseUrl,
});
assert.equal(google.schema, 'meeting_platform_adapter_runtime_recipe');
assert.equal(google.accepted, true);
assert.equal(google.runtime_ready, true);
assert.equal(google.status, 'ready_to_wire_runtime');
assert.equal(google.platform, 'google_meet');
assert.equal(google.selected_surface, 'browser_extension');
assert.equal(google.host_wiring.bridge_kind, 'browser_content_script');
assert.equal(google.host_wiring.first_required_method, 'observePlatformCandidates');
assert.equal(google.host_wiring.insert_method, 'insertAnnotation');
assert.equal(google.host_wiring.optional_track_methods.includes('speakerTrack'), true);
assert.equal(google.runtime_contract.provider_events_block_realtime, false);
assert.equal(google.runtime_contract.transcript_blocks_realtime, false);
assert.equal(google.sequence.find((step) => step.action === 'insert_realtime_annotation').timestamp_field, 'captured_at_ms');
assert.equal(google.sequence.find((step) => step.action === 'speaker_position_markers').text_required, false);
assert.equal(google.sequence.find((step) => step.action === 'provider_reconcile').blocks_realtime_annotation, false);
assert.equal(google.raw_signal_examples.runtime_event_count, 4);
assert.equal(google.raw_signal_examples.filtered_speaker_event_count, 1);
assert.equal(google.host_wiring.runtime_event_actions.includes('observe_meeting_app'), true);
assert.equal(google.host_wiring.runtime_event_actions.includes('insert_annotation'), true);
assert.equal(google.host_wiring.runtime_event_actions.includes('speaker_track'), true);
assert.equal(google.readiness.sample_speaker_track_ready, true);
assert.equal(google.next_actions.includes('wire_recipe_sequence_into_host_runtime'), true);
assert.equal(assertMeetingPlatformAdapterRuntimeRecipe({
  url: 'https://meet.google.com/abc-defg-hij',
}, {
  baseUrl,
}).platform, 'google_meet');

const zoom = buildMeetingPlatformAdapterRuntimeRecipe({
  platform: 'zoom',
  window: { title: 'Zoom Meeting' },
  process: { name: 'Zoom Workplace' },
}, {
  baseUrl,
});
assert.equal(zoom.accepted, true);
assert.equal(zoom.runtime_ready, true);
assert.equal(zoom.selected_surface, 'native_detector');
assert.equal(zoom.install_target, 'native_or_desktop_observer');
assert.equal(zoom.host_wiring.bridge_kind, 'native_detector_runtime_event_client');
assert.equal(zoom.host_wiring.runtime_event_action_counts.speaker_track, 2);

const larkProvider = buildMeetingPlatformAdapterRuntimeRecipe({
  platform: 'lark',
  provider_event: { type: 'vc.meeting.all_meeting_started_v1' },
}, {
  baseUrl,
  hostProfile: 'provider_reconcile_only',
});
assert.equal(larkProvider.accepted, false);
assert.equal(larkProvider.runtime_ready, false);
assert.equal(larkProvider.status, 'needs_local_surface_for_realtime_axis');
assert.equal(larkProvider.selected_surface, 'provider_reconcile');
assert.equal(larkProvider.host_wiring.first_required_method, 'add_local_surface_for_realtime_annotations');
assert.equal(larkProvider.next_actions.includes('choose_browser_or_native_host_profile_before_realtime_marks'), true);
assert.throws(
  () => assertMeetingPlatformAdapterRuntimeRecipe({ platform: 'lark' }, { baseUrl, hostProfile: 'provider_reconcile_only' }),
  /Meeting platform adapter runtime recipe is not accepted/,
);

const matrix = buildMeetingPlatformAdapterRuntimeRecipeMatrix({}, {
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom'],
});
assert.equal(matrix.schema, 'meeting_platform_adapter_runtime_recipe_matrix');
assert.equal(matrix.platform_count, 3);
assert.equal(matrix.accepted_count, 3);
assert.equal(matrix.runtime_ready_count, 3);
assert.equal(matrix.browser_surface_count, 1);
assert.equal(matrix.native_surface_count, 2);
assert.equal(matrix.raw_signal_runtime_event_count, 12);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').bridge_kind, 'browser_content_script');
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').bridge_kind, 'native_detector_runtime_event_client');
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').speaker_track_sample_ready, true);
assert.equal(assertMeetingPlatformAdapterRuntimeRecipeMatrix({}, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
}).accepted_count, 2);

const manifest = buildMeetingPlatformAdapterRuntimeManifest({}, {
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom'],
});
assert.equal(manifest.schema, 'meeting_platform_adapter_runtime_manifest');
assert.equal(manifest.accepted, true);
assert.equal(manifest.runtime_ready, true);
assert.equal(manifest.platform_count, 3);
assert.equal(manifest.accepted_count, 3);
assert.equal(manifest.runtime_ready_count, 3);
assert.equal(manifest.local_surface_count, 3);
assert.equal(manifest.browser_surface_count, 1);
assert.equal(manifest.native_surface_count, 2);
assert.equal(manifest.provider_reconcile_surface_count, 0);
assert.equal(manifest.host_endpoints.runtime_events, `${baseUrl}/api/meeting-platform/runtime-events`);
assert.equal(manifest.host_endpoints.insert_annotation, `${baseUrl}/api/annotations`);
assert.equal(manifest.runtime_contract.timestamp_field, 'captured_at_ms');
assert.equal(manifest.runtime_contract.observe_before_insert, true);
assert.equal(manifest.runtime_contract.provider_events_block_realtime, false);
assert.equal(manifest.runtime_contract.transcript_blocks_realtime, false);
assert.equal(manifest.dispatch_policy.browser_content_script, 'use_for_google_meet_and_browser_meeting_surfaces');
assert.equal(manifest.bridge_groups.find((row) => row.bridge_kind === 'browser_content_script').platforms.includes('google_meet'), true);
assert.equal(manifest.bridge_groups.find((row) => row.bridge_kind === 'native_detector_runtime_event_client').platform_count, 2);
assert.equal(manifest.platform_registry.row_count, 3);
assert.equal(manifest.platform_registry.rows.find((row) => row.platform === 'google_meet').host_kind, 'browser_extension_content_script');
assert.equal(manifest.platform_registry.rows.find((row) => row.platform === 'google_meet').adapter_module, '@ai-annotation/meeting-timeline-sdk/adapters/google-meet');
assert.equal(manifest.platform_registry.rows.find((row) => row.platform === 'google_meet').first_required_method, 'observePlatformCandidates');
assert.equal(manifest.platform_registry.rows.find((row) => row.platform === 'google_meet').insert_method, 'insertAnnotation');
assert.equal(manifest.platform_registry.rows.find((row) => row.platform === 'microsoft_teams').host_kind, 'native_desktop_detector');
assert.equal(manifest.platform_registry.rows.find((row) => row.platform === 'microsoft_teams').adapter_module, '@ai-annotation/meeting-timeline-sdk/adapters/microsoft-teams');
assert.equal(manifest.platform_registry.rows.find((row) => row.platform === 'zoom').bridge_kind, 'native_detector_runtime_event_client');
assert.equal(manifest.platform_registry.rows.find((row) => row.platform === 'zoom').speaker_position_markers.enabled, true);
assert.equal(manifest.platform_registry.rows.find((row) => row.platform === 'zoom').speaker_position_markers.text_required, false);
assert.equal(manifest.platform_registry.rows.find((row) => row.platform === 'zoom').provider_reconcile.blocks_realtime_annotation, false);
assert.equal(manifest.matrix_summary.platforms.includes('google_meet'), true);
assert.equal(manifest.recipes, undefined);
assert.equal(assertMeetingPlatformAdapterRuntimeManifest({}, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
}).runtime_ready, true);

const googleTarget = buildMeetingPlatformAdapterRuntimeTarget(manifest, {
  url: 'https://meet.google.com/abc-defg-hij',
  captured_at_ms: 1_782_614_400_000,
});
assert.equal(googleTarget.schema, 'meeting_platform_adapter_runtime_target');
assert.equal(googleTarget.accepted, true);
assert.equal(googleTarget.status, 'ready_to_start_runtime_target');
assert.equal(googleTarget.platform, 'google_meet');
assert.equal(googleTarget.detection_reason, 'meeting_url');
assert.equal(googleTarget.selected_surface, 'browser_extension');
assert.equal(googleTarget.host_kind, 'browser_extension_content_script');
assert.equal(googleTarget.bridge_kind, 'browser_content_script');
assert.equal(googleTarget.first_required_method, 'observePlatformCandidates');
assert.equal(googleTarget.insert_method, 'insertAnnotation');
assert.equal(googleTarget.runtime_actions.find((row) => row.id === 'observe_platform_candidates').timestamp_field, 'captured_at_ms');
assert.equal(googleTarget.runtime_actions.find((row) => row.id === 'insert_realtime_annotation').sdk_method, 'insertAnnotation');
assert.equal(googleTarget.mark_template.captured_at_ms, 1_782_614_400_000);
assert.equal(googleTarget.next_actions.includes('start_browser_extension_content_script'), true);
assert.equal(assertMeetingPlatformAdapterRuntimeTarget(manifest, {
  url: 'https://meet.google.com/abc-defg-hij',
}).accepted, true);

const zoomTarget = buildMeetingPlatformAdapterRuntimeTarget(manifest, {
  platform: 'zoom',
  window: { title: 'Zoom Meeting' },
});
assert.equal(zoomTarget.accepted, true);
assert.equal(zoomTarget.detection_reason, 'explicit_platform');
assert.equal(zoomTarget.host_kind, 'native_desktop_detector');
assert.equal(zoomTarget.bridge_kind, 'native_detector_runtime_event_client');
assert.equal(zoomTarget.runtime_actions.find((row) => row.id === 'emit_speaker_position').text_required, false);

const mismatchTarget = buildMeetingPlatformAdapterRuntimeTarget(manifest, {
  platform: 'zoom',
  surface: 'browser-extension',
});
assert.equal(mismatchTarget.accepted, false);
assert.equal(mismatchTarget.readiness.issues.find((item) => item.code === 'surface_mismatch').requested_surface, 'browser_extension');
assert.throws(
  () => assertMeetingPlatformAdapterRuntimeTarget(manifest, { platform: 'zoom', surface: 'browser-extension' }),
  /Meeting platform adapter runtime target is not ready/,
);

assert.equal(buildMeetingPlatformAdapterRuntimeTarget({
  platform: 'zoom',
}, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
}).host_kind, 'native_desktop_detector');

const browserProfileMatrix = buildMeetingPlatformAdapterRuntimeRecipeMatrix({}, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
  hostProfile: 'browser_extension',
});
assert.equal(browserProfileMatrix.browser_surface_count, 2);
assert.equal(browserProfileMatrix.rows.every((row) => row.host_profile === 'browser_extension'), true);

const kit = createMeetingPlatformTimelineKit({}, {
  baseUrl,
});
assert.equal(kit.platformAdapterRuntimeRecipe({
  url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
}).platform, 'microsoft_teams');
assert.equal(kit.platformAdapterRuntimeRecipeMatrix({}, {
  platforms: ['webex', 'lark'],
}).platform_count, 2);
assert.equal(kit.platformAdapterRuntimeManifest({}, {
  platforms: ['google-meet', 'zoom'],
}).bridge_groups.find((row) => row.bridge_kind === 'native_detector_runtime_event_client').platforms.includes('zoom'), true);
assert.equal(kit.platformAdapterRuntimeTarget(manifest, {
  url: 'https://meet.google.com/abc-defg-hij',
}).host_kind, 'browser_extension_content_script');
assert.equal(kit.assertPlatformAdapterRuntimeRecipe({
  platform: 'webex',
}).accepted, true);
assert.equal(kit.assertPlatformAdapterRuntimeManifest({}, {
  platforms: ['google-meet'],
}).platform_registry.rows[0].host_kind, 'browser_extension_content_script');
assert.equal(kit.assertPlatformAdapterRuntimeTarget(manifest, {
  platform: 'zoom',
}).host_kind, 'native_desktop_detector');

const sdk = createMeetingAppTimelineSdk({
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(sdk.platformAdapterRuntimeRecipe({
  platform: 'zoom',
  window: { title: 'Zoom Meeting' },
  process: { name: 'Zoom Workplace' },
}).selected_surface, 'native_detector');
assert.equal(sdk.adapterRuntimeRecipe({
  url: 'https://meet.google.com/abc-defg-hij',
}).host_wiring.bridge_kind, 'browser_content_script');
assert.equal(sdk.platformAdapterRuntimeRecipeMatrix({}, {
  platforms: ['google-meet', 'zoom'],
}).runtime_ready_count, 2);
assert.equal(sdk.platformAdapterRuntimeManifest({}, {
  platforms: ['google-meet', 'zoom'],
}).runtime_contract.timestamp_field, 'captured_at_ms');
assert.equal(sdk.adapterRuntimeManifest({}, {
  platforms: ['google-meet', 'zoom'],
}).platform_registry.rows.find((row) => row.platform === 'zoom').host_kind, 'native_desktop_detector');
assert.equal(sdk.platformAdapterRuntimeTarget(manifest, {
  platform: 'zoom',
}).host_kind, 'native_desktop_detector');
assert.equal(sdk.adapterRuntimeTarget(manifest, {
  url: 'https://meet.google.com/abc-defg-hij',
}).host_kind, 'browser_extension_content_script');
assert.equal(sdk.assertAdapterRuntimeRecipe({
  url: 'https://meet.google.com/abc-defg-hij',
}).accepted, true);
assert.equal(sdk.assertAdapterRuntimeManifest({}, {
  platforms: ['google-meet'],
}).runtime_ready, true);
assert.equal(sdk.assertAdapterRuntimeTarget(manifest, {
  platform: 'zoom',
}).accepted, true);

console.log('ok meeting platform adapter runtime recipe');
