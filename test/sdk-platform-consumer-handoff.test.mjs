import assert from 'node:assert/strict';

import {
  assertMeetingPlatformConsumerHandoff,
  buildMeetingPlatformConsumerHandoff,
} from '../packages/meeting-timeline-sdk/adapters/platform-consumer-handoff.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';

const handoff = buildMeetingPlatformConsumerHandoff({
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

assert.equal(handoff.schema, 'meeting_platform_consumer_handoff');
assert.equal(handoff.accepted, true);
assert.equal(handoff.platform_count, 5);
assert.equal(handoff.consumer_ready_count, 5);
assert.equal(handoff.conformance_accepted_count, 5);
assert.equal(handoff.runtime_ready_count, 5);
assert.equal(handoff.lightweight_connector_ready, true);
assert.equal(handoff.lightweight_connector_platform_count, 5);
assert.equal(handoff.adapter_route_ready_count, 5);
assert.equal(handoff.candidate_observer_count, 5);
assert.equal(handoff.speaker_track_ready_count, 5);
assert.equal(handoff.participant_track_ready_count, 5);
assert.equal(handoff.production_ready_count, 0);
assert.equal(handoff.blocking_count, 0);
assert.equal(handoff.warning_count, 5);
assert.equal(handoff.entrypoints.primary_modules.consumer_handoff, '@ai-annotation/meeting-timeline-sdk/adapters/platform-consumer-handoff');
assert.equal(handoff.entrypoints.primary_modules.app_adapter_manifest, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-manifest');
assert.equal(handoff.entrypoints.primary_modules.app_adapter_spec, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-spec');
assert.equal(handoff.entrypoints.primary_modules.app_adapter_runtime_config, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-runtime-config');
assert.equal(handoff.entrypoints.primary_modules.meeting_platform_connector, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-connector');
assert.equal(handoff.entrypoints.kit_methods.includes('platformConsumerHandoff'), true);
assert.equal(handoff.entrypoints.kit_methods.includes('meetingAppAdapterManifestMatrix'), true);
assert.equal(handoff.entrypoints.kit_methods.includes('meetingAppAdapterSpecMatrix'), true);
assert.equal(handoff.entrypoints.kit_methods.includes('meetingAppAdapterRuntimeConfigMatrix'), true);
assert.equal(handoff.entrypoints.kit_methods.includes('createPlatformConnectorHub'), true);
assert.equal(handoff.entrypoints.kit_methods.includes('installPlatformConnectorContentScriptBridge'), true);
assert.equal(handoff.entrypoints.host_methods.includes('consumerHandoff'), true);
assert.match(handoff.entrypoints.commands.app_adapter_manifest, /meeting-app:adapter-manifest/);
assert.equal(handoff.endpoints.consumer_handoff, '/api/meeting-platform/consumer-handoff');
assert.equal(handoff.hard_contracts.timestamp_field, 'captured_at_ms');
assert.equal(handoff.hard_contracts.provider_events_block_realtime, false);
assert.equal(handoff.hard_contracts.transcript_blocks_realtime, false);
assert.equal(handoff.hard_contracts.lightweight_connector_bridge_supported, true);
assert.equal(handoff.hard_contracts.connector_bridge_message_types.includes('meeting_timeline.sample_tracks'), true);
assert.equal(handoff.lightweight_connector_handoff.accepted, true);
assert.equal(handoff.lightweight_connector_handoff.module, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-connector');
assert.equal(handoff.lightweight_connector_handoff.factories.install_content_script_bridge, 'installMeetingPlatformConnectorContentScriptBridge');
assert.equal(handoff.lightweight_connector_handoff.content_script_bridge.message_types.includes('meeting_timeline.insert_mark'), true);
assert.equal(handoff.lightweight_connector_handoff.content_script_bridge.output_runtime_actions.includes('insert_annotation'), true);
assert.equal(handoff.lightweight_connector_handoff.host_requirements.timestamp_field, 'captured_at_ms');
assert.equal(handoff.sdk_facade_handoff.type, 'meeting_platform_sdk_facade_handoff');
assert.equal(handoff.sdk_facade_handoff.package, '@ai-annotation/meeting-timeline-sdk');
assert.equal(handoff.sdk_facade_handoff.create_function, 'createMeetingAppTimelineSdk');
assert.equal(handoff.sdk_facade_handoff.constructor_options.baseUrl, baseUrl);
assert.equal(handoff.sdk_facade_handoff.required_facade_methods.includes('platformAdaptationPackage'), true);
assert.equal(handoff.sdk_facade_handoff.required_facade_methods.includes('platformRuntimeBundle'), true);
assert.equal(handoff.sdk_facade_handoff.minimal_realtime_flow[0].method, 'createMeetingAppTimelineSdk({ baseUrl, platforms })');
assert.equal(handoff.sdk_facade_handoff.minimal_realtime_flow.some((step) => String(step.method).includes('sdk.insertAnnotation')), true);
assert.equal(handoff.sdk_facade_handoff.surface_wiring.browser_extension.content_script_messages.includes('meeting_timeline.sample_tracks'), true);
assert.equal(handoff.sdk_facade_handoff.surface_wiring.provider_adapter.realtime_blocking, false);
assert.equal(handoff.sdk_facade_handoff.timestamp_field, 'captured_at_ms');
assert.equal(handoff.sdk_facade_handoff.platform_rows.find((row) => row.platform === 'google_meet').provider_path, 'google_workspace_events_pubsub');
assert.equal(handoff.sdk_facade_handoff.platform_rows.find((row) => row.platform === 'google_meet').facade_methods.runtime_bundle, "sdk.platformRuntimeBundle('google_meet')");
assert.equal(handoff.surface_coverage_matrix.schema, 'meeting_platform_surface_coverage_matrix');
assert.equal(handoff.surface_coverage_matrix.platform_count, 5);
assert.equal(handoff.surface_coverage_matrix.browser_extension_ready_count, 5);
assert.equal(handoff.surface_coverage_matrix.webview_preload_ready_count, 5);
assert.equal(handoff.surface_coverage_matrix.native_detector_ready_count, 5);
assert.equal(handoff.surface_coverage_matrix.provider_reconcile_ready_count, 5);
assert.equal(handoff.surface_coverage_matrix.post_meeting_backfill_supported_count, 5);
assert.equal(handoff.surface_coverage_matrix.lightweight_connector_ready_count, 5);
assert.equal(handoff.surface_coverage_matrix.rows.find((row) => row.platform === 'google_meet').browser_extension.message_type, 'meeting_timeline.observe_candidates');
assert.equal(handoff.surface_coverage_matrix.rows.find((row) => row.platform === 'google_meet').provider_reconcile.provider_path, 'google_workspace_events_pubsub');
assert.equal(handoff.surface_coverage_matrix.rows.find((row) => row.platform === 'microsoft_teams').provider_reconcile.provider_path, 'microsoft_graph_change_notifications');
assert.equal(handoff.surface_coverage_matrix.rows.find((row) => row.platform === 'zoom').post_meeting_backfill.supported, true);
assert.equal(handoff.boot_order[0].action, 'run_static_consumer_handoff');
assert.equal(handoff.boot_order.some((step) => step.action === 'choose_lightweight_connector_or_full_integration_runtime'), true);
assert.equal(handoff.rows.find((row) => row.platform === 'google_meet').adapter_first_route, 'local_observer_axis');
assert.equal(handoff.rows.every((row) => row.consumer_ready === true), true);
assert.equal(handoff.rows.every((row) => row.production_ready === false), true);
assert.equal(handoff.conformance_report.accepted, true);
assert.equal(handoff.host_integration_plan.runtime_contract.annotation_timestamp_field, 'captured_at_ms');
assert.equal(handoff.adaptation_package_matrix.packages, undefined);
assert.equal(handoff.handoff_readiness_matrix.reports, undefined);
assert.equal(assertMeetingPlatformConsumerHandoff(handoff).accepted, true);

assert.throws(() => assertMeetingPlatformConsumerHandoff({
  baseUrl,
  platforms: ['google-meet'],
  requireProductionReady: true,
}), /Meeting platform consumer handoff is not ready/);

const kit = createMeetingPlatformTimelineKit({ baseUrl });
const zoomHandoff = kit.platformConsumerHandoff({ platforms: ['zoom'] });
assert.equal(zoomHandoff.platform_count, 1);
assert.equal(zoomHandoff.rows[0].platform, 'zoom');
assert.equal(kit.assertPlatformConsumerHandoff({ platforms: ['zoom'] }).accepted, true);
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_consumer_handoff.platform_count, 1);

console.log('ok meeting platform consumer handoff');
