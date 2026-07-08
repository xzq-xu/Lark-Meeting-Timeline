import assert from 'node:assert/strict';

import {
  buildMeetingPlatformImplementationHandoff,
  buildMeetingPlatformImplementationHandoffMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-implementation-handoff.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import {
  createMeetingAppTimelineSdk,
} from '../packages/meeting-timeline-sdk/index.mjs';

const baseUrl = 'https://timeline.example.com';

const google = buildMeetingPlatformImplementationHandoff('google-meet', { baseUrl });
assert.equal(google.schema, 'meeting_platform_implementation_handoff');
assert.equal(google.platform, 'google_meet');
assert.equal(google.implementation_ready, true);
assert.equal(google.pilot_ready, true);
assert.equal(google.production_ready, false);
assert.equal(google.recommended_first_surface, 'browser_extension');
assert.equal(google.package_entrypoints.root_create_function, 'createMeetingAppTimelineSdk');
assert.equal(google.package_entrypoints.implementation_handoff_module, '@ai-annotation/meeting-timeline-sdk/adapters/platform-implementation-handoff');
assert.equal(google.package_entrypoints.adapter_preflight_module, '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-preflight');
assert.equal(google.install_surface.browser_matches.includes('https://meet.google.com/*'), true);
assert.equal(google.install_surface.lightweight_connector_bridge.install_function, 'installMeetingPlatformConnectorContentScriptBridge');
assert.equal(google.runtime_events.endpoint, `${baseUrl}/api/meeting-platform/runtime-events`);
assert.equal(google.runtime_events.message_types.includes('meeting_timeline.insert_mark'), true);
assert.equal(google.runtime_events.actions.includes('insert_annotation'), true);
assert.equal(google.runtime_events.candidate_observation.message_type, 'meeting_timeline.observe_candidates');
assert.equal(google.provider_reconcile.path, 'google_workspace_events_pubsub');
assert.equal(google.provider_reconcile.required_for_realtime, false);
assert.equal(google.provider_reconcile.realtime_blocking, false);
assert.equal(google.adapter_preflight.status, 'needs_live_page_evidence');
assert.equal(google.adapter_preflight.selected_surface, 'browser_extension');
assert.equal(google.adapter_preflight.startup_ready, true);
assert.equal(google.adapter_preflight.live_evidence_ready, false);
assert.equal(google.adapter_preflight.realtime_annotation_ready, false);
assert.equal(google.adapter_preflight.live_evidence_required, true);
assert.equal(google.adapter_preflight.url_only_status, 'needs_live_page_evidence');
assert.equal(google.adapter_preflight.sdk_methods.includes('sdk.platformAdapterPreflight({ platform, snapshots/nativeEvidence })'), true);
assert.equal(google.adapter_preflight.bridge_messages.includes('meeting_timeline.preflight_current_window'), true);
assert.equal(google.contracts.timestamp_field, 'captured_at_ms');
assert.equal(google.contracts.provider_events_block_realtime, false);
assert.equal(google.contracts.transcript_blocks_realtime, false);
assert.equal(google.contracts.adapter_preflight_required_before_realtime_insert, true);
assert.equal(google.contracts.adapter_preflight_live_evidence_required, true);
assert.equal(google.contracts.adapter_preflight_url_only_status, 'needs_live_page_evidence');
assert.equal(google.implementation_flow[0].id, 'install_sdk');
assert.equal(google.implementation_flow.find((step) => step.id === 'run_adapter_preflight').required_before, 'bind_current_axis_or_insert_realtime_annotation');
assert.equal(google.implementation_flow.find((step) => step.id === 'run_adapter_preflight').url_only_status, 'needs_live_page_evidence');
assert.equal(google.implementation_flow.find((step) => step.id === 'bind_current_axis').required_before, 'insert_annotation');
assert.equal(google.implementation_flow.find((step) => step.id === 'insert_realtime_annotations').required_field, 'captured_at_ms');
assert.equal(google.acceptance.commands.implementation_handoff.includes('meeting-platform:implementation-handoff'), true);
assert.equal(google.acceptance.commands.consumer_handoff.includes('--platforms=google_meet'), true);
assert.equal(google.acceptance.commands.adapter_preflight.includes('meeting-platform:adapter-preflight'), true);
assert.equal(google.acceptance.live_evidence_condition, 'adapter_preflight.realtime_annotation_ready === true before first insertAnnotation');
assert.equal(google.readiness.runtime_ready, true);
assert.equal(google.readiness.lightweight_connector_ready, true);
assert.equal(google.readiness.adapter_preflight_startup_ready, true);
assert.equal(google.readiness.adapter_preflight_live_evidence_ready, false);
assert.equal(google.readiness.adapter_preflight_realtime_ready, false);
assert.equal(google.readiness.provider_required_for_realtime, false);
assert.equal(google.production_gaps.includes('production_evidence_pending'), true);
assert.equal(google.next_actions.includes('collect_live_dom_or_native_window_evidence_for_adapter_preflight'), true);
assert.equal(google.next_actions.includes('install_meeting_platform_connector_content_script_bridge'), true);

const matrix = buildMeetingPlatformImplementationHandoffMatrix({
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(matrix.schema, 'meeting_platform_implementation_handoff_matrix');
assert.equal(matrix.platform_count, 5);
assert.equal(matrix.implementation_ready_count, 5);
assert.equal(matrix.pilot_ready_count, 5);
assert.equal(matrix.production_ready_count, 0);
assert.equal(matrix.adapter_preflight_startup_ready_count, 5);
assert.equal(matrix.adapter_preflight_realtime_ready_count, 0);
assert.equal(matrix.recommended_first_platform, 'google_meet');
assert.equal(matrix.recommended_first_surface, 'browser_extension');
assert.deepEqual(matrix.priority_order.slice(0, 3), ['google_meet', 'zoom', 'microsoft_teams']);
assert.equal(matrix.rows[0].platform, 'google_meet');
assert.equal(matrix.rows[1].platform, 'zoom');
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').provider_path, 'microsoft_graph_change_notifications');
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').adapter_preflight_status, 'needs_live_page_evidence');
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').adapter_preflight_selected_surface, 'native_detector');
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').adapter_preflight_startup_ready, true);
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').adapter_preflight_realtime_ready, false);
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').browser_match_count, 2);
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').browser_match_count, 3);
assert.equal(matrix.handoffs.find((handoff) => handoff.platform === 'webex').provider_reconcile.path, 'webex_webhooks');

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
assert.equal(kit.platformImplementationHandoff('zoom').provider_reconcile.path, 'zoom_meeting_webhooks');
assert.equal(kit.platformImplementationHandoffMatrix().platform_count, 2);
assert.equal(kit.report().platform_implementation_handoff_matrix.implementation_ready_count, 2);

const sdk = createMeetingAppTimelineSdk({
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(sdk.platformImplementationHandoff('google-meet').platform, 'google_meet');
assert.equal(sdk.implementationHandoff('zoom').provider_reconcile.path, 'zoom_meeting_webhooks');
assert.equal(sdk.platformImplementationHandoffMatrix().platform_count, 2);
assert.equal(sdk.implementationHandoffMatrix().recommended_first_platform, 'google_meet');

console.log('ok meeting platform implementation handoff');
