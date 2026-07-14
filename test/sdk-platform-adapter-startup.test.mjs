import assert from 'node:assert/strict';

import {
  assertMeetingPlatformAdapterStartupPlan,
  assertMeetingPlatformAdapterStartupPlanMatrix,
  buildMeetingPlatformAdapterStartupPlan,
  buildMeetingPlatformAdapterStartupPlanMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-startup.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import {
  createMeetingAppTimelineSdk,
} from '../packages/meeting-timeline-sdk/index.mjs';

const baseUrl = 'https://timeline.example.com';

const google = buildMeetingPlatformAdapterStartupPlan({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Google Meet',
}, {
  baseUrl,
});
assert.equal(google.schema, 'meeting_platform_adapter_startup_plan');
assert.equal(google.accepted, true);
assert.equal(google.realtime_startup_ready, true);
assert.equal(google.platform, 'google_meet');
assert.equal(google.selected_surface, 'browser_extension');
assert.equal(google.install_target, 'manifest_v3_content_script');
assert.equal(google.adapter_blueprint.ready, true);
assert.equal(google.adapter_blueprint.primary_surface, 'browser_extension');
assert.equal(google.adapter_blueprint.first_acceptance_gate, 'local_candidate_preflight_accepts_active_meeting');
assert.equal(google.browser.matches.includes('https://meet.google.com/*'), true);
assert.equal(google.browser.content_script.matches.includes('https://meet.google.com/*'), true);
assert.equal(google.bridge.install_function, 'installMeetingPlatformConnectorContentScriptBridge');
assert.equal(google.runtime.preset, 'google_meet');
assert.equal(google.runtime_contract.annotation_timestamp_field, 'captured_at_ms');
assert.equal(google.runtime_contract.adapter_blueprint_required_before_host_wiring, true);
assert.equal(google.runtime_contract.provider_events_block_realtime, false);
assert.equal(google.runtime_contract.transcript_blocks_realtime, false);
assert.equal(google.message_contract.observe_candidates, 'meeting_timeline.observe_candidates');
assert.equal(google.message_contract.insert_annotation, 'meeting_timeline.insert_mark');
assert.equal(google.actions.some((action) => action.id === 'install_page_bridge'), true);
assert.equal(google.actions.some((action) => action.id === 'read_adapter_blueprint'), true);
assert.equal(google.actions.some((action) => action.id === 'observe_axis'), true);
assert.equal(google.actions.some((action) => action.id === 'insert_realtime_annotation'), true);
assert.equal(google.provider_reconcile.required_for_realtime, false);
assert.equal(google.next_actions.includes('wire_startup_plan_into_host_runtime'), true);
assert.equal(assertMeetingPlatformAdapterStartupPlan({
  url: 'https://meet.google.com/abc-defg-hij',
}, {
  baseUrl,
}).platform, 'google_meet');

const zoomNative = buildMeetingPlatformAdapterStartupPlan({
  platform: 'zoom',
  window: { title: 'Zoom Meeting' },
  process: { name: 'Zoom Workplace' },
}, {
  baseUrl,
});
assert.equal(zoomNative.selected_surface, 'native_detector');
assert.equal(zoomNative.install_target, 'native_or_desktop_observer');
assert.equal(zoomNative.realtime_startup_ready, true);
assert.equal(zoomNative.bridge.first_method, 'observePlatformCandidates');
assert.equal(zoomNative.actions.some((action) => action.id === 'install_page_bridge'), false);
assert.equal(zoomNative.code_refs.insert_call.includes('insertAnnotation'), true);

const larkProvider = buildMeetingPlatformAdapterStartupPlan({
  platform: 'lark',
  provider_event: { type: 'vc.meeting.all_meeting_started_v1' },
}, {
  baseUrl,
  providerOnly: true,
});
assert.equal(larkProvider.selected_surface, 'provider_reconcile');
assert.equal(larkProvider.accepted, false);
assert.equal(larkProvider.realtime_startup_ready, false);
assert.equal(larkProvider.issues[0].code, 'decision_not_accepted');
assert.equal(larkProvider.message_contract.realtime_axis_source, 'not_allowed_as_primary_realtime_axis');
assert.equal(larkProvider.next_actions.includes('add_local_surface_for_realtime_annotations'), true);

const missing = buildMeetingPlatformAdapterStartupPlan({});
assert.equal(missing.accepted, false);
assert.equal(missing.status, 'missing_platform');
assert.throws(
  () => assertMeetingPlatformAdapterStartupPlan({}),
  /Meeting platform adapter startup plan is not accepted/,
);

const matrix = buildMeetingPlatformAdapterStartupPlanMatrix({}, {
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom'],
});
assert.equal(matrix.schema, 'meeting_platform_adapter_startup_plan_matrix');
assert.equal(matrix.platform_count, 3);
assert.equal(matrix.accepted_count, 3);
assert.equal(matrix.realtime_startup_ready_count, 3);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').adapter_blueprint_ready, true);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').adapter_blueprint_primary_surface, 'browser_extension');
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').selected_surface, 'native_detector');
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').install_target, 'native_or_desktop_observer');
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').selected_surface, 'native_detector');
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').provider_events_block_realtime, false);
assert.equal(assertMeetingPlatformAdapterStartupPlanMatrix({}, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
}).accepted_count, 2);

const kit = createMeetingPlatformTimelineKit({}, {
  baseUrl,
});
assert.equal(kit.platformAdapterStartupPlan({
  url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
}).platform, 'microsoft_teams');
assert.equal(kit.platformAdapterStartupPlanMatrix({}, {
  platforms: ['webex', 'lark'],
}).platform_count, 2);
assert.equal(kit.assertPlatformAdapterStartupPlan({
  platform: 'webex',
}).accepted, true);

const sdk = createMeetingAppTimelineSdk({
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(sdk.platformAdapterStartupPlan({
  url: 'https://zoom.us/j/987654321',
}).platform, 'zoom');
assert.equal(sdk.adapterStartupPlan({
  url: 'https://meet.google.com/abc-defg-hij',
}).selected_surface, 'browser_extension');
assert.equal(sdk.platformAdapterStartupPlanMatrix({}, {
  platforms: ['google-meet', 'zoom'],
}).realtime_startup_ready_count, 2);
assert.equal(sdk.assertAdapterStartupPlan({
  url: 'https://meet.google.com/abc-defg-hij',
}).accepted, true);

console.log('ok meeting platform adapter startup plan');
