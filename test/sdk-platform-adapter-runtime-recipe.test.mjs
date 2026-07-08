import assert from 'node:assert/strict';

import {
  assertMeetingPlatformAdapterRuntimeRecipe,
  assertMeetingPlatformAdapterRuntimeRecipeMatrix,
  buildMeetingPlatformAdapterRuntimeRecipe,
  buildMeetingPlatformAdapterRuntimeRecipeMatrix,
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
assert.equal(kit.assertPlatformAdapterRuntimeRecipe({
  platform: 'webex',
}).accepted, true);

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
assert.equal(sdk.assertAdapterRuntimeRecipe({
  url: 'https://meet.google.com/abc-defg-hij',
}).accepted, true);

console.log('ok meeting platform adapter runtime recipe');
