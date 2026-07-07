import assert from 'node:assert/strict';

import {
  assertMeetingPlatformAdapterDecision,
  assertMeetingPlatformAdapterDecisionMatrix,
  buildMeetingPlatformAdapterDecision,
  buildMeetingPlatformAdapterDecisionMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-decision.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import {
  createMeetingAppTimelineSdk,
} from '../packages/meeting-timeline-sdk/index.mjs';

const baseUrl = 'https://timeline.example.com';

const google = buildMeetingPlatformAdapterDecision({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Google Meet',
}, {
  baseUrl,
});
assert.equal(google.schema, 'meeting_platform_adapter_decision');
assert.equal(google.accepted, true);
assert.equal(google.realtime_ready, true);
assert.equal(google.platform, 'google_meet');
assert.equal(google.platform_source, 'url_detection');
assert.equal(google.selected_surface, 'browser_extension');
assert.equal(google.selected_route, 'local_observer_axis');
assert.equal(google.adapter_blueprint.ready, true);
assert.equal(google.adapter_blueprint.primary_surface, 'browser_extension');
assert.equal(google.adapter_blueprint.selected_surface_recommended, true);
assert.equal(google.adapter_blueprint.first_acceptance_gate, 'local_candidate_preflight_accepts_active_meeting');
assert.equal(google.contracts.timestamp_field, 'captured_at_ms');
assert.equal(google.contracts.adapter_blueprint_required_before_host_wiring, true);
assert.equal(google.contracts.provider_events_block_realtime, false);
assert.equal(google.contracts.transcript_blocks_realtime, false);
assert.equal(google.route_summary.provider_reconcile.required_for_realtime, false);
assert.equal(google.route_summary.post_meeting_artifact.blocks_realtime_if_missing, false);
assert.equal(google.runtime_actions.some((action) => action.action === 'observe_platform_candidates'), true);
assert.equal(google.runtime_actions.some((action) => action.action === 'insert_annotation'), true);
assert.equal(google.runtime_actions.some((action) => action.action === 'provider_event' && action.blocks_realtime_annotation === false), true);
assert.equal(google.evidence_requirements.realtime_minimum.includes('annotation_captured_at_ms'), true);
assert.equal(assertMeetingPlatformAdapterDecision({
  url: 'https://meet.google.com/abc-defg-hij',
}, {
  baseUrl,
}).platform, 'google_meet');

const liveSnapshotDecision = buildMeetingPlatformAdapterDecision({
  platform: 'google-meet',
  url: 'https://meet.google.com/abc-defg-hij',
  page: {
    controls: [{ label: 'Leave call' }],
    participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
  },
}, {
  baseUrl,
});
assert.equal(liveSnapshotDecision.timeline_capabilities.realtime_axis.status, 'local_ready');
assert.equal(liveSnapshotDecision.timeline_capabilities.speaker_track.status, 'local_ready');
assert.equal(liveSnapshotDecision.runtime_actions.some((action) => action.action === 'speaker_track'), true);

const zoomNative = buildMeetingPlatformAdapterDecision({
  platform: 'zoom',
  window: { title: 'Zoom Meeting' },
  process: { name: 'Zoom Workplace' },
}, {
  baseUrl,
});
assert.equal(zoomNative.platform, 'zoom');
assert.equal(zoomNative.selected_surface, 'native_detector');
assert.equal(zoomNative.surface_source, 'native_context');

const larkProvider = buildMeetingPlatformAdapterDecision({
  platform: 'lark',
  provider_event: { type: 'vc.meeting.all_meeting_started_v1' },
}, {
  baseUrl,
  providerOnly: true,
});
assert.equal(larkProvider.selected_surface, 'provider_reconcile');
assert.equal(larkProvider.runtime_actions.some((action) => action.action === 'provider_event'), true);
assert.equal(larkProvider.contracts.provider_events_block_realtime, false);

const missing = buildMeetingPlatformAdapterDecision({});
assert.equal(missing.accepted, false);
assert.equal(missing.status, 'missing_platform');
assert.throws(
  () => assertMeetingPlatformAdapterDecision({}),
  /Meeting platform adapter decision is not accepted/,
);

const matrix = buildMeetingPlatformAdapterDecisionMatrix({}, {
  baseUrl,
  platforms: ['google-meet', 'zoom', 'teams'],
});
assert.equal(matrix.schema, 'meeting_platform_adapter_decision_matrix');
assert.deepEqual(matrix.platforms, ['google_meet', 'zoom', 'microsoft_teams']);
assert.equal(matrix.platform_count, 3);
assert.equal(matrix.accepted_count, 3);
assert.equal(matrix.realtime_ready_count, 3);
assert.equal(matrix.adapter_blueprint_ready_count, 3);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').adapter_blueprint_primary_surface, 'browser_extension');
assert.equal(assertMeetingPlatformAdapterDecisionMatrix({}, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
}).accepted_count, 2);

const kit = createMeetingPlatformTimelineKit({}, {
  baseUrl,
});
assert.equal(kit.platformAdapterDecision({
  url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
}).platform, 'microsoft_teams');
assert.equal(kit.platformAdapterDecisionMatrix({}, {
  platforms: ['webex', 'lark'],
}).platform_count, 2);
assert.equal(kit.assertPlatformAdapterDecision({
  platform: 'webex',
}).accepted, true);

const sdk = createMeetingAppTimelineSdk({
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(sdk.platformAdapterDecision({
  url: 'https://zoom.us/j/987654321',
}).platform, 'zoom');
assert.equal(sdk.adapterDecision({
  url: 'https://meet.google.com/abc-defg-hij',
}).selected_surface, 'browser_extension');
assert.equal(sdk.platformAdapterDecisionMatrix({}, {
  platforms: ['google-meet', 'zoom'],
}).accepted_count, 2);
assert.equal(sdk.adapterDecisionMatrix({}, {
  platforms: ['google-meet'],
}).rows[0].provider_events_block_realtime, false);
assert.equal(sdk.assertAdapterDecision({
  url: 'https://meet.google.com/abc-defg-hij',
}).accepted, true);

console.log('ok meeting platform adapter decision');
