import assert from 'node:assert/strict';

import {
  createMeetingAppTimelineSdk,
} from '../packages/meeting-timeline-sdk/index.mjs';
import {
  buildMeetingPlatformAdapterMatrix,
  buildMeetingPlatformAdapterMatrixRow,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-matrix.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const googleRow = buildMeetingPlatformAdapterMatrixRow('google-meet');
assert.equal(googleRow.schema, 'meeting_platform_adapter_matrix_row');
assert.equal(googleRow.platform, 'google_meet');
assert.equal(googleRow.adapter_kind, 'browser_first');
assert.equal(googleRow.primary_surface, 'browser_extension');
assert.equal(googleRow.realtime_axis.timestamp_field, 'captured_at_ms');
assert.equal(googleRow.realtime_axis.provider_events_block_realtime, false);
assert.equal(googleRow.realtime_axis.transcript_blocks_realtime, false);
assert.equal(googleRow.provider_reconcile.required_for_realtime, false);
assert.equal(googleRow.provider_reconcile.required_for_production, true);
assert.equal(googleRow.speaker_positions.enabled, true);
assert.equal(googleRow.speaker_positions.content_policy, 'speaker_position_only_no_transcript_text_required');
assert.equal(googleRow.post_meeting.transcript_available, true);
assert.equal(googleRow.browser_integration.matches.includes('https://meet.google.com/*'), true);
assert.equal(googleRow.next_actions.includes('insert_annotations_with_captured_at_ms'), true);

const zoomRow = buildMeetingPlatformAdapterMatrixRow('zoom');
assert.equal(zoomRow.adapter_kind, 'native_first');
assert.equal(zoomRow.primary_surface, 'native_detector');
assert.equal(zoomRow.provider_reconcile.transport, 'Zoom Meeting webhooks');
assert.equal(zoomRow.realtime_axis.end_fallbacks.includes('manual_stop_when_zoom_webhook_is_late'), true);

const matrix = buildMeetingPlatformAdapterMatrix({
  platforms: ['google-meet', 'teams', 'zoom'],
});
assert.equal(matrix.schema, 'meeting_platform_adapter_matrix');
assert.equal(matrix.platform_count, 3);
assert.equal(matrix.realtime_ready_count, 3);
assert.equal(matrix.accepted_count, 3);
assert.equal(matrix.browser_first_count, 1);
assert.equal(matrix.native_first_count, 2);
assert.equal(matrix.provider_reconcile_required_count, 3);
assert.equal(matrix.speaker_position_enabled_count, 3);
assert.equal(matrix.post_meeting_transcript_count, 3);
assert.deepEqual(matrix.platforms, ['google_meet', 'microsoft_teams', 'zoom']);
assert.equal(matrix.rows[0].provider_reconcile_required, true);
assert.equal(matrix.adapters.find((row) => row.platform === 'microsoft_teams').primary_surface, 'native_detector');

const kit = createMeetingPlatformTimelineKit({ baseUrl: 'http://localhost:8787' });
assert.equal(kit.platformAdapterMatrix({ platforms: ['google-meet'] }).platform_count, 1);
assert.equal(kit.platformAdapterMatrixRow('google-meet').runtime.runtime_target_message_type, 'meeting_timeline.runtime_target');

const sdk = createMeetingAppTimelineSdk({ baseUrl: 'http://localhost:8787' });
assert.equal(sdk.platformAdapterMatrix({ platforms: ['google-meet'] }).rows[0].primary_surface, 'browser_extension');
assert.equal(sdk.adapterMatrixRow('zoom').primary_surface, 'native_detector');

console.log('ok platform adapter matrix');
