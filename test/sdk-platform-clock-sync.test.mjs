import assert from 'node:assert/strict';

import {
  buildMeetingPlatformClockSyncMatrix,
  buildMeetingPlatformClockSyncPlan,
  buildMeetingPlatformClockSyncReport,
} from '../packages/meeting-timeline-sdk/adapters/platform-clock-sync.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const startMs = 1_782_614_400_000;

const plan = buildMeetingPlatformClockSyncPlan('google-meet');
assert.equal(plan.schema, 'meeting_platform_clock_sync_plan');
assert.equal(plan.platform, 'google_meet');
assert.equal(plan.required, true);
assert.equal(plan.algorithm.name, 'midpoint_offset');
assert.equal(plan.thresholds.max_recommended_skew_ms, 500);
assert.equal(plan.realtime_policy.provider_events_block_realtime, false);
assert.equal(plan.realtime_policy.transcript_blocks_realtime, false);
assert.equal(plan.realtime_policy.device_clock_sync_required, true);

const matrix = buildMeetingPlatformClockSyncMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(matrix.schema, 'meeting_platform_clock_sync_matrix');
assert.equal(matrix.platform_count, 5);
assert.equal(matrix.required_count, 5);
assert.equal(matrix.provider_blocking_count, 0);
assert.equal(matrix.transcript_blocking_count, 0);

const ready = buildMeetingPlatformClockSyncReport('google-meet', {
  samples: [
    {
      id: 'slow',
      client_send_at_ms: startMs,
      server_time_ms: startMs + 100,
      client_receive_at_ms: startMs + 200,
    },
    {
      id: 'best',
      client_send_at_ms: startMs + 1_000,
      server_time_ms: startMs + 1_060,
      client_receive_at_ms: startMs + 1_100,
    },
  ],
  annotation: {
    id: 'note-1',
    label: 'why?',
    captured_at_ms: startMs + 2_000,
  },
});
assert.equal(ready.schema, 'meeting_platform_clock_sync_report');
assert.equal(ready.status, 'clock_sync_ready');
assert.equal(ready.accepted_for_realtime, true);
assert.equal(ready.selected_sample_id, 'best');
assert.equal(ready.recommended_offset_ms, 10);
assert.equal(ready.selected_rtt_ms, 100);
assert.equal(ready.annotation_time.raw_captured_at_ms, startMs + 2_000);
assert.equal(ready.annotation_time.calibrated_captured_at_ms, startMs + 2_010);
assert.equal(ready.calibrated_annotation.captured_at_ms, startMs + 2_010);
assert.equal(ready.calibrated_annotation.clock_sync.offset_ms, 10);

const requiresOffset = buildMeetingPlatformClockSyncReport('zoom', {
  samples: [
    {
      client_send_at_ms: startMs,
      server_time_ms: startMs + 1_050,
      client_receive_at_ms: startMs + 100,
    },
  ],
  annotation: {
    id: 'note-skew',
    label: 'skew',
    captured_at_ms: startMs + 10_000,
  },
});
assert.equal(requiresOffset.status, 'clock_sync_requires_offset');
assert.equal(requiresOffset.accepted_for_realtime, true);
assert.equal(requiresOffset.recommended_offset_ms, 1_000);
assert.equal(requiresOffset.operator_attention_required, true);
assert.equal(requiresOffset.warnings.includes('clock_skew_above_recommended'), true);
assert.equal(requiresOffset.calibrated_annotation.captured_at_ms, startMs + 11_000);

const directOffset = buildMeetingPlatformClockSyncReport('lark', {
  clock_sync: {
    offset_ms: 120,
    rtt_ms: 40,
  },
  annotation: {
    id: 'note-direct',
    label: 'direct',
    captured_at_ms: startMs + 5_000,
  },
});
assert.equal(directOffset.status, 'clock_sync_ready');
assert.equal(directOffset.samples[0].source, 'direct_offset');
assert.equal(directOffset.recommended_offset_ms, 120);
assert.equal(directOffset.calibrated_annotation.captured_at_ms, startMs + 5_120);

const unstable = buildMeetingPlatformClockSyncReport('webex', {
  samples: [
    {
      client_send_at_ms: startMs,
      server_time_ms: startMs + 2_000,
      client_receive_at_ms: startMs + 4_000,
    },
  ],
});
assert.equal(unstable.status, 'clock_sync_unstable');
assert.equal(unstable.accepted_for_realtime, false);
assert.equal(unstable.warnings.includes('rtt_above_recommended'), true);
assert.equal(unstable.warnings.includes('uncertainty_above_recommended'), true);

const missing = buildMeetingPlatformClockSyncReport('teams', {});
assert.equal(missing.status, 'clock_sync_missing');
assert.equal(missing.accepted_for_realtime, false);
assert.equal(missing.next_actions.includes('call_time_sync_endpoint_before_sending_annotations'), true);

const client = {
  async startMeeting(input) { return { ok: true, input }; },
  async endMeeting(input) { return { ok: true, input }; },
  async insertMark(input) { return { ok: true, input }; },
  async insertMarks(input) { return { ok: true, input }; },
};
const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl: 'https://timeline.example.com',
  verify: false,
});
assert.equal(kit.platformClockSyncPlan('zoom').schema, 'meeting_platform_clock_sync_plan');
assert.equal(kit.platformClockSyncMatrix({ platforms: ['zoom'] }).platform_count, 1);
assert.equal(kit.platformClockSync('zoom', {
  clock_sync: { offset_ms: 1, rtt_ms: 10 },
}).status, 'clock_sync_ready');
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_clock_sync_matrix.platform_count, 1);

console.log('ok meeting platform clock sync');
