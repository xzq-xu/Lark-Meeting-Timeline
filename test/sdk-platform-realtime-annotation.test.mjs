import assert from 'node:assert/strict';

import {
  buildMeetingPlatformRealtimeAnnotation,
  buildMeetingPlatformRealtimeAnnotationMatrix,
  buildMeetingPlatformRealtimeAnnotationPlan,
} from '../packages/meeting-timeline-sdk/adapters/platform-realtime-annotation.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const startMs = 1_782_614_400_000;

const plan = buildMeetingPlatformRealtimeAnnotationPlan('google-meet');
assert.equal(plan.schema, 'meeting_platform_realtime_annotation_plan');
assert.equal(plan.platform, 'google_meet');
assert.deepEqual(plan.pipeline, ['clock_sync', 'session_binding', 'annotation_intake']);
assert.equal(plan.realtime_policy.provider_events_block_realtime, false);
assert.equal(plan.realtime_policy.transcript_blocks_realtime, false);
assert.equal(plan.clock_sync.schema, 'meeting_platform_clock_sync_plan');
assert.equal(plan.session_binding.schema, 'meeting_platform_session_binding_plan');
assert.equal(plan.annotation_intake.schema, 'meeting_platform_annotation_intake_plan');

const matrix = buildMeetingPlatformRealtimeAnnotationMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(matrix.schema, 'meeting_platform_realtime_annotation_matrix');
assert.equal(matrix.platform_count, 5);
assert.equal(matrix.provider_blocking_count, 0);
assert.equal(matrix.transcript_blocking_count, 0);
assert.equal(matrix.dependencies.clock_sync.length, 5);
assert.equal(matrix.dependencies.session_binding.length, 5);
assert.equal(matrix.dependencies.annotation_intake.length, 5);

const ready = buildMeetingPlatformRealtimeAnnotation('google-meet', {
  clock_sync: {
    offset_ms: 10,
    rtt_ms: 40,
  },
  current_meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    title: 'Design review',
    start_time_ms: startMs,
  },
  local_observer: {
    platform: 'google_meet',
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Design review',
    observed_at_ms: startMs + 1_000,
  },
  annotation: {
    id: 'note-1',
    label: 'why?',
    kind: 'handwriting_trigger',
    captured_at_ms: startMs + 90_000,
  },
});
assert.equal(ready.schema, 'meeting_platform_realtime_annotation');
assert.equal(ready.status, 'ready_to_insert');
assert.equal(ready.accepted_for_realtime, true);
assert.deepEqual(ready.actions, ['insert_mark']);
assert.equal(ready.clock_sync.status, 'clock_sync_ready');
assert.equal(ready.session_binding.status, 'bound_to_current_axis');
assert.equal(ready.annotation_intake.status, 'ready_to_insert_current_axis');
assert.equal(ready.calibrated_annotation.captured_at_ms, startMs + 90_010);
assert.equal(ready.insert_payload.captured_at_ms, startMs + 90_010);
assert.equal(ready.insert_payload.time_ms, 90_010);
assert.equal(ready.selected_meeting.meeting_id, 'abc-defg-hij');

const startAxis = buildMeetingPlatformRealtimeAnnotation('zoom', {
  clock_sync: {
    offset_ms: 0,
    rtt_ms: 40,
  },
  local_observer: {
    url: 'https://zoom.us/j/987654321?pwd=secret',
    title: 'Weekly review - Zoom',
    observed_at_ms: startMs + 1_000,
  },
  annotation: {
    id: 'note-2',
    label: 'follow up',
    captured_at_ms: startMs + 120_000,
  },
});
assert.equal(startAxis.status, 'start_axis_then_insert');
assert.equal(startAxis.accepted_for_realtime, true);
assert.deepEqual(startAxis.actions, ['start_meeting_session', 'insert_mark']);
assert.equal(startAxis.start_payload.meeting_id, '987654321');
assert.equal(startAxis.start_payload.detector_source, 'local_observer_session_binding');
assert.equal(startAxis.insert_payload.time_ms, 119_000);

const pending = buildMeetingPlatformRealtimeAnnotation('lark', {
  clock_sync: {
    offset_ms: 0,
    rtt_ms: 40,
  },
  annotation: {
    id: 'note-3',
    label: 'why?',
    captured_at_ms: startMs + 10_000,
  },
});
assert.equal(pending.status, 'pending_real_meeting');
assert.equal(pending.accepted_for_realtime, true);
assert.deepEqual(pending.actions, ['store_pending_mark', 'rebind_when_meeting_identity_arrives']);
assert.equal(pending.pending_payload.id, 'note-3');
assert.equal(pending.insert_payload, undefined);

const needsClock = buildMeetingPlatformRealtimeAnnotation('google-meet', {
  current_meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    start_time_ms: startMs,
  },
  local_observer: {
    url: 'https://meet.google.com/abc-defg-hij',
    observed_at_ms: startMs + 1_000,
  },
  annotation: {
    id: 'note-4',
    label: 'why?',
    captured_at_ms: startMs + 10_000,
  },
});
assert.equal(needsClock.status, 'needs_clock_sync');
assert.equal(needsClock.accepted_for_realtime, false);
assert.deepEqual(needsClock.actions, ['run_clock_sync_before_realtime_insert']);
assert.equal(needsClock.clock_sync.status, 'clock_sync_missing');
assert.equal(needsClock.insert_payload, undefined);

const conflict = buildMeetingPlatformRealtimeAnnotation('google-meet', {
  clock_sync: {
    offset_ms: 0,
    rtt_ms: 40,
  },
  current_meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    start_time_ms: startMs,
  },
  signals: [
    {
      type: 'meeting_started',
      meeting: {
        platform: 'google_meet',
        meeting_id: 'xyz-uvwx-rst',
      },
      occurred_at_ms: startMs + 1_000,
    },
  ],
  annotation: {
    id: 'note-5',
    label: 'why?',
    captured_at_ms: startMs + 11_000,
  },
});
assert.equal(conflict.status, 'binding_conflict');
assert.equal(conflict.accepted_for_realtime, false);
assert.deepEqual(conflict.actions, ['do_not_insert_until_meeting_identity_conflict_is_resolved']);
assert.equal(conflict.warnings.some((warning) => warning.startsWith('binding_conflict:')), true);

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
assert.equal(kit.platformRealtimeAnnotationPlan('zoom').schema, 'meeting_platform_realtime_annotation_plan');
assert.equal(kit.platformRealtimeAnnotationMatrix({ platforms: ['zoom'] }).platform_count, 1);
assert.equal(kit.platformRealtimeAnnotation('zoom', {
  clock_sync: {
    offset_ms: 0,
    rtt_ms: 40,
  },
  local_observer: {
    url: 'https://zoom.us/j/987654321',
    observed_at_ms: startMs,
  },
  annotation: {
    id: 'kit-note',
    label: 'kit',
    captured_at_ms: startMs + 3_000,
  },
}).status, 'start_axis_then_insert');
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_realtime_annotation_matrix.platform_count, 1);

console.log('ok meeting platform realtime annotation');
