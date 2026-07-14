import assert from 'node:assert/strict';

import {
  buildMeetingPlatformAnnotationIntake,
  buildMeetingPlatformAnnotationIntakeMatrix,
  buildMeetingPlatformAnnotationIntakePlan,
} from '../packages/meeting-timeline-sdk/adapters/platform-annotation-intake.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const startMs = 1_782_614_400_000;
const currentMeeting = {
  platform: 'google_meet',
  meeting_id: 'abc-defg-hij',
  meeting_url: 'https://meet.google.com/abc-defg-hij',
  title: 'Design review',
  start_time_ms: startMs,
};

const plan = buildMeetingPlatformAnnotationIntakePlan('google-meet');
assert.equal(plan.schema, 'meeting_platform_annotation_intake_plan');
assert.equal(plan.platform, 'google_meet');
assert.equal(plan.accepted_time_fields.includes('captured_at_ms'), true);
assert.equal(plan.realtime_policy.provider_events_block_realtime, false);
assert.equal(plan.realtime_policy.transcript_blocks_realtime, false);
assert.equal(plan.route_policy.allow_open_session_when_no_axis, true);
assert.equal(plan.output_contract.annotation_timestamp_field, 'captured_at_ms');

const matrix = buildMeetingPlatformAnnotationIntakeMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(matrix.schema, 'meeting_platform_annotation_intake_matrix');
assert.equal(matrix.platform_count, 5);
assert.equal(matrix.provider_blocking_count, 0);
assert.equal(matrix.transcript_blocking_count, 0);

const ready = buildMeetingPlatformAnnotationIntake('google-meet', {
  current_meeting: currentMeeting,
  annotation: {
    id: 'note-1',
    label: 'why?',
    kind: 'handwriting_trigger',
    captured_at_ms: startMs + 90_000,
  },
});
assert.equal(ready.schema, 'meeting_platform_annotation_intake');
assert.equal(ready.status, 'ready_to_insert_current_axis');
assert.equal(ready.action, 'insert_mark');
assert.equal(ready.accepted_for_realtime, true);
assert.equal(ready.normalized_time_ms, 90_000);
assert.equal(ready.insert_payload.time_ms, 90_000);
assert.equal(ready.insert_payload.captured_at_ms, startMs + 90_000);
assert.deepEqual(ready.warnings, []);

const openSession = buildMeetingPlatformAnnotationIntake('zoom', {
  annotation: {
    id: 'note-open',
    label: 'follow up',
    captured_at_ms: startMs + 10_000,
  },
});
assert.equal(openSession.status, 'start_open_session_then_insert');
assert.equal(openSession.action, 'start_open_session_then_insert_mark');
assert.equal(openSession.accepted_for_realtime, true);
assert.equal(openSession.open_session_payload.meeting_id, `annotation-open-session-${startMs + 10_000}`);
assert.equal(openSession.open_session_payload.start_time_ms, startMs + 10_000);
assert.equal(openSession.insert_payload.time_ms, undefined);

const pendingWhenNoAxis = buildMeetingPlatformAnnotationIntake('zoom', {
  annotation: {
    id: 'note-pending',
    label: 'pending',
    captured_at_ms: startMs + 10_000,
  },
}, {
  allowOpenSessionWhenNoAxis: false,
});
assert.equal(pendingWhenNoAxis.status, 'pending_real_meeting');
assert.equal(pendingWhenNoAxis.action, 'store_pending_mark');
assert.equal(pendingWhenNoAxis.warnings.includes('pending_real_meeting'), true);

const missingTime = buildMeetingPlatformAnnotationIntake('lark', {
  annotation: {
    id: 'note-missing',
    label: 'why?',
  },
});
assert.equal(missingTime.status, 'needs_device_captured_at');
assert.equal(missingTime.action, 'reject_or_store_unbound');
assert.equal(missingTime.requires_device_captured_at, true);
assert.equal(missingTime.warnings.includes('missing_captured_at_ms'), true);
assert.equal(missingTime.accepted_for_realtime, false);

const ended = buildMeetingPlatformAnnotationIntake('google-meet', {
  current_meeting: {
    ...currentMeeting,
    end_time_ms: startMs + 600_000,
  },
  annotation: {
    id: 'note-after-end',
    label: 'late',
    captured_at_ms: startMs + 720_000,
  },
});
assert.equal(ended.status, 'after_meeting_end');
assert.equal(ended.action, 'store_for_audit_only');
assert.equal(ended.after_meeting_end_ms, 120_000);
assert.equal(ended.warnings.includes('normalized_after_meeting_end'), true);
assert.equal(ended.accepted_for_realtime, false);

const strokeInferred = buildMeetingPlatformAnnotationIntake('google-meet', {
  current_meeting: currentMeeting,
  annotation: {
    id: 'note-stroke',
    label: 'stroke inferred',
    strokes: [
      [{ x: 1, y: 1, t: startMs + 1_000 }],
      [{ x: 2, y: 2, t: startMs + 2_000 }],
    ],
  },
});
assert.equal(strokeInferred.status, 'ready_to_insert_current_axis');
assert.equal(strokeInferred.captured_at_source, 'stroke_point_time');
assert.equal(strokeInferred.captured_at_ms, startMs + 2_000);
assert.equal(strokeInferred.normalized_time_ms, 2_000);
assert.equal(strokeInferred.insert_payload.captured_at_ms, startMs + 2_000);

const pendingAxis = buildMeetingPlatformAnnotationIntake('webex', {
  current_meeting: {
    platform: 'webex',
    pending_binding: true,
  },
  annotation: {
    id: 'note-pending-axis',
    label: 'pending axis',
    captured_at_ms: startMs + 3_000,
  },
});
assert.equal(pendingAxis.axis_state, 'pending_axis');
assert.equal(pendingAxis.status, 'pending_real_meeting');

const localSimulation = buildMeetingPlatformAnnotationIntake('teams', {
  current_meeting: {
    platform: 'microsoft_teams',
    meeting_id: 'local-demo',
    source: 'local_simulation',
  },
  annotation: {
    id: 'note-local',
    label: 'local',
    captured_at_ms: startMs + 3_000,
  },
});
assert.equal(localSimulation.status, 'local_simulation_axis');
assert.equal(localSimulation.accepted_for_realtime, false);

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
assert.equal(kit.platformAnnotationIntakePlan('zoom').schema, 'meeting_platform_annotation_intake_plan');
assert.equal(kit.platformAnnotationIntakeMatrix({ platforms: ['zoom'] }).platform_count, 1);
assert.equal(kit.platformAnnotationIntake('google-meet', {
  current_meeting: currentMeeting,
  annotation: { id: 'kit-note', label: 'kit', captured_at_ms: startMs + 4_000 },
}).status, 'ready_to_insert_current_axis');
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_annotation_intake_matrix.platform_count, 1);

console.log('ok meeting platform annotation intake');
