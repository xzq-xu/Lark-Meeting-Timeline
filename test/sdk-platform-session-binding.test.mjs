import assert from 'node:assert/strict';

import {
  buildMeetingPlatformSessionBinding,
  buildMeetingPlatformSessionBindingMatrix,
  buildMeetingPlatformSessionBindingPlan,
} from '../packages/meeting-timeline-sdk/adapters/platform-session-binding.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const startMs = 1_782_614_400_000;

const plan = buildMeetingPlatformSessionBindingPlan('google-meet');
assert.equal(plan.schema, 'meeting_platform_session_binding_plan');
assert.equal(plan.platform, 'google_meet');
assert.equal(plan.match_policy.accept_score, 70);
assert.equal(plan.realtime_policy.local_observer_preferred, true);
assert.equal(plan.realtime_policy.provider_events_block_realtime, false);
assert.equal(plan.realtime_policy.transcript_blocks_realtime, false);

const matrix = buildMeetingPlatformSessionBindingMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(matrix.schema, 'meeting_platform_session_binding_matrix');
assert.equal(matrix.platform_count, 5);
assert.equal(matrix.provider_blocking_count, 0);
assert.equal(matrix.transcript_blocking_count, 0);

const bound = buildMeetingPlatformSessionBinding('google-meet', {
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
  signals: [
    {
      type: 'meeting_started',
      meeting: {
        platform: 'google_meet',
        meeting_id: 'abc-defg-hij',
        meeting_url: 'https://meet.google.com/abc-defg-hij',
      },
      occurred_at_ms: startMs,
    },
  ],
});
assert.equal(bound.schema, 'meeting_platform_session_binding');
assert.equal(bound.status, 'bound_to_current_axis');
assert.equal(bound.bind_to_current_axis, true);
assert.equal(bound.accepted_for_realtime, true);
assert.equal(bound.selected_meeting.meeting_id, 'abc-defg-hij');
assert.equal(bound.positive_match_count, 2);
assert.equal(bound.conflict_count, 0);
assert.equal(bound.matches.some((match) => match.evidence.includes('meeting_id_match')), true);

const openFromLocal = buildMeetingPlatformSessionBinding('zoom', {
  local_observer: {
    url: 'https://zoom.us/j/987654321?pwd=secret',
    title: 'Weekly review - Zoom',
    observed_at_ms: startMs + 2_000,
  },
});
assert.equal(openFromLocal.status, 'open_axis_from_local_observer');
assert.equal(openFromLocal.should_start_axis, true);
assert.equal(openFromLocal.start_payload.platform, 'zoom');
assert.equal(openFromLocal.start_payload.meeting_id, '987654321');
assert.equal(openFromLocal.start_payload.detector_source, 'local_observer_session_binding');
assert.equal(openFromLocal.start_payload.suppress_auto_annotations, true);

const openFromProvider = buildMeetingPlatformSessionBinding('webex', {
  signals: [
    {
      type: 'meeting_started',
      meeting: {
        platform: 'webex',
        meeting_id: 'webex-meeting-001',
        meeting_url: 'https://example.webex.com/meet/ada',
      },
      occurred_at_ms: startMs + 3_000,
    },
  ],
});
assert.equal(openFromProvider.status, 'open_axis_from_provider_start');
assert.equal(openFromProvider.start_payload.detector_source, 'provider_event_session_binding');

const conflict = buildMeetingPlatformSessionBinding('google-meet', {
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
});
assert.equal(conflict.status, 'binding_conflict');
assert.equal(conflict.accepted_for_realtime, false);
assert.equal(conflict.conflict_count, 1);
assert.equal(conflict.conflicts[0].evidence.includes('meeting_id_conflict'), true);
assert.equal(conflict.next_actions.includes('do_not_switch_current_axis_automatically'), true);

const pending = buildMeetingPlatformSessionBinding('lark', {
  annotation: {
    id: 'note-1',
    label: 'why?',
    captured_at_ms: startMs + 4_000,
  },
});
assert.equal(pending.status, 'pending_binding');
assert.equal(pending.should_store_pending, true);
assert.equal(pending.accepted_for_realtime, true);

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
assert.equal(kit.platformSessionBindingPlan('zoom').schema, 'meeting_platform_session_binding_plan');
assert.equal(kit.platformSessionBindingMatrix({ platforms: ['zoom'] }).platform_count, 1);
assert.equal(kit.platformSessionBinding('zoom', {
  local_observer: {
    url: 'https://zoom.us/j/987654321',
    observed_at_ms: startMs,
  },
}).status, 'open_axis_from_local_observer');
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_session_binding_matrix.platform_count, 1);

console.log('ok meeting platform session binding');
