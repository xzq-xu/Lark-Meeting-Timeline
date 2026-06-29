import assert from 'node:assert/strict';
import {
  isDemoMeetingAxis,
  isLocalSimulationAxis,
  isCurrentPreparedRealAxis,
  isRealMeetingAxisClient,
  shouldHideDemoTimelineForProbe,
  sourceLabelForMeeting,
} from '../public/uiState.mjs';

const demoState = {
  meeting: {
    meeting_id: 'demo-lark-meeting-001',
    source: null,
    pending_binding: false,
  },
};

assert.equal(isDemoMeetingAxis(demoState.meeting), true);
assert.equal(shouldHideDemoTimelineForProbe({
  probe: { active: true },
  state: demoState,
}), true);

assert.equal(shouldHideDemoTimelineForProbe({
  probe: { active: false, status: 'timeout', started_at: '2026-06-26T12:00:00.000Z' },
  state: demoState,
}), true);

assert.equal(shouldHideDemoTimelineForProbe({
  probe: {
    active: false,
    status: 'passed',
    started_at: '2026-06-26T12:00:00.000Z',
    observed_event: { event_type: 'vc.meeting.all_meeting_started_v1' },
  },
  state: demoState,
}), false);

assert.equal(shouldHideDemoTimelineForProbe({
  probe: { active: false },
  state: demoState,
}), false);

assert.equal(shouldHideDemoTimelineForProbe({
  probe: { active: false },
  state: {
    ...demoState,
    presentation: {
      hide_timeline: true,
      hidden_reason: 'demo_sample_axis_is_not_current_real_demo',
    },
  },
}), true);

assert.equal(shouldHideDemoTimelineForProbe({
  probe: { active: false },
  state: demoState,
  realDemoSession: { active: true, prepared_at: '2026-06-26T12:00:00.000Z' },
}), true);

const localSimulationState = {
  meeting: {
    meeting_id: 'local-live-001',
    source: 'local_simulation',
    pending_binding: false,
  },
};
assert.equal(isLocalSimulationAxis(localSimulationState.meeting), true);
assert.equal(shouldHideDemoTimelineForProbe({
  probe: { active: false },
  state: localSimulationState,
  realDemoSession: { active: true, prepared_at: '2026-06-26T12:00:00.000Z' },
}), true);

const realWsState = {
  meeting: {
    meeting_id: 'real-meeting-001',
    source: 'lark_ws_event',
    pending_binding: false,
  },
};
assert.equal(isRealMeetingAxisClient(realWsState.meeting), true);
assert.equal(shouldHideDemoTimelineForProbe({
  probe: { active: true },
  state: realWsState,
}), false);

assert.equal(isCurrentPreparedRealAxis({
  meeting: realWsState.meeting,
  realDemoSession: {
    active: true,
    prepared_at: '2026-06-26T12:00:00.000Z',
    last_real_axis_at: '2026-06-26T12:00:01.000Z',
  },
}), true);

assert.equal(isCurrentPreparedRealAxis({
  meeting: realWsState.meeting,
  realDemoSession: {
    active: true,
    prepared_at: '2026-06-26T12:00:00.000Z',
    last_real_axis_at: '2026-06-26T11:59:59.000Z',
  },
}), false);

assert.equal(shouldHideDemoTimelineForProbe({
  probe: { active: false },
  state: realWsState,
  realDemoSession: {
    active: true,
    prepared_at: '2026-06-26T12:00:00.000Z',
    last_real_axis_at: null,
  },
}), true);

assert.equal(shouldHideDemoTimelineForProbe({
  probe: { active: false },
  state: realWsState,
  realDemoSession: {
    active: true,
    prepared_at: '2026-06-26T12:00:00.000Z',
    last_real_axis_at: '2026-06-26T12:00:01.000Z',
  },
}), false);

const pendingState = {
  meeting: {
    meeting_id: 'pending-live-001',
    source: 'annotation_fallback',
    pending_binding: true,
  },
};
assert.equal(isRealMeetingAxisClient(pendingState.meeting), false);
assert.equal(shouldHideDemoTimelineForProbe({
  probe: { active: true },
  state: pendingState,
}), false);

assert.equal(shouldHideDemoTimelineForProbe({
  probe: { active: false },
  state: pendingState,
  realDemoSession: { active: true, prepared_at: '2026-06-26T12:00:00.000Z' },
}), true);

assert.equal(sourceLabelForMeeting({
  source: 'lark_tenant_meeting_search_api',
  pending_binding: false,
}), '飞书租户会议扫描轴');

assert.equal(sourceLabelForMeeting({
  source: 'open_meeting_session',
  pending_binding: false,
}), '开放会议会话轴');

assert.equal(isRealMeetingAxisClient({
  meeting_id: 'open-session-001',
  source: 'open_meeting_session',
  pending_binding: false,
}), true);

console.log('ok frontend timeline gate state');
