export function sourceLabelForMeeting(meeting = {}) {
  if (meeting.pending_binding) return '等待真实飞书事件';
  const labels = {
    local_simulation: '本地模拟轴',
    open_meeting_session: '开放会议会话轴',
    lark_ws_event: '飞书长连接事件轴',
    lark_http_event: '飞书 HTTP 事件轴',
    lark_http_local_event: '本机 HTTP 模拟轴',
    lark_event: '飞书事件轴',
    lark_reserve_api: '飞书预约轴',
    lark_reserve_pending: '飞书会议待开始',
    lark_active_meeting_api: '飞书活跃会议轴',
    lark_meeting_search_api: '飞书会议扫描轴',
    lark_tenant_meeting_search_api: '飞书租户会议扫描轴',
    lark_meeting_lookup_api: '飞书会议号绑定轴',
    lark_probe_auto_search: '飞书 probe 自动扫描轴',
    lark_passive_meeting_scan: '飞书被动扫描轴',
    annotation_fallback: '等待真实飞书事件',
  };
  return labels[meeting.source] ?? meeting.source ?? '未标记来源';
}

export function isRealMeetingAxisClient(meeting = {}) {
  return [
    'open_meeting_session',
    'lark_ws_event',
    'lark_http_event',
    'lark_active_meeting_api',
    'lark_meeting_search_api',
    'lark_tenant_meeting_search_api',
    'lark_tenant_passive_meeting_scan',
    'lark_meeting_lookup_api',
    'lark_probe_auto_search',
    'lark_passive_meeting_scan',
  ].includes(meeting.source) && !meeting.pending_binding;
}

export function isDemoMeetingAxis(meeting = {}) {
  return meeting.meeting_id === 'demo-lark-meeting-001' && !meeting.source && !meeting.pending_binding;
}

export function isLocalSimulationAxis(meeting = {}) {
  return meeting.source === 'local_simulation' && !meeting.pending_binding;
}

function isoMs(value) {
  if (value == null || value === '') return null;
  const ms = Date.parse(value);
  return Number.isNaN(ms) ? null : ms;
}

export function isCurrentPreparedRealAxis({ meeting, realDemoSession } = {}) {
  if (!isRealMeetingAxisClient(meeting)) return false;
  if (!realDemoSession?.active) return true;

  const preparedMs = isoMs(realDemoSession.prepared_at);
  if (preparedMs == null) return true;

  const axisMs = isoMs(realDemoSession.last_real_axis_at);
  return axisMs != null && axisMs >= preparedMs;
}

export function shouldHideDemoTimelineForProbe({ probe, state, realDemoSession } = {}) {
  if (state?.presentation?.hide_timeline === true) return true;
  const probeStartedWithoutEvent = Boolean(
    probe?.started_at
      && probe.status !== 'idle'
      && probe.status !== 'passed'
      && !probe.observed_event,
  );
  const hasMeeting = Boolean(state?.meeting);
  const meeting = state?.meeting ?? {};
  const realDemoWaiting = Boolean(
    realDemoSession?.active
      && !isCurrentPreparedRealAxis({ meeting, realDemoSession }),
  );
  if (realDemoWaiting && hasMeeting) return true;

  const shouldHideForRealDemo = Boolean(
    realDemoWaiting
      && meeting
      && !meeting.pending_binding,
  );
  const hiddenWhileWaiting = realDemoWaiting
    ? (shouldHideForRealDemo || isDemoMeetingAxis(meeting) || isLocalSimulationAxis(meeting))
    : isDemoMeetingAxis(meeting);
  return Boolean(
    (probe?.active || probeStartedWithoutEvent || realDemoWaiting)
      && hasMeeting
      && hiddenWhileWaiting
      && !isCurrentPreparedRealAxis({ meeting, realDemoSession }),
  );
}
