import {
  isRealMeetingAxisClient,
  shouldHideDemoTimelineForProbe as shouldHideDemoTimelineForProbeState,
  sourceLabelForMeeting,
} from './uiState.mjs';

let state = null;
let stream = null;
let larkConfig = null;
let readiness = null;
let realMeetingProbe = null;
let deliveryDiagnostics = null;
let acceptanceReport = null;
let realDemoStatus = null;
let realDemoProgress = null;
let realDemoAcceptance = null;
let annotationIngestInfo = null;
let autoAcceptance = null;
let deviceSimulator = null;
let passiveMeetingScan = null;
let realDemoSession = null;
let probeAutoBindInFlight = false;
let realDemoMonitor = null;
let realDemoMonitorInFlight = false;
let realDemoProgressStream = null;

const $ = (id) => document.getElementById(id);

function fmtTime(ms) {
  const total = Math.max(0, Math.round(ms / 1000));
  const m = Math.floor(total / 60);
  const s = total % 60;
  return `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
}

function escapeHtml(text) {
  return String(text ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

function fmtAbsoluteTime(value) {
  if (value == null || value === '') return '';
  const date = new Date(Number(value) || value);
  return Number.isNaN(date.getTime()) ? String(value) : date.toISOString();
}

function fmtLocalClock(value) {
  if (value == null || value === '') return '';
  const date = new Date(Number(value) || value);
  if (Number.isNaN(date.getTime())) return String(value);
  return date.toLocaleTimeString('zh-CN', {
    hour12: false,
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  });
}

function startTimeSourceLabel(meeting = {}) {
  const labels = {
    lark_payload_start_time: '飞书事件开始时间',
    lark_record_start_time: '飞书会议记录开始时间',
    event_create_time: '飞书事件投递时间',
    manual_start_override: '人工指定开始时间',
    operator_supplied_start_time: '人工/检测器指定开始时间',
    session_opened_at: '会话打开时刻',
    current_meeting: '沿用当前轴零点',
    server_now_fallback: '服务端当前时间兜底',
  };
  return labels[meeting.start_time_source] ?? meeting.start_time_source ?? '未标注来源';
}

function axisSourceIsMeetingStartEvent(source) {
  return ['lark_ws_event', 'lark_http_event'].includes(source);
}

function axisSourceIsFallback(source) {
  return [
    'lark_active_meeting_api',
    'lark_meeting_search_api',
    'lark_tenant_meeting_search_api',
    'lark_tenant_passive_meeting_scan',
    'lark_meeting_lookup_api',
    'lark_probe_auto_search',
    'lark_passive_meeting_scan',
  ].includes(source);
}

function axisSourceIsOpenSession(source) {
  return source === 'open_meeting_session';
}

function timelineAxisBanner() {
  const meeting = state?.meeting ?? {};
  if (!isRealMeetingAxisClient(meeting)) return '';
  const source = meeting.source ?? 'unknown';
  if (axisSourceIsMeetingStartEvent(source)) {
    return `
      <div class="timeline-axis-banner ok">
        轴来源：${escapeHtml(sourceLabelForMeeting(meeting))}，由飞书会议开始/结束事件维护。
      </div>
    `;
  }
  if (axisSourceIsOpenSession(source)) {
    return `
      <div class="timeline-axis-banner ok">
        轴来源：${escapeHtml(sourceLabelForMeeting(meeting))}。这是产品主路径的本地会议会话入口；严格飞书事件验收仍需 meeting_start 官方事件。
      </div>
    `;
  }
  if (axisSourceIsFallback(source)) {
    return `
      <div class="timeline-axis-banner warn">
        轴来源：${escapeHtml(sourceLabelForMeeting(meeting))}。这是扫描/绑定兜底，不是飞书 meeting_start 事件建轴；只能证明真实会议已绑定，不能证明会议开始时自动建轴。
      </div>
    `;
  }
  return `
    <div class="timeline-axis-banner warn">
      轴来源：${escapeHtml(sourceLabelForMeeting(meeting))}。当前不能证明它来自飞书会议开始事件。
    </div>
  `;
}

function currentMeetingEndOffsetMs() {
  const meeting = state?.meeting ?? {};
  const start = Date.parse(meeting.start_time ?? '');
  const end = Date.parse(meeting.end_time ?? '');
  if (Number.isFinite(start) && Number.isFinite(end)) return Math.max(0, end - start);

  const eventOffsets = (state?.events ?? [])
    .filter((event) => event?.type === 'meeting_end')
    .map((event) => Number(event.time_ms))
    .filter((value) => Number.isFinite(value) && value >= 0);
  return eventOffsets.length ? Math.min(...eventOffsets) : null;
}

function currentPreparedRealAxisActive() {
  if (state?.presentation?.real_axis_active) return true;
  if (shouldHideDemoTimelineForProbe()) return false;
  return isRealMeetingAxisClient(state?.meeting);
}

function currentVisibleMeetingEnded() {
  return Boolean(state?.meeting?.end_time) && !shouldHideDemoTimelineForProbe();
}

function itemAfterMeetingEndMs(item = {}) {
  const stored = Number(item.payload?.timing?.after_meeting_end_ms ?? 0);
  const endOffset = currentMeetingEndOffsetMs();
  const computed = endOffset == null ? 0 : Math.max(0, Number(item.time_ms ?? 0) - endOffset);
  return Math.max(Number.isFinite(stored) ? stored : 0, computed);
}

function timeSourceInfo(item = {}) {
  const timing = item.payload?.timing ?? {};
  const source = item.time_source || timing.source || 'unknown';
  const labels = {
    explicit_time: '显式相对时间',
    captured_at: '采集时间',
    stroke_point_time: 'stroke 时间',
    server_received_at: '服务端收到时间',
    unknown: '未知时间源',
  };
  const afterEndMs = itemAfterMeetingEndMs(item);
  const warning = source === 'server_received_at' || afterEndMs > 0;
  const afterEnd = afterEndMs > 0;
  const detail = [
    `时间源：${labels[source] ?? source}`,
    timing.source_field ? `字段：${timing.source_field}` : '',
    timing.normalized_time_ms != null ? `会议内：${fmtTime(timing.normalized_time_ms)}` : '',
    timing.captured_at_ms != null ? `采集：${fmtAbsoluteTime(timing.captured_at_ms)}` : '',
    timing.server_received_at_ms != null ? `收到：${fmtAbsoluteTime(timing.server_received_at_ms)}` : '',
    timing.server_receive_delay_ms != null ? `上传延迟：${Math.round(timing.server_receive_delay_ms / 1000)}s` : '',
    afterEndMs > 0 ? `晚于会议结束：${fmtTime(afterEndMs)}` : '',
  ].filter(Boolean).join(' · ');
  return {
    source,
    label: afterEnd ? '晚于会议结束' : labels[source] ?? source,
    detail,
    warning,
    color: warning ? '#c46a13' : '#2563eb',
    tagClass: warning ? 'warn' : 'ok',
  };
}

function eventSourceLabel(source) {
  const labels = {
    lark_ws_event: '飞书长连接',
    lark_http_event: '飞书 HTTP',
    lark_http_local_event: '本机事件调试',
    lark_active_meeting_api: '活跃会议接口',
    lark_meeting_search_api: '手动扫描',
    lark_tenant_meeting_search_api: '租户扫描',
    lark_tenant_passive_meeting_scan: '租户被动扫描',
    lark_meeting_lookup_api: '会议号绑定',
    lark_probe_auto_search: 'probe 自动扫描',
    lark_passive_meeting_scan: '被动扫描',
    open_meeting_session: '开放会话',
    local_simulation: '本地模拟',
    local_pending: 'pending 临时轴',
    lark_event: '飞书事件',
  };
  return labels[source] ?? source ?? '未知来源';
}

function timelineMarkerMeta(item = {}, source = timeSourceInfo(item)) {
  const timing = item.payload?.timing ?? {};
  const parts = [source.label];
  if (timing.captured_at_ms != null) parts.push(`采集 ${fmtLocalClock(timing.captured_at_ms)}`);
  else if (timing.server_received_at_ms != null) parts.push(`收到 ${fmtLocalClock(timing.server_received_at_ms)}`);
  const afterEndMs = itemAfterMeetingEndMs(item);
  if (afterEndMs > 0) parts.push(`晚于结束 ${fmtTime(afterEndMs)}`);
  return parts.filter(Boolean).join(' · ');
}

function annotationBindingLabel(value) {
  const labels = {
    real_meeting_bound: '已绑定真实会议轴',
    pending_real_meeting: 'pending 等待真实会议事件',
    local_simulation: '本地模拟轴',
    demo_ignored: '样例轴已忽略',
    unbound: '未绑定',
  };
  return labels[value] ?? value ?? '未知';
}

function requirementLabel(id) {
  const labels = {
    event_receiver_ready: '链路已就绪',
    real_meeting_axis_active: '真实会议轴',
    event_meeting_axis_built: '飞书事件建轴',
    open_annotation_on_real_axis: '标注落真实轴',
    realtime_state_broadcast: '实时广播',
    transcript_post_meeting_only: '转写会后处理',
  };
  return labels[id] ?? id ?? '未知验收项';
}

function requirementSummary(requirements = []) {
  if (!Array.isArray(requirements) || !requirements.length) return '';
  return requirements
    .map((item) => `${requirementLabel(item.id)}=${item.ok ? 'OK' : 'TODO'}`)
    .join(' · ');
}

function realAxisSourceLabel(source) {
  const labels = {
    lark_ws_event: '飞书长连接事件建轴',
    lark_http_event: '飞书 HTTP webhook 建轴',
    lark_http_local_event: '本地 HTTP 事件调试建轴',
    lark_active_meeting_api: '活跃会议接口建轴',
    lark_meeting_search_api: '手动扫描建轴',
    lark_tenant_meeting_search_api: '租户扫描建轴',
    lark_tenant_passive_meeting_scan: '租户被动扫描建轴',
    lark_meeting_lookup_api: '会议号/链接绑定建轴',
    lark_probe_auto_search: 'probe 自动扫描建轴',
    lark_passive_meeting_scan: '被动扫描建轴',
    open_meeting_session: '开放会议会话建轴',
    local_simulation: '本地模拟轴',
    annotation_fallback: 'pending 标注临时轴',
    demo_sample: '样例轴',
  };
  return labels[source] ?? source ?? '尚未建轴';
}

function meetingSearchFallbackStatus(search = deliveryDiagnostics?.active_search) {
  const auth = larkConfig?.oauth ?? {};
  if (auth.meeting_search) {
    return {
      ready: Boolean(auth.meeting_search.usable),
      needsReauth: Boolean(auth.meeting_search.needs_reauth),
      detail: auth.meeting_search.usable
        ? '当前账号可在 probe 等待期间自动扫描近期真实会议作为建轴兜底'
        : auth.meeting_search.reason === 'oauth_token_expired'
          ? '飞书 OAuth token 已过期；重新授权后自动扫描兜底才能使用'
          : auth.meeting_search.reason === 'missing_scope'
            ? '当前 OAuth token 缺少 vc:meeting.search:read；重新授权后可自动扫描当前账号真实会议'
            : '尚未登录飞书账号；自动扫描兜底需要当前账号 OAuth',
    };
  }
  if (!auth.authenticated) {
    return {
      ready: false,
      needsReauth: true,
      detail: auth.token_present && auth.expired
        ? '飞书 OAuth token 已过期；重新授权后自动扫描兜底才能使用'
        : '尚未登录飞书账号；自动扫描兜底需要当前账号 OAuth',
    };
  }
  if (!search?.has_oauth_scope) {
    return {
      ready: false,
      needsReauth: true,
      detail: '当前 OAuth token 缺少 vc:meeting.search:read；重新授权后可自动扫描当前账号真实会议',
    };
  }
  return {
    ready: true,
    needsReauth: false,
    detail: '当前账号可在 probe 等待期间自动扫描近期真实会议作为建轴兜底',
  };
}

async function api(path, options = {}) {
  const response = await fetch(path, {
    ...options,
    headers: {
      'content-type': 'application/json; charset=utf-8',
      ...(options.headers ?? {}),
    },
  });
  const json = await response.json();
  if (!response.ok) {
    const error = new Error(json.error || `HTTP ${response.status}`);
    error.details = json;
    throw error;
  }
  return json;
}

function parseOffsetMs(value) {
  const text = String(value ?? '').trim();
  if (!text) return null;
  if (/^\d+(\.\d+)?$/.test(text)) return Math.round(Number(text));
  const match = text.match(/^(?:(\d{1,2}):)?(\d{1,2}):(\d{1,2})(?:[.,](\d{1,3}))?$/);
  if (!match) return null;
  const hours = Number(match[1] ?? 0);
  const minutes = Number(match[2]);
  const seconds = Number(match[3]);
  const millis = Number((match[4] ?? '').padEnd(3, '0') || 0);
  return ((hours * 60 + minutes) * 60 + seconds) * 1000 + millis;
}

function renderStats() {
  if (shouldHideDemoTimelineForProbe()) {
    $('stats').innerHTML = [
      ['转写片段', 0],
      ['会议事件', 0],
      ['实时标注', 0],
      ['建轴状态', '等待真实事件'],
    ].map(([label, value]) => `
      <div class="stat">
        <span>${label}</span>
        <strong>${escapeHtml(value)}</strong>
      </div>
    `).join('');
    return;
  }
  $('stats').innerHTML = [
    ['转写片段', state.segments.length],
    ['会议事件', state.events.length],
    ['实时标注', state.sequence.length],
    ['建轴来源', sourceLabelForMeeting(state.meeting)],
    ['已对齐', state.alignments.length],
  ].map(([label, value]) => `
    <div class="stat">
      <span>${label}</span>
      <strong>${value}</strong>
    </div>
  `).join('');
}

function shouldHideDemoTimelineForProbe() {
  const effectiveRealDemoSession = realDemoSession
    ?? realDemoStatus?.real_demo_session
    ?? realDemoProgress?.result?.real_demo_session
    ?? readiness?.real_demo_session
    ?? acceptanceReport?.current_evidence?.real_demo_session
    ?? null;
  return shouldHideDemoTimelineForProbeState({
    probe: realMeetingProbe,
    state,
    realDemoSession: effectiveRealDemoSession,
  });
}

function renderDebugInputControls() {
  const localLocked = Boolean(realDemoSession?.active);
  for (const id of ['startLiveBtn', 'endLiveBtn']) {
    const button = $(id);
    if (!button) continue;
    button.disabled = localLocked;
    button.title = localLocked
      ? '真实演示模式只等待飞书真实会议事件或当前用户会议扫描；本地模拟按钮不会建轴'
      : '仅用于本地模拟调试，不代表真实飞书会议';
  }

  const capturedAtInput = $('liveCapturedAtInput');
  const liveMarkBtn = $('liveMarkBtn');
  const realAxisReady = currentPreparedRealAxisActive();
  const visibleMeetingEnded = currentVisibleMeetingEnded();
  if (liveMarkBtn) {
    const missingCapturedAt = !capturedAtInput?.value.trim();
    liveMarkBtn.disabled = visibleMeetingEnded && missingCapturedAt;
    liveMarkBtn.title = liveMarkBtn.disabled
      ? '当前会议已结束；会后补传标注必须填写真实采集时间 captured_at/captured_at_ms'
      : realAxisReady
        ? '写入一条开放标注；留空采集时间时会使用当前提交时间'
        : '真实轴尚未建立；这里写入会成为 pending 标注，真实 meeting_start 到达后按 captured_at_ms 重算';
  }

  const acceptanceBtn = $('writeAcceptanceMarkBtn');
  if (acceptanceBtn) {
    acceptanceBtn.disabled = !realAxisReady || visibleMeetingEnded;
    acceptanceBtn.title = !realAxisReady
      ? '真实会议轴尚未建立；验收标注不能用按钮轴代替真实会议开始建轴'
      : visibleMeetingEnded
        ? '会议已结束；验收标注必须在真实会议轴进行中写入'
        : '通过开放标注接口写入一条验收标注';
  }
}

function renderMeta() {
  if (shouldHideDemoTimelineForProbe()) {
    const preparedAt = realDemoSession?.prepared_at ? ` · 准备时间 ${realDemoSession.prepared_at}` : '';
    const reason = state?.presentation?.hidden_reason ? ` · ${state.presentation.hidden_reason}` : '';
    $('meetingMeta').textContent = `等待真实飞书会议开始事件 · 当前尚未建轴 · 旧样例时间轴已隐藏${preparedAt}${reason}`;
    $('durationLabel').textContent = '尚未建轴';
    $('meetingStartInput').value = '';
    return;
  }
  const meeting = state.meeting;
  const suffix = meeting.meeting_url ? ` · ${meeting.meeting_url}` : '';
  const sourceCaveat = axisSourceIsFallback(meeting.source)
    ? '（扫描/绑定兜底，不是 meeting_start 事件）'
    : axisSourceIsOpenSession(meeting.source)
      ? '（开放会话入口，非飞书官方事件）'
    : '';
  const axisBuiltAt = realDemoSession?.last_real_axis_at
    ? ` · 建轴 ${realDemoSession.last_real_axis_source || meeting.source || 'unknown'} @ ${realDemoSession.last_real_axis_at}`
    : '';
  $('meetingMeta').textContent = `${sourceLabelForMeeting(meeting)}${sourceCaveat} · 轴零点 ${meeting.start_time}（${startTimeSourceLabel(meeting)}）${axisBuiltAt} · ${meeting.title} · ${meeting.meeting_id}${suffix}`;
  const realDuration = currentMeetingEndOffsetMs();
  const visualDuration = Number(state.duration_ms ?? 0);
  $('durationLabel').textContent = realDuration != null
    ? `会议时长 ${fmtTime(realDuration)}${visualDuration > realDuration ? ` · 显示范围 ${fmtTime(visualDuration)}` : ''}`
    : `显示范围 ${fmtTime(visualDuration)}`;
  $('meetingStartInput').value = meeting.start_time ?? '';
}

function timelineSvg() {
  const visualDurationMs = Math.max(1, Number(state.duration_ms ?? 0));
  const sequenceWithSource = state.sequence.map((item) => ({ item, source: timeSourceInfo(item) }));
  const calibratedSequence = sequenceWithSource.filter((entry) => !entry.source.warning);
  const uncalibratedSequence = sequenceWithSource.filter((entry) => entry.source.warning);
  const uncalibratedRows = Math.ceil(uncalibratedSequence.length / 2);
  const width = 1040;
  const height = uncalibratedSequence.length ? 430 + uncalibratedRows * 42 : 430;
  const left = 92;
  const right = 32;
  const laneWidth = width - left - right;
  const scale = (ms) => left + Math.max(0, Math.min(1, ms / visualDurationMs)) * laneWidth;
  const textPosition = (x, preferredGap = 8, reserve = 260) => (
    x > width - right - reserve
      ? { x: x - preferredGap, anchor: 'end' }
      : { x: x + preferredGap, anchor: 'start' }
  );
  const ticks = [];
  const step = [
    60 * 1000,
    2 * 60 * 1000,
    5 * 60 * 1000,
    10 * 60 * 1000,
    15 * 60 * 1000,
    30 * 60 * 1000,
    60 * 60 * 1000,
    2 * 60 * 60 * 1000,
  ].find((candidate) => visualDurationMs / candidate <= 8) ?? 4 * 60 * 60 * 1000;
  for (let ms = 0; ms <= visualDurationMs; ms += step) ticks.push(ms);
  if (ticks.at(-1) !== visualDurationMs) ticks.push(visualDurationMs);

  const segmentRects = state.segments.map((seg, index) => {
    const x = scale(seg.start_ms);
    const w = Math.max(4, scale(seg.end_ms) - x);
    const y = 92 + (index % 3) * 34;
    return `
      <rect x="${x}" y="${y}" width="${w}" height="22" rx="4" fill="#dbeafe" stroke="#93c5fd" />
      <text x="${x + 6}" y="${y + 15}" fill="#1e3a8a" font-size="11">${escapeHtml(seg.speaker_name)}</text>
    `;
  }).join('');

  const eventMarkers = state.events.map((event) => {
    const x = scale(event.time_ms);
    const label = textPosition(x, 7, 180);
    const color = event.type.includes('screen') ? '#c46a13' : event.type.includes('record') ? '#c43d3d' : '#0f8b62';
    const sourceLabel = eventSourceLabel(event.source);
    return `
      <line x1="${x}" y1="186" x2="${x}" y2="238" stroke="${color}" stroke-width="2" />
      <circle cx="${x}" cy="188" r="5" fill="${color}" />
      <text x="${label.x}" y="192" text-anchor="${label.anchor}" fill="${color}" font-size="11">${escapeHtml(event.label)}</text>
      <text x="${label.x}" y="205" text-anchor="${label.anchor}" fill="${color}" font-size="10" opacity="0.78">${escapeHtml(sourceLabel)}</text>
    `;
  }).join('');

  const sequenceMarkers = calibratedSequence.map(({ item, source }, index) => {
    const x = scale(item.time_ms);
    const label = textPosition(x);
    const row = index % 4;
    const tipY = 258 + row * 18;
    const baseY = 270 + row * 18;
    const labelY = 274 + row * 18;
    return `
      <g>
        <title>${escapeHtml(source.detail || item.label)}</title>
        <line x1="${x}" y1="${baseY}" x2="${x}" y2="354" stroke="${source.color}" stroke-width="2" />
        <path d="M ${x - 6} ${baseY} L ${x + 6} ${baseY} L ${x} ${tipY} Z" fill="${source.color}" />
        <text x="${label.x}" y="${labelY}" text-anchor="${label.anchor}" fill="${source.color}" font-size="12">${escapeHtml(item.label)}</text>
        <text x="${label.x}" y="${labelY + 13}" text-anchor="${label.anchor}" fill="${source.color}" font-size="10">${escapeHtml(timelineMarkerMeta(item, source))}</text>
      </g>
    `;
  }).join('');

  const uncalibratedMarkers = uncalibratedSequence.map(({ item, source }, index) => {
    const colWidth = Math.max(320, Math.floor(laneWidth / 2));
    const x = left + (index % 2) * colWidth;
    const y = 416 + Math.floor(index / 2) * 42;
    return `
      <g>
        <title>${escapeHtml(source.detail || item.label)}</title>
        <circle cx="${x}" cy="${y - 4}" r="4" fill="${source.color}" />
        <text x="${x + 10}" y="${y}" fill="${source.color}" font-size="12">未校准：${escapeHtml(item.label)}</text>
        <text x="${x + 10}" y="${y + 13}" fill="${source.color}" font-size="10">${escapeHtml(timelineMarkerMeta(item, source))}</text>
      </g>
    `;
  }).join('');

  const alignmentLines = state.alignments.map((alignment) => {
    const item = state.sequence.find((x) => x.id === alignment.sequence_id);
    const seg = state.segments.find((x) => x.id === alignment.active_segment_id);
    if (!item || !seg) return '';
    if (timeSourceInfo(item).warning) return '';
    const x1 = scale(item.time_ms);
    const x2 = scale(Math.min(Math.max(item.time_ms, seg.start_ms), seg.end_ms));
    return `<line x1="${x1}" y1="334" x2="${x2}" y2="126" stroke="#94a3b8" stroke-width="1.5" stroke-dasharray="4 4" />`;
  }).join('');

  return `
    <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="meeting timeline">
      <rect x="0" y="0" width="${width}" height="${height}" fill="#fcfcfd" />
      ${ticks.map((ms) => {
        const x = scale(ms);
        return `
          <line x1="${x}" y1="42" x2="${x}" y2="380" stroke="#e2e8f0" />
          <text x="${x - 17}" y="28" fill="#667085" font-size="11">${fmtTime(ms)}</text>
        `;
      }).join('')}
      <text x="18" y="108" fill="#475467" font-size="12" font-weight="700">转写</text>
      <text x="18" y="204" fill="#475467" font-size="12" font-weight="700">事件</text>
      <text x="18" y="292" fill="#475467" font-size="12" font-weight="700">标注</text>
      <line x1="${left}" y1="384" x2="${width - right}" y2="384" stroke="#98a2b3" stroke-width="2" />
      ${segmentRects}
      ${eventMarkers}
      ${alignmentLines}
      ${sequenceMarkers}
      ${uncalibratedSequence.length ? `<text x="18" y="416" fill="#c46a13" font-size="12" font-weight="700">未校准</text>` : ''}
      ${uncalibratedMarkers}
    </svg>
  `;
}

function renderTimeline() {
  if (shouldHideDemoTimelineForProbe()) {
    const pendingCount = state?.presentation?.pending_annotation_count ?? state?.sequence?.length ?? 0;
    const pendingText = pendingCount
      ? `已收到 ${pendingCount} 条待回绑标注；当前不画临时轴，避免误认为已经按会议开始时间建轴。`
      : '此时没有创建会议轴；收到飞书真实 meeting_start 事件后才会建轴。';
    $('timeline').innerHTML = `
      <div class="timeline-empty">
        <strong>等待真实会议开始事件</strong>
        <p class="hint">${escapeHtml(pendingText)}带 captured_at_ms 的标注如果先到，会进入 pending 队列，后续按真实 start_time 重算位置。</p>
      </div>
    `;
    return;
  }
  $('timeline').innerHTML = `${timelineAxisBanner()}${timelineSvg()}`;
}

function renderAlignments() {
  if (shouldHideDemoTimelineForProbe()) {
    if (state?.sequence?.length) {
      $('alignmentList').innerHTML = state.sequence.map((item) => {
        const source = timeSourceInfo(item);
        return `
          <article class="item">
            <div class="item-title">
              <span>${escapeHtml(item.label)}</span>
              <span class="tag">${escapeHtml(source.label)}</span>
            </div>
            <div class="item-meta">待回绑真实会议轴 · ${escapeHtml(source.detail || '时间源未知')}</div>
          </article>
        `;
      }).join('');
      return;
    }
    $('alignmentList').innerHTML = '<p class="hint">等待真实会议轴。样例标注已隐藏，避免和本次验收混淆。</p>';
    return;
  }
  if (!state.alignments.length) {
    $('alignmentList').innerHTML = '<p class="hint">还没有实时标注。会议中写入后会立即出现在这里。</p>';
    return;
  }
  $('alignmentList').innerHTML = state.alignments.map((alignment) => {
    const item = state.sequence.find((x) => x.id === alignment.sequence_id);
    const active = state.segments.find((x) => x.id === alignment.active_segment_id);
    const eventText = alignment.events.map((x) => `${x.event.label} ${fmtTime(x.event.time_ms)}`).join(' · ');
    const source = timeSourceInfo(item);
    return `
      <article class="item">
        <div class="item-title">
          <span>${escapeHtml(item?.label ?? alignment.label)}</span>
          <span class="tag ${source.tagClass}">${escapeHtml(source.label)}</span>
        </div>
        <div class="item-meta">${fmtTime(alignment.time_ms)} · ${escapeHtml(item?.kind ?? alignment.kind)}</div>
        <div class="item-meta">${escapeHtml(source.detail || '时间源未知')}</div>
        <div class="item-text">${active ? escapeHtml(active.text) : '未找到邻近转写'}</div>
        <div class="item-meta">${Math.round(alignment.confidence * 100)}% · ${escapeHtml(eventText || '无邻近会议事件')}</div>
      </article>
    `;
  }).join('');
}

function renderSegments() {
  if (shouldHideDemoTimelineForProbe()) {
    $('segmentList').innerHTML = '<p class="hint">等待真实会议结束后同步飞书妙记或导入转写。</p>';
    return;
  }
  if (!state.segments.length) {
    $('segmentList').innerHTML = '<p class="hint">暂无转写。会议结束后同步飞书妙记或导入转写 JSON 后，会自动补齐这里。</p>';
    return;
  }
  $('segmentList').innerHTML = state.segments.map((seg) => `
    <article class="item">
      <div class="item-title">
        <span>${escapeHtml(seg.speaker_name)}</span>
        <span>${fmtTime(seg.start_ms)}-${fmtTime(seg.end_ms)}</span>
      </div>
      <div class="item-text">${escapeHtml(seg.text)}</div>
    </article>
  `).join('');
}

function minuteTokenFromItem(item) {
  return item.minute_token
    || item.token
    || item.object_token
    || item.minute?.token
    || item.url
    || item.share_url
    || '';
}

function minuteTitleFromItem(item) {
  return item.title
    || item.topic
    || item.name
    || item.meeting_topic
    || item.minute?.title
    || '未命名妙记';
}

function renderMinuteResults(items, raw) {
  const box = $('minuteResults');
  if (!items?.length) {
    box.innerHTML = `<p class="hint">没有返回妙记。原始字段：${escapeHtml(JSON.stringify(Object.keys(raw?.data ?? raw ?? {})))}</p>`;
    return;
  }
  box.innerHTML = items.map((item, index) => {
    const token = minuteTokenFromItem(item);
    const title = minuteTitleFromItem(item);
    const time = item.start_time || item.create_time || item.created_at || item.minute?.create_time || '';
    return `
      <div class="mini-result">
        <button type="button" data-minute-token="${escapeHtml(token)}" data-minute-title="${escapeHtml(title)}">
          <div class="mini-title">${escapeHtml(title)}</div>
          <div class="mini-meta">${escapeHtml(time || `result ${index + 1}`)}</div>
          <div class="mini-meta">${escapeHtml(token || '未找到 token 字段')}</div>
        </button>
      </div>
    `;
  }).join('');
  box.querySelectorAll('button[data-minute-token]').forEach((button) => {
    button.addEventListener('click', () => {
      $('minuteTokenInput').value = button.dataset.minuteToken || '';
      if (button.dataset.minuteTitle) $('liveTitleInput').value = button.dataset.minuteTitle;
    });
  });
}

function canonicalizeLocalhostFromConfig(config) {
  if (window.location.hostname !== '127.0.0.1') return false;
  if (!config?.redirect_uri) return false;

  let redirectUrl;
  try {
    redirectUrl = new URL(config.redirect_uri);
  } catch {
    return false;
  }

  if (redirectUrl.hostname !== 'localhost') return false;
  if (redirectUrl.port && redirectUrl.port !== window.location.port) return false;

  const next = new URL(window.location.href);
  next.hostname = 'localhost';
  window.location.replace(next.toString());
  return true;
}

function renderWebhookStatus(config) {
  const urlInput = $('larkWebhookUrlInput');
  const status = $('larkWebhookStatus');
  if (!urlInput || !status) return;

  urlInput.value = config.event_callback_url || '';
  const ws = config.ws_event_receiver ?? {};
  const eventSummary = config.event_log_summary ?? {};
  const wsReady = ws.enabled && ws.state === 'connected';
  const callbackReady = Boolean(config.event_callback_public_https);
  const tokenText = config.verification_token_configured
    ? 'verification token 已配置'
    : 'verification token 未配置';
  if (wsReady) {
    const registered = Array.isArray(ws.registered_event_types) && ws.registered_event_types.length
      ? `已注册：${ws.registered_event_types.join(' / ')}。`
      : '';
    const counts = `原始 WS 事件 ${eventSummary.ws_event_count ?? config.recent_ws_event_count ?? 0} 条，建轴候选 ${eventSummary.ws_timeline_candidate_count ?? 0} 条，已建轴 ${eventSummary.ws_timeline_started_count ?? 0} 条。`;
    const last = ws.last_timeline_event_type
      ? `最近会议事件：${ws.last_timeline_event_type} @ ${ws.last_timeline_event_at}`
      : eventSummary.last_ws_event_type
        ? `最近原始事件：${eventSummary.last_ws_event_type} · ${eventSummary.last_ws_event_status || 'unknown'} @ ${eventSummary.last_ws_event_at}`
        : '等待真实会议事件';
    status.textContent = `长连接已连接：本机可直接接收飞书事件并自动建轴。${registered}${counts}${last}。`;
    renderLarkEventLog(config.recent_events ?? []);
    return;
  }
  const wsText = ws.enabled
    ? `长连接状态：${ws.state}${ws.error ? `，${ws.error}` : ''}`
    : `长连接未启用：${ws.state}`;
  status.textContent = callbackReady
    ? `${wsText}。HTTP webhook 可用：把该 URL 配到飞书开放平台事件订阅。${tokenText}。`
    : `${wsText}。HTTP webhook 仍是本地地址；如不用长连接，需要 ngrok/cloudflared 暴露 HTTPS 后配置 LARK_EVENT_CALLBACK_URL。${tokenText}。`;
  renderLarkEventLog(config.recent_events ?? []);
}

function renderLarkEventLog(events) {
  const box = $('larkEventLog');
  if (!box) return;
  if (!events.length) {
    box.innerHTML = '<p class="hint">暂无飞书事件。若长连接已连接但这里一直为空，说明开放平台还没有把事件推给这个应用。</p>';
    return;
  }
  box.innerHTML = events.map((event) => `
    <article class="event-log-row">
      <div class="item-title">
        <span>${escapeHtml(event.event_type)}</span>
        <span class="tag">${event.timeline_processed ? 'timeline' : event.timeline_candidate ? 'candidate' : 'ignored'}</span>
      </div>
      <div class="item-meta">${escapeHtml(event.transport)} · ${escapeHtml(event.at)}</div>
      <div class="item-meta">${escapeHtml([
        event.timeline_started ? 'started' : '',
        event.ignored_reason ? `reason=${event.ignored_reason}` : '',
        (event.parsed_keys || []).length ? `keys=${event.parsed_keys.join(', ')}` : '',
      ].filter(Boolean).join(' · '))}</div>
    </article>
  `).join('');
}

function renderRealMeetingProbe() {
  const box = $('realMeetingProbeStatus');
  if (!box || !realMeetingProbe) return;
  const elapsed = realMeetingProbe.elapsed_ms != null ? fmtTime(realMeetingProbe.elapsed_ms) : '00:00';
  const statusLabels = {
    idle: '未开始',
    waiting: '等待中',
    passed: '已通过',
    timeout: '已超时',
  };
  const eventText = realMeetingProbe.observed_event
    ? ` · ${realMeetingProbe.observed_event.event_type} @ ${realMeetingProbe.observed_event.at}`
    : '';
  const auto = realMeetingProbe.auto_search;
  const autoText = auto?.enabled
    ? ` · 自动扫描：${auto.last_result?.status || '等待'}${auto.last_result?.reason ? `/${auto.last_result.reason}` : ''}${auto.server_loop?.scheduled ? ' · 服务端轮询中' : ''}`
    : '';
  const missed = realMeetingProbe.missed_event;
  const missedText = missed
    ? ` · 最近真实开始事件早于本次等待 ${fmtTime(missed.start_before_probe_ms || 0)}${missed.ended_before_probe ? '，且会议已结束' : ''}`
    : '';
  box.textContent = `${statusLabels[realMeetingProbe.status] || realMeetingProbe.status} · ${elapsed} · ${realMeetingProbe.next_step}${eventText}${autoText}${missedText}`;
  const restoreBtn = $('restoreLatestRealAxisBtn');
  if (restoreBtn) {
    restoreBtn.disabled = !missed;
    restoreBtn.title = missed
      ? '用本地已收到的飞书会议开始事件恢复当前真实会议轴'
      : '只有当最近真实会议开始事件早于本次等待窗口时可用';
  }
}

function renderDeliveryDiagnostics() {
  const box = $('larkDeliveryDiagnostics');
  if (!box || !deliveryDiagnostics) return;
  const labels = {
    missing_app_credentials: '凭证缺失',
    receiver_not_ready: '接收未就绪',
    missing_local_event_handlers: '本地 handler 缺失',
    direct_meeting_event_received: '直开事件已收到',
    meeting_start_event_received: '会议开始事件已收到',
    meeting_context_event_received: '会议上下文事件已收到',
    real_events_seen_but_no_direct_start: '有真实事件但无直开开始',
    ws_events_seen_but_no_meeting_start: '有 WS 事件但无会议开始',
    waiting_for_lark_delivery: '等待飞书投递',
  };
  const auditLabels = {
    missing_app_credentials: '凭证缺失',
    receiver_not_ready: '接收未就绪',
    local_handler_missing: '本地 handler 缺失',
    parser_self_test_failed: '解析自检失败',
    event_delivery_ok: '事件投递正常',
    no_event_delivery_observed: '尚未观察到投递',
    wrong_event_subscription: '订阅事件不匹配',
    parser_payload_gap: '事件 payload 待适配',
    meeting_event_seen_but_no_axis: '会议事件未建轴',
  };
  const action = deliveryDiagnostics.next_actions?.[0]
    ? ` 下一步：${deliveryDiagnostics.next_actions[0]}`
    : '';
  const checklistNext = deliveryDiagnostics.open_platform_checklist?.find((item) => item.status !== 'ok');
  const checklistText = checklistNext
    ? ` 开放平台核查：${checklistNext.label} - ${checklistNext.required}`
    : '';
  const activeSearch = deliveryDiagnostics.active_search;
  const searchText = activeSearch
    ? ` 当前账号会议扫描：${activeSearch.has_oauth_scope ? 'OAuth scope 已包含 vc:meeting.search:read' : 'OAuth scope 缺少 vc:meeting.search:read'}。${activeSearch.api_status === 'invalid_token' ? '应用身份扫描已被飞书拒绝，建议走当前用户 OAuth。' : ''}`
    : '';
  const currentValidation = acceptanceReport?.current_validation;
  const annotationBinding = acceptanceReport?.current_evidence?.annotation_binding
    ?? deliveryDiagnostics.evidence?.annotation_binding;
  const evidence = deliveryDiagnostics.evidence ?? {};
  const eventCountsText = ` 原始WS=${evidence.ws_event_count ?? 0}，候选=${evidence.ws_timeline_candidate_count ?? 0}，已建轴=${evidence.ws_timeline_started_count ?? 0}。`;
  const audit = deliveryDiagnostics.real_meeting_event_audit;
  const auditText = audit
    ? ` 事件审计：${auditLabels[audit.status] || audit.status}；${audit.summary}`
    : '';
  const parserText = deliveryDiagnostics.parser_self_test
    ? ` 本地解析自检：${deliveryDiagnostics.parser_self_test.passed ? 'OK' : '失败'}。`
    : '';
  const validationText = currentValidation?.scoped_by_probe
    ? ` 本次验收：${currentValidation.real_event_after_probe ? '已收到 probe 后真实会议事件' : currentValidation.auto_search_binding_after_probe ? '已通过自动扫描绑定真实会议' : '尚未建立 probe 后真实会议入口'}，probe 后真实轴标注 ${currentValidation.real_axis_annotation_count_after_probe ?? 0} 条，当前标注绑定：${annotationBindingLabel(currentValidation.annotation_binding_state)}。`
    : annotationBinding?.total
      ? ` 标注绑定：${annotationBindingLabel(annotationBinding.binding_state)}，共 ${annotationBinding.total} 条。`
    : '';
  const nextStep = acceptanceReport?.acceptance_steps?.find((step) => !step.done);
  const reportText = nextStep ? ` 验收单下一项：${nextStep.text}` : '';
  box.textContent = `投递诊断：${labels[deliveryDiagnostics.status] || deliveryDiagnostics.status}。${deliveryDiagnostics.summary}${eventCountsText}${auditText}${parserText}${searchText}${validationText}${action}${checklistText}${reportText}`;
  renderAcceptanceGuide();
}

function readinessCheck(id) {
  return readiness?.checks?.find((check) => check.id === id) ?? null;
}

function renderAcceptanceGuide() {
  const box = $('acceptanceGuideSteps');
  const badge = $('acceptanceBadge');
  const action = $('acceptanceNextAction');
  if (!box || !badge || !action || !readiness) return;

  const wsCheck = readinessCheck('lark_ws_connected');
  const eventCheck = readinessCheck('direct_meeting_event_seen');
  const realAnnotationCheck = readinessCheck('real_annotation_seen');
  const apiCheck = readinessCheck('open_annotation_api');
  const realtimeCheck = readinessCheck('realtime_timeline');
  const search = deliveryDiagnostics?.active_search;
  const searchFallback = meetingSearchFallbackStatus(search);
  const passiveScan = passiveMeetingScan
    ?? readiness?.passive_meeting_scan
    ?? acceptanceReport?.current_evidence?.passive_meeting_scan
    ?? search?.passive_scan
    ?? null;
  const evidence = deliveryDiagnostics?.evidence ?? {};
  const binding = acceptanceReport?.current_evidence?.annotation_binding
    ?? evidence.annotation_binding
    ?? readiness.current?.annotation_binding
    ?? {};
  const currentValidation = acceptanceReport?.current_validation;
  const probe = realMeetingProbe ?? readiness.real_meeting_probe;
  const hasRealEvent = Number(evidence.real_event_count ?? 0) > 0;
  const hasStartEvent = Boolean(evidence.start_event || evidence.direct_start_event || evidence.reserve_start_event || evidence.context_start_event || eventCheck?.ok);
  const hasRealEntry = Boolean(currentValidation?.real_entry_after_probe || hasStartEvent || evidence.real_meeting_axis_active);
  const hasRealAnnotation = Boolean(realAnnotationCheck?.ok);
  const realAxisAnnotationCount = currentValidation?.scoped_by_probe
    ? Number(currentValidation.real_axis_annotation_count_after_probe ?? 0)
    : Number(binding.real_axis_count ?? 0);
  const pendingAnnotationCount = Number(binding.pending_count ?? 0);
  const completionEvidence = realDemoMonitor?.completion_evidence ?? realDemoProgress?.completion_evidence ?? null;
  const completionRequirements = completionEvidence?.requirements ?? [];
  const requirementText = requirementSummary(completionRequirements);
  const axisSource = completionEvidence?.meeting_source
    ?? realDemoStatus?.evidence?.meeting?.source
    ?? evidence.meeting?.source
    ?? readiness.current?.meeting?.source
    ?? state?.meeting?.source
    ?? null;
  const productPassed = Boolean(realDemoAcceptance?.product_acceptance_complete);
  const strictEventPassed = Boolean(realDemoAcceptance?.strict_event_acceptance_complete);
  const acceptanceSummary = $('realDemoAcceptanceSummary');
  if (acceptanceSummary) {
    const missing = realDemoAcceptance?.missing_product_requirements?.map((item) => item.id).join(', ') || '无';
    acceptanceSummary.textContent = realDemoAcceptance
      ? `验收摘要：产品验收=${productPassed ? '通过' : '未完成'} · 严格事件投递=${strictEventPassed ? '通过' : realDemoAcceptance.event_delivery_verified ? '已建事件轴但未完成标注' : '未证明'} · verdict=${realDemoAcceptance.verdict} · 推荐命令=${realDemoAcceptance.recommended_command || '无需继续'} · 缺失=${missing}`
      : '验收摘要加载中';
  }
  const statusSummary = $('realDemoStatusSummary');
  if (statusSummary) {
    const gates = realDemoStatus?.gates ?? {};
    const runbook = realDemoStatus?.operator_runbook;
    const blockerText = realDemoStatus?.blockers?.length
      ? ` 阻塞/提醒：${realDemoStatus.blockers.map((item) => item.text).slice(0, 2).join('；')}`
      : '';
    statusSummary.textContent = realDemoStatus
      ? `${realDemoStatus.summary} 状态=${realDemoStatus.status} · 阶段=${runbook?.phase || 'unknown'} · 下一步=${runbook?.primary_next_action || realDemoStatus.next_action || '等待'} · 事件接收=${gates.event_receiver_ready ? 'OK' : 'TODO'} · 真实事件=${gates.real_event_seen ? 'OK' : '未到'} · 事件审计=${gates.event_audit_status || 'unknown'} · 扫描兜底=${gates.scan_fallback_ready ? 'OK' : 'TODO'} · 虚拟设备=${gates.device_simulator_enabled ? 'OK' : 'TODO'} · 设备流=${gates.device_stream_enabled ? gates.device_stream_status || 'ON' : 'OFF'}。${blockerText}`
      : '真实演示状态加载中';
  }
  const monitorSummary = $('realDemoMonitorSummary');
  if (monitorSummary) {
    const progress = realDemoMonitor ?? realDemoProgress;
    const evidence = progress?.completion_evidence;
    const streamStatus = progress?.stream_status ?? realDemoStatus?.realtime_stream;
    monitorSummary.textContent = progress
      ? `完成证据：${progress.observed ? '已完成' : progress.timed_out ? '超时未完成' : '等待中'} · 建轴来源=${realAxisSourceLabel(evidence?.meeting_source)} · 真实轴=${evidence?.real_meeting_axis_active ? 'OK' : 'TODO'} · 真实轴标注=${evidence?.real_axis_annotation_count ?? 0} · 事件审计=${evidence?.event_audit_status || 'unknown'} · 设备流=${evidence?.device_stream_status || 'unknown'} ${evidence?.device_stream_count ?? 0}/${evidence?.device_stream_max_count ?? '-'} · 最近广播标注=${evidence?.last_broadcast_real_axis_annotation_count ?? streamStatus?.last_state_summary?.real_axis_annotation_count ?? 0} · SSE客户端=${evidence?.current_sse_clients ?? streamStatus?.current_clients ?? 0}${requirementText ? ` · 验收项：${requirementText}` : ''}`
      : `实时流证据：SSE客户端=${streamStatus?.current_clients ?? 0} · 广播=${streamStatus?.broadcast_count ?? 0} · 最近广播=${streamStatus?.last_broadcast_at ?? '无'}`;
  }

  const baseRows = [
    {
      label: '事件接收通道',
      ok: Boolean(wsCheck?.ok),
      detail: wsCheck?.detail ?? '等待长连接状态',
    },
    {
      label: '建轴来源',
      ok: Boolean(hasRealEntry),
      detail: hasRealEntry
        ? `${realAxisSourceLabel(axisSource)}；按钮只进入等待或扫描，不应把本地模拟轴当真实轴`
        : '尚未建真实会议轴；直接开启飞书会议后应由事件或明确扫描绑定创建',
    },
    {
      label: '直接开会建轴',
      ok: hasRealEntry,
      detail: hasStartEvent
        ? '真实会议开始事件已进入时间轴'
        : currentValidation?.auto_search_binding_after_probe
          ? '飞书事件未投递，但 probe 自动扫描已绑定真实会议轴'
          : evidence.real_meeting_axis_active
            ? '当前已有真实会议轴'
        : hasRealEvent
          ? `已有真实飞书事件 ${evidence.real_event_count} 条，但还没有直开会议开始事件`
          : '尚未收到真实飞书会议开始事件',
    },
    {
      label: '实时标注通道',
      ok: Boolean(apiCheck?.ok && realtimeCheck?.ok),
      detail: 'POST /api/annotations 写入后会通过 SSE 刷新统一时间轴',
    },
    {
      label: '真实会议标注',
      ok: hasRealAnnotation && realAxisAnnotationCount > 0,
      detail: realAxisAnnotationCount > 0
        ? `真实会议轴已有 ${realAxisAnnotationCount} 条开放标注`
        : pendingAnnotationCount > 0
          ? `已有 ${pendingAnnotationCount} 条标注在 pending 轴，等待飞书真实会议开始事件后重绑定`
          : realAnnotationCheck?.detail ?? '等待真实会议轴建立后写入一条标注',
    },
    {
      label: '虚拟墨水屏标注',
      ok: Boolean(deviceSimulator?.enabled),
      detail: deviceSimulator?.enabled
        ? `已开启；真实会议轴建立时自动写入 ${deviceSimulator.device_type || 'simulator'} 标注；设备流=${deviceSimulator.stream?.status || 'idle'} ${deviceSimulator.stream?.count ?? 0}/${deviceSimulator.stream?.max_count ?? '-'}`
        : '已关闭；真实开会后需要手动点“发送设备标注”或由外部设备调用 /api/annotations',
    },
    {
      label: '兜底扫描权限',
      ok: searchFallback.ready,
      detail: deliveryDiagnostics?.active_search?.api_status === 'invalid_token'
        ? `${searchFallback.detail}；应用身份扫描当前不可作为兜底`
        : searchFallback.detail,
    },
    {
      label: '被动扫描建轴',
      ok: Boolean(passiveScan?.enabled && (passiveScan?.last_result?.status === 'bound' || searchFallback.ready)),
      detail: passiveScan?.enabled
        ? passiveScan.last_result?.status === 'bound'
          ? `已通过被动扫描绑定真实会议：${passiveScan.last_result.meeting_title || passiveScan.last_result.selected_meeting_id || 'unknown'}`
          : searchFallback.ready
            ? `已开启服务端轮询；每 ${Math.round((passiveScan.server_loop?.interval_ms ?? passiveScan.interval_ms ?? 10000) / 1000)} 秒扫描当前账号进行中会议`
            : `已开启但不可用：${passiveScan.last_result?.reason || searchFallback.detail}`
        : '已关闭；不会在用户直接开会时主动扫描当前账号会议',
    },
    {
      label: 'probe 自动扫描兜底',
      ok: Boolean(currentValidation?.auto_search_binding_after_probe || searchFallback.ready),
      detail: currentValidation?.auto_search_binding_after_probe
        ? '本次 probe 已通过自动扫描绑定真实会议轴'
        : searchFallback.ready
        ? '启动验收探针后，页面会每 5 秒自动扫描一次当前账号近期会议'
          : '自动扫描不可用；主路径仍会等待飞书会议事件投递',
    },
    {
      label: '手动扫描兜底',
      ok: Boolean(deliveryDiagnostics?.evidence?.real_meeting_axis_active),
      detail: deliveryDiagnostics?.evidence?.real_meeting_axis_active
        ? '当前已绑定真实会议轴'
        : search?.has_oauth_scope
          ? '事件未投递时，可手动点击“扫描我的真实会议”绑定；不会在开始验收时自动建轴'
          : search?.api_status === 'invalid_token'
            ? '应用身份扫描当前不可用；请重新授权 vc:meeting.search:read 后用“扫描我的真实会议”'
          : '缺少 vc:meeting.search:read，只影响手动扫描兜底，不影响事件建轴主路径',
    },
  ];

  const requirementRows = completionRequirements.map((item) => ({
    label: `完成验收：${requirementLabel(item.id)}`,
    ok: Boolean(item.ok),
    detail: item.ok
      ? item.observed || item.required || '已满足'
      : `${item.required || '等待满足'} 当前：${item.observed || '暂无证据'}`,
  }));

  const rows = [...baseRows, ...requirementRows];

  box.innerHTML = rows.map((row) => `
    <div class="guide-step">
      <div>
        <div class="guide-step-title">${escapeHtml(row.label)}</div>
        <div class="hint">${escapeHtml(row.detail)}</div>
      </div>
      <span class="tag ${row.ok ? 'ok' : 'warn'}">${row.ok ? 'OK' : 'TODO'}</span>
    </div>
  `).join('');

  const passed = hasRealEntry && apiCheck?.ok && realtimeCheck?.ok && hasRealAnnotation;
  const waiting = probe?.status === 'waiting';
  const meetingEnded = currentVisibleMeetingEnded();
  const canWriteAcceptanceMark = !meetingEnded && currentPreparedRealAxisActive();
  const acceptanceMarkBtn = $('writeAcceptanceMarkBtn');
  if (acceptanceMarkBtn) {
    acceptanceMarkBtn.disabled = !canWriteAcceptanceMark;
    acceptanceMarkBtn.title = canWriteAcceptanceMark
      ? '通过开放标注接口写入一条验收标注'
      : meetingEnded
        ? '当前会议已经结束；会后补传必须由设备 payload 带真实 captured_at_ms，不能用浏览器当前时间伪造'
        : '等待真实会议轴；验收标注不能用 probe/pending/按钮轴代替真实会议开始建轴';
  }
  const autoBtn = $('toggleAutoAcceptanceBtn');
  if (autoBtn) {
    const enabled = Boolean(autoAcceptance?.enabled);
    autoBtn.textContent = enabled ? '自动验收标注：开' : '自动验收标注：关';
    autoBtn.className = enabled ? 'primary' : '';
    autoBtn.title = enabled
      ? '真实会议轴建立后会自动通过开放标注接口写入一条验收标注'
      : '开启后，直接开会建轴时自动写入一条验收标注用于端到端验证';
  }
  const simulatorBtn = $('toggleDeviceSimulatorBtn');
  if (simulatorBtn) {
    const enabled = Boolean(deviceSimulator?.enabled);
    simulatorBtn.textContent = enabled ? '虚拟设备标注：开' : '虚拟设备标注：关';
    simulatorBtn.className = enabled ? 'primary' : '';
    simulatorBtn.title = enabled
      ? '真实会议轴建立后会自动写入一条模拟墨水屏设备标注'
      : '开启后，真实会议轴建立时自动写入一条设备来源标注，便于验证会中实时标注';
  }
  const passiveBtn = $('togglePassiveScanBtn');
  if (passiveBtn) {
    const enabled = Boolean(passiveScan?.enabled);
    passiveBtn.textContent = enabled ? '被动扫描建轴：开' : '被动扫描建轴：关';
    passiveBtn.className = enabled ? 'primary' : '';
    passiveBtn.title = enabled
      ? '服务端会用当前 OAuth 用户定期扫描正在进行的真实会议，作为飞书事件未投递时的兜底建轴'
      : '关闭后只等待飞书事件投递或手动扫描/绑定';
  }
  const reauthBtn = $('reauthMeetingScopeBtn');
  if (reauthBtn) {
    const oauthRedirect = realDemoStatus?.auth_start?.redirect_url;
    reauthBtn.disabled = !searchFallback.needsReauth && !oauthRedirect;
    reauthBtn.dataset.redirectUrl = oauthRedirect || '';
    reauthBtn.title = oauthRedirect
      ? '打开飞书 OAuth，授权 vc:meeting.search:read 以启用当前账号会议扫描兜底'
      : searchFallback.needsReauth
        ? searchFallback.detail
      : '当前账号已具备会议扫描权限';
  }
  const monitorBtn = $('monitorRealDemoBtn');
  if (monitorBtn) {
    monitorBtn.disabled = realDemoMonitorInFlight;
    monitorBtn.className = realDemoMonitorInFlight ? 'primary' : '';
    monitorBtn.textContent = realDemoMonitorInFlight ? '正在等待证据' : '等待完成证据';
    monitorBtn.title = '等待真实会议轴、真实轴开放标注和 SSE 广播证据';
  }
  badge.className = `tag ${productPassed || passed ? 'ok' : 'warn'}`;
  badge.textContent = productPassed ? (strictEventPassed ? '事件验收通过' : '产品验收通过') : passed ? '可验收' : hasRealEntry ? '待写标注' : waiting ? '等待事件' : '待验证';
  action.textContent = productPassed
    ? strictEventPassed
      ? '真实会议事件建轴、实时标注和 SSE 广播都已验证；转写可在会议结束后同步。'
      : '真实会议轴和实时标注已完成产品验收；严格飞书事件投递仍需单独确认。'
    : passed
    ? '真实会议建轴和实时标注都已验证；转写可在会议结束后同步。'
    : hasRealEntry
        ? '真实会议轴已建立，点击“写入验收标注”验证标注实时进入时间轴。'
    : waiting
        ? searchFallback.ready
          ? '保持此页面打开，然后在飞书客户端直接开启一次会议；事件未投递时会自动扫描真实会议兜底。'
          : '保持此页面打开，然后在飞书客户端直接开启一次会议；自动扫描兜底需先授权会议扫描兜底。'
      : searchFallback.needsReauth
          ? '建议先授权会议扫描兜底；也可以保持服务运行后直接开会，等待飞书事件投递。'
          : '保持服务运行后直接在飞书客户端开启会议；“启动验收探针”只是可选 probe，不会建轴。';
}

function renderReadiness() {
  if (!readiness) return;
  const summary = $('readinessSummary');
  const list = $('readinessList');
  if (!summary || !list) return;
  summary.textContent = readiness.ready
    ? '真实链路可验收'
    : readiness.ready_for_local_demo
      ? '本地链路可跑，真实飞书入口待确认'
      : '存在阻塞项';
  list.innerHTML = readiness.checks.map((check) => `
    <article class="readiness-row">
      <span class="tag ${check.ok ? 'ok' : 'warn'}">${check.ok ? 'OK' : 'TODO'}</span>
      <div>
        <strong>${escapeHtml(check.label)}</strong>
        <p class="hint">${escapeHtml(check.detail)}</p>
        ${check.action_url ? `<a href="${escapeHtml(check.action_url)}" target="_blank" rel="noreferrer">打开配置页</a>` : ''}
      </div>
    </article>
  `).join('');
  renderAcceptanceGuide();
}

function renderAll() {
  renderMeta();
  renderStats();
  renderTimeline();
  renderAlignments();
  renderSegments();
  renderDebugInputControls();
}

function renderDeviceIngestInfo() {
  const status = $('deviceIngestStatus');
  if (!status || !annotationIngestInfo) return;
  const route = annotationIngestInfo.annotation_route ?? {};
  const meeting = annotationIngestInfo.current_meeting ?? {};
  const ignored = annotationIngestInfo.ignored_current_meeting;
  const session = annotationIngestInfo.real_demo_session ?? realDemoSession;
  const simulator = annotationIngestInfo.device_simulator ?? deviceSimulator;
  const meetingSearch = annotationIngestInfo.meeting_search_oauth;
  const streamSim = simulator?.stream ?? {};
  const routeText = route.mode ? `落轴：${route.mode}` : '落轴：未知';
  const meetingText = meeting.meeting_id
    ? `会议：${meeting.title || meeting.meeting_id}`
    : ignored?.meeting_id
      ? `当前样例轴已忽略：${ignored.meeting_id}`
      : '当前没有可见会议轴';
  const sessionText = session?.active ? '真实演示等待中' : '真实演示未准备';
  const authText = meetingSearch?.usable
    ? '扫描兜底可用'
    : meetingSearch?.reason === 'oauth_token_expired'
      ? '扫描兜底不可用：OAuth 已过期'
      : meetingSearch?.reason === 'missing_scope'
        ? '扫描兜底不可用：缺少 vc:meeting.search:read'
        : '扫描兜底不可用';
  const simulatorText = simulator?.enabled ? '虚拟设备会自动写入' : '虚拟设备关闭';
  const streamText = streamSim.enabled || streamSim.timer_active
    ? `流式模拟：${streamSim.status || 'running'} ${streamSim.count ?? 0}/${streamSim.max_count ?? '-'}`
    : streamSim.status && streamSim.status !== 'idle'
      ? `流式模拟：${streamSim.status}`
      : '流式模拟关闭';
  status.textContent = `${routeText} · ${meetingText} · ${sessionText} · ${simulatorText} · ${streamText} · ${authText} · endpoint=${annotationIngestInfo.endpoint}`;
  const startStreamBtn = $('startDeviceStreamBtn');
  if (startStreamBtn) {
    const running = Boolean(streamSim.enabled || streamSim.timer_active);
    startStreamBtn.disabled = running;
    startStreamBtn.className = running ? 'primary' : '';
  }
  const stopStreamBtn = $('stopDeviceStreamBtn');
  if (stopStreamBtn) {
    stopStreamBtn.disabled = !(streamSim.enabled || streamSim.timer_active);
  }
  renderDeviceCurl();
}

function buildDeviceDebugPayload(info = annotationIngestInfo) {
  const base = info?.minimal_payload ?? {
    id: 'device-mark-001',
    kind: 'handwriting_trigger',
    label: 'why?',
    text_candidates: ['why?', 'why'],
    strokes: [],
  };
  const now = Date.now();
  return {
    ...base,
    id: `device-debug-${now}`,
    captured_at_ms: now,
    realtime: true,
    mode: 'realtime',
    payload: {
      ...(base.payload ?? {}),
      origin: 'browser_device_debug',
      realtime: true,
      mode: 'realtime',
    },
  };
}

function isBrowserDeviceDebugPayload(payload = {}) {
  return payload?.payload?.origin === 'browser_device_debug'
    || /^device-(debug|mark)-/.test(String(payload.id ?? ''));
}

function refreshBrowserDeviceDebugPayload(payload = {}) {
  if (!isBrowserDeviceDebugPayload(payload)) return payload;
  const now = Date.now();
  return {
    ...payload,
    id: `device-debug-${now}`,
    captured_at_ms: now,
    realtime: true,
    mode: 'realtime',
    payload: {
      ...(payload.payload ?? {}),
      origin: 'browser_device_debug',
      realtime: true,
      mode: 'realtime',
      refreshed_at_ms: now,
    },
  };
}

function readDevicePayload({ refreshDebug = false } = {}) {
  const input = $('devicePayloadInput');
  const payload = JSON.parse(input?.value || '{}');
  const normalized = refreshDebug ? refreshBrowserDeviceDebugPayload(payload) : payload;
  if (refreshDebug && normalized !== payload && input) {
    input.value = JSON.stringify(normalized, null, 2);
  }
  return normalized;
}

function shellSingleQuote(value) {
  return `'${String(value).replaceAll("'", "'\\''")}'`;
}

function buildDeviceCurlCommand() {
  if (!annotationIngestInfo) return '';
  let payload;
  try {
    payload = readDevicePayload();
  } catch {
    return 'payload 不是合法 JSON，无法生成 curl';
  }
  const deviceId = $('deviceIdInput')?.value.trim();
  const deviceType = $('deviceTypeInput')?.value.trim();
  const lines = [
    `curl -sS -X POST ${shellSingleQuote(annotationIngestInfo.endpoint)} \\`,
    `  -H ${shellSingleQuote('content-type: application/json; charset=utf-8')} \\`,
  ];
  if (deviceId) lines.push(`  -H ${shellSingleQuote(`x-hmp-device-id: ${deviceId}`)} \\`);
  if (deviceType) lines.push(`  -H ${shellSingleQuote(`x-hmp-device-type: ${deviceType}`)} \\`);
  lines.push(`  --data-binary ${shellSingleQuote(JSON.stringify(payload))}`);
  const annotationId = String(payload.id ?? payload.annotation_id ?? payload.mark_id ?? payload.event_id ?? '');
  const statusUrl = annotationId && annotationIngestInfo.annotation_status_url
    ? annotationIngestInfo.annotation_status_url.replace('{annotation_id}', encodeURIComponent(annotationId))
    : annotationIngestInfo.annotation_status_url;
  if (statusUrl) {
    lines.push(
      '',
      '# 查询这条标注是否 pending / 已回绑真实会议轴 / 缺少采集时间',
      `curl -sS ${shellSingleQuote(statusUrl)}`,
    );
  }
  if (annotationIngestInfo.stream_url) {
    lines.push(
      '',
      '# 观察实时 SSE；页面时间轴也通过这个 state 广播刷新',
      `curl -N ${shellSingleQuote(annotationIngestInfo.stream_url)}`,
    );
  }
  return lines.join('\n');
}

function renderDeviceCurl() {
  const box = $('deviceCurlInput');
  if (!box) return;
  box.value = buildDeviceCurlCommand();
}

function sampleSequenceText() {
  return JSON.stringify([
    {
      id: 'mark-live-1',
      time_ms: 205000,
      kind: 'handwriting_question',
      label: '手写：这里为什么?',
      payload: { text_candidates: ['为什么?', '为什幺?', 'why?'] },
    },
    {
      id: 'mark-live-2',
      time_ms: 360000,
      kind: 'attention',
      label: '圈出：权限风险',
      payload: { action: 'enclosure' },
    },
  ], null, 2);
}

function sampleTranscriptText() {
  return JSON.stringify({
    transcript: [
      {
        start_ms: 0,
        end_ms: 8000,
        speaker_name: '飞书用户',
        text: '这是一段手动导入的飞书转写样例。',
      },
    ],
  }, null, 2);
}

async function load() {
  state = await api('/api/state');
  renderAll();
}

async function loadConfig() {
  larkConfig = await api('/api/lark/config');
  if (canonicalizeLocalhostFromConfig(larkConfig)) return;
  renderWebhookStatus(larkConfig);
  $('larkConfig').textContent = larkConfig.configured
    ? `飞书应用已配置：${larkConfig.base_url} · redirect_uri=${larkConfig.redirect_uri}`
    : `飞书应用未配置：请先在 .env 填 LARK_APP_ID / LARK_APP_SECRET。`;
  const auth = larkConfig.oauth;
  const user = auth?.user?.data ?? auth?.user;
  const name = user?.name || user?.en_name || user?.email || user?.open_id || user?.union_id || '';
  const meetingSearch = auth?.meeting_search;
  const scanAuthText = meetingSearch?.usable
    ? '会议扫描兜底 OAuth 已包含 vc:meeting.search:read。'
    : meetingSearch?.reason === 'missing_scope'
      ? '当前账号已登录，但缺少 vc:meeting.search:read；点“授权会议扫描兜底”重新登录补 scope。'
      : meetingSearch?.reason === 'oauth_token_expired'
        ? '当前 OAuth 已过期；点“授权会议扫描兜底”重新登录补会议扫描 scope。'
        : '会议扫描兜底需要单独授权 vc:meeting.search:read。';
  const scopeHint = $('larkAuthScopeHint');
  if (scopeHint) {
    scopeHint.textContent = `实时建轴主路径不是这里登录拿权限，而是在飞书开放平台给应用开通并发布 vc:meeting.all_meeting:readonly；这里的用户登录只用于会后妙记，同级的会议扫描授权只用于事件未投递时的兜底扫描。${scanAuthText}`;
  }
  const scanLoginBtn = $('larkMeetingScanLoginBtn');
  if (scanLoginBtn) {
    scanLoginBtn.textContent = meetingSearch?.usable ? '会议扫描已授权' : '授权会议扫描兜底';
    scanLoginBtn.title = '打开飞书 OAuth，请求 vc:meeting.search:read；只用于当前账号会议扫描兜底，不是实时事件建轴主路径';
  }
  $('larkAuthStatus').textContent = auth?.authenticated
    ? `飞书账号已授权${name ? `：${name}` : ''}${auth.token_expires_at ? ` · token 到期 ${auth.token_expires_at}` : ''}`
    : auth?.token_present && auth?.expired
      ? `飞书账号授权已过期${auth.token_expires_at ? `：${auth.token_expires_at}` : ''}，请重新登录。`
      : '飞书账号未授权。真实同步前请先登录飞书账号。';
}

async function loadDeliveryDiagnostics() {
  deliveryDiagnostics = await api('/api/lark/delivery-diagnostics');
  passiveMeetingScan = deliveryDiagnostics.active_search?.passive_scan ?? passiveMeetingScan;
  renderDeliveryDiagnostics();
}

async function loadAcceptanceReport() {
  acceptanceReport = await api('/api/lark/acceptance-report');
  autoAcceptance = acceptanceReport.current_evidence?.auto_acceptance ?? autoAcceptance;
  deviceSimulator = acceptanceReport.current_evidence?.device_simulator ?? deviceSimulator;
  passiveMeetingScan = acceptanceReport.current_evidence?.passive_meeting_scan ?? passiveMeetingScan;
  realDemoSession = acceptanceReport.current_evidence?.real_demo_session ?? realDemoSession;
  renderDeliveryDiagnostics();
  renderAcceptanceGuide();
}

async function loadRealDemoStatus() {
  realDemoProgress = await api('/api/lark/real-demo/progress');
  realDemoAcceptance = await api('/api/lark/real-demo/acceptance');
  realDemoStatus = realDemoProgress.result ?? await api('/api/lark/real-demo/status');
  deviceSimulator = realDemoStatus.device_simulator ?? deviceSimulator;
  passiveMeetingScan = realDemoStatus.passive_meeting_scan ?? passiveMeetingScan;
  realDemoSession = realDemoStatus.real_demo_session ?? realDemoSession;
  renderAcceptanceGuide();
}

async function loadReadiness() {
  readiness = await api('/api/readiness');
  realMeetingProbe = readiness.real_meeting_probe ?? realMeetingProbe;
  autoAcceptance = readiness.auto_acceptance ?? autoAcceptance;
  deviceSimulator = readiness.device_simulator ?? deviceSimulator;
  passiveMeetingScan = readiness.passive_meeting_scan ?? passiveMeetingScan;
  realDemoSession = readiness.real_demo_session ?? realDemoSession;
  renderReadiness();
  renderRealMeetingProbe();
  await loadDeliveryDiagnostics();
  await loadAcceptanceReport();
  await loadRealDemoStatus();
  if (state) renderAll();
}

async function loadAnnotationIngestInfo() {
  annotationIngestInfo = await api('/api/annotation-ingest-info');
  deviceSimulator = annotationIngestInfo.device_simulator ?? deviceSimulator;
  if (!$('deviceIdInput').value.trim()) $('deviceIdInput').value = 'hanwang-alpha-001';
  if (!$('deviceTypeInput').value.trim()) $('deviceTypeInput').value = 'hanwang_epaper';
  $('devicePayloadInput').value = JSON.stringify(buildDeviceDebugPayload(annotationIngestInfo), null, 2);
  renderDeviceIngestInfo();
}

async function loadRealMeetingProbe() {
  realMeetingProbe = await api('/api/lark/real-meeting-probe');
  renderRealMeetingProbe();
  if (state) renderAll();
}

async function maybeAutoBindProbeMeeting() {
  if (probeAutoBindInFlight) return;
  if (!realMeetingProbe?.active || realMeetingProbe?.status !== 'waiting') return;
  if (!realMeetingProbe.auto_search?.enabled) return;
  if (deliveryDiagnostics?.evidence?.real_meeting_axis_active) return;
  probeAutoBindInFlight = true;
  try {
    const result = await api('/api/lark/real-meeting-probe/auto-bind', {
      method: 'POST',
      body: JSON.stringify({}),
    });
    realMeetingProbe = {
      ...realMeetingProbe,
      auto_search: {
        ...(realMeetingProbe.auto_search ?? {}),
        last_attempt_at: result.at ?? realMeetingProbe.auto_search?.last_attempt_at,
        last_result: {
          status: result.status,
          reason: result.reason ?? null,
          at: result.at ?? null,
          selected_meeting_id: result.selected_meeting_id ?? result.state?.meeting?.meeting_id ?? null,
          error: result.error ?? null,
        },
      },
    };
    renderRealMeetingProbe();
    if (result.status === 'bound' && result.state) {
      state = result.state;
      renderAll();
      $('streamStatus').textContent = `自动扫描已绑定真实会议：${result.state.meeting.title || result.state.meeting.meeting_id}`;
      await loadReadiness();
    }
  } catch (error) {
    console.warn(error);
  } finally {
    probeAutoBindInFlight = false;
  }
}

function authStartRedirectUrl(scopes = []) {
  const query = new URLSearchParams();
  if (scopes.length) query.set('scope', scopes.join(' '));
  else query.set('purpose', 'minutes');
  query.set('redirect', '1');
  return `/api/auth/lark/start?${query.toString()}`;
}

function normalizeAuthPopupUrl(url) {
  const parsed = new URL(url, window.location.href);
  if (parsed.pathname === '/api/auth/lark/start') parsed.searchParams.set('redirect', '1');
  return parsed.toString();
}

function openAuthPopup(url, name) {
  const target = normalizeAuthPopupUrl(url);
  const popup = window.open(target, name, 'popup,width=960,height=760');
  if (!popup) window.location.href = target;
  return popup;
}

async function startLarkAuth(reason = 'lark_oauth', scopes = []) {
  const redirectUrl = authStartRedirectUrl(scopes);
  openAuthPopup(redirectUrl, reason);
  const scopeText = scopes.length ? scopes.join(' ') : '默认 scope';
  $('streamStatus').textContent = `已打开飞书授权窗口：scope=${scopeText}`;
  return { redirect_url: redirectUrl, scopes };
}

async function wire() {
  $('sequenceInput').value = sampleSequenceText();
  $('transcriptInput').value = sampleTranscriptText();
  $('liveTitleInput').value = '实时标注会议';
  $('liveOffsetInput').value = '';
  $('liveCapturedAtInput').value = '';
  $('liveLabelInput').value = '手写：这里为什么?';
  $('refreshBtn').addEventListener('click', load);
  $('resetBtn').addEventListener('click', async () => {
    state = await api('/api/demo/reset', { method: 'POST', body: '{}' });
    realDemoSession = null;
    realDemoMonitor = null;
    realDemoMonitorInFlight = false;
    await load();
    await loadReadiness();
  });
  $('importSequenceBtn').addEventListener('click', async () => {
    const sequence = JSON.parse($('sequenceInput').value);
    state = await api('/api/import/sequence', { method: 'POST', body: JSON.stringify({ sequence }) });
    renderAll();
  });
  $('importTranscriptBtn').addEventListener('click', async () => {
    const body = JSON.parse($('transcriptInput').value);
    state = await api('/api/import/lark-transcript', { method: 'POST', body: JSON.stringify(body) });
    renderAll();
  });
  $('larkLoginBtn').addEventListener('click', async () => {
    await startLarkAuth('lark_oauth');
  });
  $('larkMeetingScanLoginBtn').addEventListener('click', async () => {
    try {
      await startLarkAuth('lark_meeting_scope_oauth', ['vc:meeting.search:read']);
      $('streamStatus').textContent = '已打开飞书授权窗口：请求 vc:meeting.search:read，用于当前账号会议扫描兜底。';
    } catch (error) {
      $('streamStatus').textContent = `打开会议扫描授权失败：${error.message}`;
    }
  });
  $('larkRefreshAuthBtn').addEventListener('click', async () => {
    try {
      await api('/api/auth/lark/refresh', { method: 'POST', body: '{}' });
      await loadConfig();
      await loadReadiness();
      $('streamStatus').textContent = '飞书授权已刷新';
    } catch (error) {
      $('streamStatus').textContent = `刷新授权失败：${error.message}`;
      await loadConfig();
    }
  });
  $('larkLogoutBtn').addEventListener('click', async () => {
    await api('/api/auth/lark/logout', { method: 'POST', body: '{}' });
    await loadConfig();
  });
  $('startRealMeetingProbeBtn').addEventListener('click', async () => {
    realMeetingProbe = await api('/api/lark/real-meeting-probe/start', {
      method: 'POST',
      body: JSON.stringify({ timeout_ms: 120_000, note: 'direct meeting validation', reset_temporary_axis: true, auto_search: true }),
    });
    renderRealMeetingProbe();
    renderAcceptanceGuide();
    await loadDeliveryDiagnostics();
    await loadReadiness();
  });
  $('prepareRealDemoBtn').addEventListener('click', async () => {
    const likelyNeedsAuth = Boolean(realDemoStatus?.auth_start?.redirect_url)
      || meetingSearchFallbackStatus().needsReauth;
    const authPopup = likelyNeedsAuth
      ? window.open('about:blank', 'lark_real_demo_oauth', 'popup,width=960,height=760')
      : null;
    try {
      const result = await api('/api/lark/real-demo/prepare', {
        method: 'POST',
        body: JSON.stringify({ auto_annotation: true, passive_scan: true, tenant_fallback_enabled: false }),
      });
      autoAcceptance = result.auto_acceptance ?? autoAcceptance;
      deviceSimulator = result.device_simulator ?? deviceSimulator;
      if (result.device_stream_simulator) {
        deviceSimulator = {
          ...(deviceSimulator ?? {}),
          stream: result.device_stream_simulator,
        };
      }
      passiveMeetingScan = result.passive_meeting_scan ?? passiveMeetingScan;
      realDemoSession = result.real_demo_session ?? realDemoSession;
      if (state) renderAll();
      if (result.trigger?.status === 'bound') {
        $('streamStatus').textContent = `真实等待状态已开启，并已通过扫描绑定会议：${result.trigger.meeting_title || result.trigger.selected_meeting_id}`;
        await load();
      } else if (result.auth_required && (result.auth_start?.redirect_url || result.auth_start?.auth_url)) {
        const authUrl = result.auth_start.redirect_url || result.auth_start.auth_url;
        openAuthPopup(authUrl, 'lark_real_demo_oauth', authPopup);
        const scopes = result.auth_start.scopes?.join(' ') || 'vc:meeting.search:read';
        $('streamStatus').textContent = `真实等待状态已开启但尚未建轴，已打开飞书授权窗口：scope=${scopes}`;
      } else {
        if (authPopup) authPopup.close();
        $('streamStatus').textContent = result.next_step || '真实等待状态已开启但尚未建轴：现在可以直接开启飞书会议。扫描兜底需要单独点击“扫描我的真实会议”或打开“被动扫描建轴”。';
      }
      await loadReadiness();
    } catch (error) {
      if (authPopup) authPopup.close();
      $('streamStatus').textContent = `进入真实等待状态失败：${error.message}`;
    }
  });
  $('monitorRealDemoBtn').addEventListener('click', async () => {
    realDemoMonitorInFlight = true;
    realDemoMonitor = null;
    renderAcceptanceGuide();
    $('streamStatus').textContent = '正在等待真实会议完成证据：请直接开启飞书会议，设备流会在真实轴建立后写入标注。';
    try {
      const result = await api('/api/lark/real-demo/monitor', {
        method: 'POST',
        body: JSON.stringify({ timeout_ms: 180_000, interval_ms: 500 }),
      });
      realDemoMonitor = result;
      realDemoStatus = result.result ?? realDemoStatus;
      if (result.result?.device_simulator) deviceSimulator = result.result.device_simulator;
      if (result.result?.passive_meeting_scan) passiveMeetingScan = result.result.passive_meeting_scan;
      if (result.result?.real_demo_session) realDemoSession = result.result.real_demo_session;
      $('streamStatus').textContent = result.observed
        ? '真实会议完成证据已捕获：真实轴、开放标注和 SSE 广播均有记录。'
        : '等待完成证据超时：请检查飞书会议开始事件是否投递。';
      await load();
      await loadReadiness();
    } catch (error) {
      $('streamStatus').textContent = `等待完成证据失败：${error.message}`;
    } finally {
      realDemoMonitorInFlight = false;
      renderAcceptanceGuide();
    }
  });
  $('startAcceptanceBtn').addEventListener('click', async () => {
    realMeetingProbe = await api('/api/lark/real-meeting-probe/start', {
      method: 'POST',
      body: JSON.stringify({ timeout_ms: 180_000, note: 'direct meeting end-to-end validation', reset_temporary_axis: true, auto_search: true }),
    });
    renderRealMeetingProbe();
    renderAcceptanceGuide();
    $('streamStatus').textContent = '验收探针已启动但不会建轴：现在请在飞书客户端直接开启会议；带 captured_at_ms 的标注先到会进入 pending，真实事件到达后自动重绑定。';
    await loadDeliveryDiagnostics();
    await loadReadiness();
  });
  $('writeAcceptanceMarkBtn').addEventListener('click', async () => {
    try {
      const result = await api('/api/annotations', {
        method: 'POST',
        body: JSON.stringify({
          id: `acceptance-${Date.now()}`,
          source: 'browser_acceptance',
          captured_at_ms: Date.now(),
          kind: 'acceptance_mark',
          label: '验收标注：实时写入',
          realtime: true,
          mode: 'realtime',
          text_candidates: ['验收标注：实时写入', 'acceptance mark'],
          intent: 'acceptance_check',
          strokes: [],
          payload: {
            origin: 'acceptance_guide',
            realtime: true,
            mode: 'realtime',
          },
        }),
      });
      state = result.state;
      renderAll();
      $('streamStatus').textContent = `验收标注已写入：${result.item.label} @ ${fmtTime(result.item.time_ms)}`;
      await loadReadiness();
    } catch (error) {
      $('streamStatus').textContent = `写入验收标注失败：${error.message}`;
    }
  });
  $('toggleAutoAcceptanceBtn').addEventListener('click', async () => {
    try {
      const nextEnabled = !autoAcceptance?.enabled;
      const result = await api('/api/acceptance/auto-annotation', {
        method: 'POST',
        body: JSON.stringify({ enabled: nextEnabled }),
      });
      autoAcceptance = result.auto_acceptance;
      renderAcceptanceGuide();
      $('streamStatus').textContent = nextEnabled
        ? '自动验收标注已开启：真实会议轴建立后会自动写入一条开放标注。'
        : '自动验收标注已关闭。';
      await loadReadiness();
    } catch (error) {
      $('streamStatus').textContent = `切换自动验收标注失败：${error.message}`;
    }
  });
  $('toggleDeviceSimulatorBtn').addEventListener('click', async () => {
    try {
      const nextEnabled = !deviceSimulator?.enabled;
      const result = await api('/api/device-simulator', {
        method: 'POST',
        body: JSON.stringify({ enabled: nextEnabled, trigger_now: nextEnabled }),
      });
      deviceSimulator = result.device_simulator;
      if (result.trigger?.annotation) {
        await load();
      }
      renderAcceptanceGuide();
      renderDeviceIngestInfo();
      $('streamStatus').textContent = nextEnabled
        ? `虚拟设备标注已开启${result.trigger?.annotation ? '，并已写入当前真实会议轴' : ''}。`
        : '虚拟设备标注已关闭；外部设备接口仍可用。';
      await loadReadiness();
    } catch (error) {
      $('streamStatus').textContent = `切换虚拟设备标注失败：${error.message}`;
    }
  });
  $('togglePassiveScanBtn').addEventListener('click', async () => {
    try {
      const nextEnabled = !passiveMeetingScan?.enabled;
      const result = await api('/api/lark/passive-meeting-scan', {
        method: 'POST',
        body: JSON.stringify({ enabled: nextEnabled, trigger_now: nextEnabled }),
      });
      passiveMeetingScan = result.passive_meeting_scan;
      const triggerText = result.trigger?.status
        ? `；本次触发：${result.trigger.status}${result.trigger.reason ? `/${result.trigger.reason}` : ''}`
        : '';
      renderAcceptanceGuide();
      $('streamStatus').textContent = nextEnabled
        ? `被动扫描建轴已开启${triggerText}`
        : '被动扫描建轴已关闭；仍会等待飞书真实事件投递。';
      if (result.trigger?.state) {
        state = result.trigger.state;
        renderAll();
      }
      await loadReadiness();
    } catch (error) {
      $('streamStatus').textContent = `切换被动扫描失败：${error.message}`;
    }
  });
  $('reauthMeetingScopeBtn').addEventListener('click', async () => {
    try {
      const redirectUrl = $('reauthMeetingScopeBtn').dataset.redirectUrl;
      if (redirectUrl) {
        openAuthPopup(redirectUrl, 'lark_meeting_scope_oauth');
      } else {
        await startLarkAuth('lark_meeting_scope_oauth', ['vc:meeting.search:read']);
      }
      $('streamStatus').textContent = '已打开飞书授权窗口；授权完成后被动扫描兜底会自动恢复。';
    } catch (error) {
      $('streamStatus').textContent = `打开飞书授权失败：${error.message}`;
    }
  });
  $('refreshProbeBtn').addEventListener('click', async () => {
    await loadRealMeetingProbe();
    renderAcceptanceGuide();
    await loadReadiness();
  });
  $('resetProbeBtn').addEventListener('click', async () => {
    realMeetingProbe = await api('/api/lark/real-meeting-probe/reset', {
      method: 'POST',
      body: '{}',
    });
    renderRealMeetingProbe();
    renderAll();
    $('streamStatus').textContent = '验收状态已清除；需要真实验收时请重新点击“启动验收探针”。';
    await loadReadiness();
  });
  $('restoreLatestRealAxisBtn').addEventListener('click', async () => {
    try {
      const result = await api('/api/lark/restore-latest-real-meeting-axis', {
        method: 'POST',
        body: JSON.stringify({ apply_end: true }),
      });
      state = result.state;
      renderAll();
      $('streamStatus').textContent = result.end_event
        ? `已从真实飞书 start/end 事件恢复会议轴：${state.meeting.title || state.meeting.meeting_id}`
        : `已从真实飞书 start 事件恢复会议轴：${state.meeting.title || state.meeting.meeting_id}`;
      await loadRealMeetingProbe();
      await loadReadiness();
    } catch (error) {
      $('streamStatus').textContent = `恢复最近真实会议轴失败：${error.message}`;
    }
  });
  $('refreshDeviceIngestBtn').addEventListener('click', async () => {
    try {
      await loadAnnotationIngestInfo();
      $('streamStatus').textContent = '设备接入信息已刷新。';
    } catch (error) {
      $('deviceIngestStatus').textContent = `读取接入信息失败：${error.message}`;
    }
  });
  $('copyDeviceCurlBtn').addEventListener('click', async () => {
    try {
      readDevicePayload({ refreshDebug: true });
      renderDeviceCurl();
      const text = $('deviceCurlInput').value;
      await navigator.clipboard.writeText(text);
      $('streamStatus').textContent = '外部设备 curl 已复制。';
    } catch {
      $('streamStatus').textContent = '复制失败：请手动选中 curl 文本复制。';
    }
  });
  $('deviceIdInput').addEventListener('input', renderDeviceCurl);
  $('deviceTypeInput').addEventListener('input', renderDeviceCurl);
  $('devicePayloadInput').addEventListener('input', renderDeviceCurl);
  $('sendDeviceMarkBtn').addEventListener('click', async () => {
    try {
      const payload = readDevicePayload({ refreshDebug: true });
      const deviceId = $('deviceIdInput').value.trim();
      const deviceType = $('deviceTypeInput').value.trim();
      const headers = {};
      if (deviceId) headers['x-hmp-device-id'] = deviceId;
      if (deviceType) headers['x-hmp-device-type'] = deviceType;
      const result = await api('/api/annotations', {
        method: 'POST',
        headers,
        body: JSON.stringify(payload),
      });
      state = result.state;
      renderAll();
      const source = timeSourceInfo(result.item);
      $('streamStatus').textContent = `设备标注已写入：${result.item.label} @ ${fmtTime(result.item.time_ms)} · ${source.label}`;
      await loadAnnotationIngestInfo();
      await loadReadiness();
    } catch (error) {
      $('deviceIngestStatus').textContent = `发送设备标注失败：${error.message}`;
    }
  });
  $('startDeviceStreamBtn').addEventListener('click', async () => {
    try {
      const result = await api('/api/device-simulator/stream', {
        method: 'POST',
        body: JSON.stringify({
          action: 'start',
          interval_ms: 1500,
          max_count: 5,
          label_prefix: '流式设备标注',
        }),
      });
      deviceSimulator = {
        ...(deviceSimulator ?? {}),
        stream: result,
      };
      renderDeviceIngestInfo();
      $('streamStatus').textContent = result.status === 'waiting_for_real_axis' || result.status === 'starting'
        ? '流式设备模拟已启动：会等真实会议轴出现后连续写入开放标注。'
        : `流式设备模拟已启动：${result.status}`;
      await loadReadiness();
    } catch (error) {
      $('deviceIngestStatus').textContent = `启动流式模拟失败：${error.message}`;
    }
  });
  $('stopDeviceStreamBtn').addEventListener('click', async () => {
    try {
      const result = await api('/api/device-simulator/stream', {
        method: 'POST',
        body: JSON.stringify({ action: 'stop' }),
      });
      deviceSimulator = {
        ...(deviceSimulator ?? {}),
        stream: result,
      };
      renderDeviceIngestInfo();
      $('streamStatus').textContent = '流式设备模拟已停止。';
    } catch (error) {
      $('deviceIngestStatus').textContent = `停止流式模拟失败：${error.message}`;
    }
  });
  $('startOpenSessionBtn').addEventListener('click', async () => {
    const title = $('liveTitleInput').value.trim() || '飞书实时会议';
    const meeting_url = $('liveMeetingUrlInput').value.trim() || null;
    const startInput = $('meetingStartInput').value.trim();
    const body = {
      platform: 'lark',
      title,
      meeting_url,
      detector_source: 'browser_operator',
    };
    if (startInput) body.start_time = startInput;
    try {
      const result = await api('/api/meeting-session/start', {
        method: 'POST',
        body: JSON.stringify(body),
      });
      state = result.state;
      realDemoSession = state.real_demo_session ?? realDemoSession;
      $('liveOffsetInput').value = '';
      renderAll();
      $('streamStatus').textContent = `开放会议会话已建轴：${result.meeting.title} · ${result.meeting.meeting_id}`;
      await loadDeliveryDiagnostics();
      await loadReadiness();
    } catch (error) {
      $('streamStatus').textContent = `开放会话建轴失败：${error.message}`;
      await load().catch(() => {});
    }
  });
  $('searchMinutesBtn').addEventListener('click', async () => {
    const query = $('minuteSearchInput').value.trim();
    $('minuteResults').innerHTML = '<p class="hint">正在搜索真实飞书妙记...</p>';
    try {
      const result = await api('/api/lark/search-minutes', {
        method: 'POST',
        body: JSON.stringify({ query, page_size: 20 }),
      });
      renderMinuteResults(result.items, result.raw);
    } catch (error) {
      $('minuteResults').innerHTML = `<p class="hint">${escapeHtml(error.message)}</p>`;
    }
  });
  $('createLarkReserveBtn').addEventListener('click', async () => {
    const title = $('liveTitleInput').value.trim() || '实时标注会议';
    try {
      const result = await api('/api/lark/create-reserve', {
        method: 'POST',
        body: JSON.stringify({ title, meeting_connect: true }),
      });
      state = result.state;
      const url = result.reserve?.url || result.reserve?.app_link || state.meeting?.meeting_url || '';
      if (url) {
        $('liveMeetingUrlInput').value = url;
        window.open(url, 'lark_meeting');
      }
      renderAll();
      $('streamStatus').textContent = '已创建飞书会议链接，等待真实会议开始事件或扫描/同步真实会议后建轴。';
      await loadReadiness();
    } catch (error) {
      $('streamStatus').textContent = error.details?.required_scope
        ? `创建飞书会议失败：需要在飞书开放平台开通 ${error.details.required_scope} 权限。`
        : `创建飞书会议失败：${error.message}`;
      await loadReadiness();
    }
  });
  $('syncActiveMeetingBtn').addEventListener('click', async () => {
    try {
      const result = await api('/api/lark/reserve-active-meeting', {
        method: 'POST',
        body: JSON.stringify({}),
      });
      state = result.state;
      renderAll();
      await loadReadiness();
    } catch (error) {
      $('streamStatus').textContent = error.details?.required_scope
        ? `同步活跃会议失败：需要在飞书开放平台开通 ${error.details.required_scope} 权限。`
        : `同步活跃会议失败：${error.message}`;
      await loadReadiness();
    }
  });
  $('bindMyMeetingBtn').addEventListener('click', async () => {
    try {
      $('streamStatus').textContent = '正在扫描当前飞书账号的近期真实会议...';
      const result = await api('/api/lark/bind-my-latest-meeting', {
        method: 'POST',
        body: JSON.stringify({
          query: $('liveTitleInput').value.trim(),
          lookback_seconds: 2 * 60 * 60,
          lookahead_seconds: 30 * 60,
        }),
      });
      state = result.state;
      renderAll();
      $('streamStatus').textContent = `已扫描并绑定真实会议：${state.meeting.title} · ${result.search.auth_mode}`;
      await loadReadiness();
    } catch (error) {
      $('streamStatus').textContent = error.details?.required_scope
        ? `扫描我的真实会议失败：当前授权缺少 ${error.details.required_scope}，请在飞书开放平台开通权限后重新登录。`
        : `扫描我的真实会议失败：${error.message}`;
      await loadReadiness();
    }
  });
  $('bindTenantMeetingBtn').addEventListener('click', async () => {
    try {
      $('streamStatus').textContent = '正在尝试实验性应用身份扫描；如果飞书拒绝 token，请改用“授权会议扫描兜底”和“扫描我的真实会议”。';
      const result = await api('/api/lark/bind-tenant-latest-meeting', {
        method: 'POST',
        body: JSON.stringify({
          query: $('liveTitleInput').value.trim(),
          lookback_seconds: 15 * 60,
          lookahead_seconds: 30 * 60,
        }),
      });
      state = result.state;
      renderAll();
      $('streamStatus').textContent = `已用应用身份绑定真实会议：${state.meeting.title} · ${result.search.auth_mode}`;
      await loadReadiness();
    } catch (error) {
      $('streamStatus').textContent = error.details?.search_status === 'invalid_token'
        ? '应用身份扫描失败：飞书拒绝应用/租户 access token。请点击“授权会议扫描兜底”，再用“扫描我的真实会议”。'
        : error.details?.required_scope
          ? `应用身份扫描失败：需要在飞书开放平台开通 ${error.details.required_scope} 权限。`
          : `应用身份扫描失败：${error.message}`;
      await loadReadiness();
    }
  });
  $('bindMeetingNoBtn').addEventListener('click', async () => {
    const meeting_url = $('liveMeetingUrlInput').value.trim();
    if (!meeting_url) {
      $('streamStatus').textContent = '按会议号绑定失败：请先粘贴飞书会议链接或会议号。';
      return;
    }
    try {
      const result = await api('/api/lark/bind-meeting-by-no', {
        method: 'POST',
        body: JSON.stringify({ meeting_url }),
      });
      state = result.state;
      renderAll();
      $('streamStatus').textContent = `已按会议号绑定：${result.meeting_no} · ${result.auth_mode}`;
      await loadReadiness();
    } catch (error) {
      $('streamStatus').textContent = error.details?.required_scope
        ? `按会议号绑定失败：需要开通 ${error.details.required_scope} 权限。`
        : `按会议号绑定失败：${error.message}`;
    }
  });
  $('startLiveBtn').addEventListener('click', async () => {
    const title = $('liveTitleInput').value.trim() || '实时标注会议';
    const meeting_url = $('liveMeetingUrlInput').value.trim() || null;
    try {
      state = await api('/api/live/start-meeting', {
        method: 'POST',
        body: JSON.stringify({ title, meeting_url }),
      });
      $('liveOffsetInput').value = '';
      renderAll();
    } catch (error) {
      $('streamStatus').textContent = `本地模拟开始失败：${error.message}`;
      await load().catch(() => {});
    }
  });
  $('endLiveBtn').addEventListener('click', async () => {
    try {
      state = await api('/api/live/end-meeting', { method: 'POST', body: '{}' });
      renderAll();
    } catch (error) {
      $('streamStatus').textContent = `本地模拟结束失败：${error.message}`;
      await load().catch(() => {});
    }
  });
  $('syncMinuteBtn').addEventListener('click', async () => {
    const minute_token = $('minuteTokenInput').value.trim();
    const start_time = $('meetingStartInput').value.trim();
    state = await api('/api/lark/sync-minute', {
      method: 'POST',
      body: JSON.stringify({ minute_token, start_time }),
    });
    renderAll();
  });
  $('liveMarkBtn').addEventListener('click', async () => {
    const offset = parseOffsetMs($('liveOffsetInput').value);
    const capturedAt = $('liveCapturedAtInput').value.trim();
    const label = $('liveLabelInput').value.trim() || '实时标注';
    const body = {
      source: 'browser_demo',
      kind: 'live_mark',
      label,
      realtime: true,
      mode: 'realtime',
      payload: {
        text_candidates: [label],
        realtime: true,
        mode: 'realtime',
      },
    };
    if (capturedAt) {
      if (/^\d+(\.\d+)?$/.test(capturedAt)) body.captured_at_ms = Number(capturedAt);
      else body.captured_at = capturedAt;
    } else {
      body.captured_at_ms = Date.now();
    }
    if (offset != null) body.time_ms = offset;
    const result = await api('/api/annotations', {
      method: 'POST',
      body: JSON.stringify(body),
    });
    state = result.state;
    if (offset == null) $('liveOffsetInput').value = '';
    $('liveCapturedAtInput').value = '';
    renderAll();
    const source = timeSourceInfo(result.item);
    $('streamStatus').textContent = `实时标注已写入：${result.item.label} @ ${fmtTime(result.item.time_ms)} · ${source.label}`;
  });
  $('liveCapturedAtInput').addEventListener('input', renderDebugInputControls);

  window.addEventListener('message', async (event) => {
    if (event.data?.type === 'lark-auth-complete') {
      await loadConfig();
      const passive = event.data.passive_scan;
      if (passive?.status === 'bound') {
        $('streamStatus').textContent = `飞书授权完成，已自动绑定会议：${passive.meeting_title || passive.selected_meeting_id}`;
        await load();
      } else if (passive?.status) {
        $('streamStatus').textContent = `飞书授权完成，会议扫描结果：${passive.status}${passive.reason ? `/${passive.reason}` : ''}`;
      }
      await loadReadiness();
    }
  });
  await load();
  await loadConfig();
  await loadReadiness();
  await loadRealMeetingProbe();
  await loadAnnotationIngestInfo();
  if (window.EventSource) {
    stream = new EventSource('/api/stream');
    $('streamStatus').textContent = '实时同步：连接中';
    stream.addEventListener('open', () => {
      $('streamStatus').textContent = '实时同步：已连接';
    });
    stream.addEventListener('state', (event) => {
      state = JSON.parse(event.data);
      renderAll();
    });
    stream.addEventListener('error', () => {
      $('streamStatus').textContent = '实时同步：断开，浏览器会自动重连';
    });

    realDemoProgressStream = new EventSource('/api/lark/real-demo/progress-stream');
    realDemoProgressStream.addEventListener('progress', (event) => {
      realDemoProgress = JSON.parse(event.data);
      realDemoStatus = realDemoProgress.result ?? realDemoStatus;
      deviceSimulator = realDemoStatus?.device_simulator ?? deviceSimulator;
      passiveMeetingScan = realDemoStatus?.passive_meeting_scan ?? passiveMeetingScan;
      realDemoSession = realDemoStatus?.real_demo_session ?? realDemoSession;
      api('/api/lark/real-demo/acceptance')
        .then((value) => {
          realDemoAcceptance = value;
          renderAcceptanceGuide();
        })
        .catch(() => {});
      renderAcceptanceGuide();
      if (state) renderAll();
    });
    realDemoProgressStream.addEventListener('error', () => {
      console.warn('real demo progress stream disconnected; browser will retry');
    });
  } else {
    $('streamStatus').textContent = '实时同步：当前浏览器不支持 EventSource';
  }
  setInterval(() => {
    loadConfig().catch((error) => console.warn(error));
    loadReadiness().catch((error) => console.warn(error));
    loadRealMeetingProbe().catch((error) => console.warn(error));
    maybeAutoBindProbeMeeting().catch((error) => console.warn(error));
  }, 5000);
}

wire().catch((error) => {
  console.error(error);
  document.body.innerHTML = `<pre>${escapeHtml(error.stack || error.message || error)}</pre>`;
});
