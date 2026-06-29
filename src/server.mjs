import { createServer } from 'node:http';
import { randomUUID } from 'node:crypto';
import { createReadStream, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { extname, join, normalize } from 'node:path';
import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import * as Lark from '@larksuiteoapi/node-sdk';
import { createLarkClient, extractMeetingNo, extractMinuteToken } from './larkClient.mjs';
import {
  annotationCapturedAbsoluteMs,
  annotationCapturedAbsoluteMsStrict,
  buildTimeline,
  extractLarkMeetingPatch,
  mergeTimeline,
  normalizeAnnotationEvent,
  normalizeLarkEventPayload,
  parseAbsoluteMs,
  normalizeTranscript,
} from './normalize.mjs';
import { TimelineStore } from './store.mjs';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const publicDir = join(root, 'public');
const dataDir = process.env.TIMELINE_DATA_DIR || join(root, 'data');
const authPath = process.env.LARK_AUTH_PATH || join(dataDir, 'lark-auth.json');
const realMeetingProbePath = process.env.REAL_MEETING_PROBE_PATH || join(dataDir, 'real-meeting-probe.json');
const larkEventLogPath = process.env.LARK_EVENT_LOG_PATH || join(dataDir, 'lark-events-log.json');
const autoAcceptancePath = process.env.AUTO_ACCEPTANCE_PATH || join(dataDir, 'auto-acceptance.json');
const deviceSimulatorPath = process.env.DEVICE_SIMULATOR_PATH || join(dataDir, 'device-simulator.json');
const passiveMeetingScanPath = process.env.PASSIVE_MEETING_SCAN_PATH || join(dataDir, 'passive-meeting-scan.json');
const realDemoSessionPath = process.env.REAL_DEMO_SESSION_PATH || join(dataDir, 'real-demo-session.json');
const store = new TimelineStore();
const streamClients = new Set();
const streamTelemetry = {
  opened_count: 0,
  closed_count: 0,
  broadcast_count: 0,
  last_opened_at: null,
  last_closed_at: null,
  last_broadcast_at: null,
  last_event: null,
  last_client_count: 0,
  last_state_summary: null,
};

function boolFromEnv(value, fallback = false) {
  if (value == null || value === '') return fallback;
  const text = String(value).trim().toLowerCase();
  if (['1', 'true', 'yes', 'on'].includes(text)) return true;
  if (['0', 'false', 'no', 'off'].includes(text)) return false;
  return fallback;
}

function boundedNumber(value, fallback, min, max) {
  const number = Number(value);
  if (!Number.isFinite(number)) return fallback;
  return Math.min(Math.max(number, min), max);
}

const larkEventLog = loadLarkEventLogState();
const defaultRealMeetingProbe = {
  active: false,
  started_at: null,
  timeout_ms: 120_000,
  note: null,
  auto_search: {
    enabled: false,
    interval_ms: 5000,
    last_attempt_at: null,
    last_result: null,
  },
};
let realMeetingProbe = loadRealMeetingProbeState();
const defaultAutoAcceptance = {
  enabled: false,
  label: '验收标注：自动写入',
  count: 0,
  last_annotation_id: null,
  last_meeting_id: null,
  last_trigger_at: null,
  updated_at: null,
};
let autoAcceptance = loadAutoAcceptanceState();
const defaultDeviceSimulator = {
  enabled: false,
  device_id: 'hanwang-demo-device',
  device_type: 'hanwang_epaper_simulator',
  label: '虚拟墨水屏标注：why?',
  count: 0,
  last_annotation_id: null,
  last_meeting_id: null,
  last_trigger_at: null,
  updated_at: null,
};
let deviceSimulator = loadDeviceSimulatorState();
const defaultPassiveMeetingScan = {
  enabled: false,
  tenant_fallback_enabled: true,
  interval_ms: 10_000,
  tenant_fallback_cooldown_ms: 5 * 60_000,
  lookback_seconds: 10 * 60,
  lookahead_seconds: 2 * 60,
  last_attempt_at: null,
  last_result: null,
  updated_at: null,
};
const oauthStateTtlMs = 15 * 60 * 1000;
const oauthStateHistoryLimit = 10;
let passiveMeetingScan = loadPassiveMeetingScanState();
const defaultRealDemoSession = {
  active: false,
  auto_open_session_on_annotation: true,
  prepared_at: null,
  last_real_axis_at: null,
  last_real_axis_source: null,
  last_note: null,
  updated_at: null,
};
let realDemoSession = loadRealDemoSessionState();
const minuteOAuthScopes = [
  'minutes:minutes.search:read',
  'minutes:minutes.basic:read',
  'minutes:minutes.transcript:export',
];

function loadDotEnv() {
  const envPath = join(root, '.env');
  if (!existsSync(envPath)) return;
  const text = readFileSync(envPath, 'utf8');
  for (const line of text.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const eq = trimmed.indexOf('=');
    if (eq <= 0) continue;
    const key = trimmed.slice(0, eq).trim();
    const value = trimmed.slice(eq + 1).trim().replace(/^["']|["']$/g, '');
    if (process.env[key] == null) process.env[key] = value;
  }
}

loadDotEnv();
const lark = createLarkClient(process.env);
const port = Number(process.env.PORT || 8787);
let authState = loadAuthState();
let larkWsClient = null;
let probeAutoSearchTimer = null;
let probeAutoSearchInFlight = false;
let passiveMeetingScanTimer = null;
let passiveMeetingScanInFlight = false;
let deviceStreamTimer = null;
let deviceStreamInFlight = false;
let deviceStreamState = {
  enabled: false,
  status: 'idle',
  interval_ms: 1500,
  max_count: 5,
  count: 0,
  started_at: null,
  stopped_at: null,
  last_annotation_id: null,
  last_meeting_id: null,
  last_trigger_at: null,
  last_error: null,
  label_prefix: '流式设备标注',
};
const capabilityStatus = {
  reserve: {
    status: 'unknown',
    checked_at: null,
    required_scope: 'vc:reserve',
    permission_url: `https://open.feishu.cn/app/${process.env.LARK_APP_ID}/auth?q=vc:reserve&op_from=openapi&token_type=tenant`,
    error: null,
  },
  meeting_lookup: {
    status: 'unknown',
    checked_at: null,
    required_scope: 'vc:meeting.search:read or vc:meeting.meetingid:read',
    permission_url: null,
    error: null,
  },
  meeting_search: {
    status: 'unknown',
    checked_at: null,
    required_scope: 'vc:meeting.search:read',
    permission_url: null,
    error: null,
  },
};
const larkWsStatus = {
  enabled: false,
  state: 'disabled',
  mode: 'long_connection',
  last_event_at: null,
  last_event_type: null,
  last_timeline_event_at: null,
  last_timeline_event_type: null,
  error: null,
};
const meetingEventRequirements = {
  direct_meeting: {
    label: '用户直接开启飞书会议',
    start_event: 'vc.meeting.all_meeting_started_v1',
    end_event: 'vc.meeting.all_meeting_ended_v1',
    required_scope: 'vc:meeting.all_meeting:readonly',
    permission_label: '获取所有视频会议信息',
    start_doc_url: 'https://open.feishu.cn/document/server-docs/vc-v1/meeting/events/all_meeting_started',
    end_doc_url: 'https://open.feishu.cn/document/server-docs/vc-v1/meeting/events/all_meeting_ended',
    note: '这是本 demo 的主路径：用户直接在飞书里开会，应用收到企业会议开始/结束事件后自动建轴。',
  },
  reserve_meeting: {
    label: 'OpenAPI 预约/创建会议',
    start_event: 'vc.meeting.meeting_started_v1',
    end_event: 'vc.meeting.meeting_ended_v1',
    required_scope: 'vc:reserve',
    permission_label: '预约会议',
    start_doc_url: 'https://open.feishu.cn/document/server-docs/vc-v1/meeting/events/meeting_started',
    end_doc_url: 'https://open.feishu.cn/document/server-docs/vc-v1/meeting/events/meeting_ended',
    note: '这是可选路径：demo 自己创建会议链接并立即建轴。',
  },
  meeting_context: {
    label: '会议上下文兜底事件',
    start_event: 'vc.meeting.join_meeting_v1',
    end_event: 'vc.meeting.leave_meeting_v1',
    required_scope: 'vc:meeting.search:read',
    permission_label: '获取会议信息',
    start_doc_url: 'https://open.feishu.cn/document/server-docs/vc-v1/meeting/events/join_meeting',
    end_doc_url: 'https://open.feishu.cn/document/server-docs/vc-v1/meeting/events/leave_meeting',
    note: '这是事件订阅兜底：如果企业会议开始事件未到，首次真实加入会议事件也可以用来创建真实会议时间轴。',
  },
  meeting_lookup: {
    label: '会议号/链接查询兜底',
    required_scope: 'vc:meeting.search:read',
    note: '事件订阅未打通时，可粘贴会议号/会议链接查询真实会议并建轴；这不是主路径。',
  },
};

function openPlatformChecklist(config = {}, missing = {}) {
  const ws = config.ws_event_receiver ?? {};
  const missingDirectEvents = missing.direct_events ?? [];
  const missingFallbackEvents = missing.fallback_events ?? [];
  const longConnectionReady = Boolean(ws.enabled && ws.state === 'connected');
  const publicWebhookReady = Boolean(config.event_callback_public_https);
  const recentWsEventCount = Number(config.recent_ws_event_count ?? 0);
  const recentPublicWebhookEventCount = publicWebhookReady
    ? (config.recent_events ?? []).filter((event) => event.transport === 'http_webhook').length
    : 0;
  const realDeliveryEventCount = recentWsEventCount + recentPublicWebhookEventCount;
  return [
    {
      id: 'event_delivery_mode',
      label: '事件接收方式',
      status: longConnectionReady || publicWebhookReady ? 'ok' : 'todo',
      required: '飞书开放平台事件订阅选择长连接；不用长连接时必须配置公网 HTTPS webhook。',
      evidence: longConnectionReady
        ? `long_connection ${ws.state}`
        : publicWebhookReady
          ? `public webhook ${config.event_callback_url}`
          : `long_connection ${ws.state ?? 'unknown'}, webhook_public_https=${publicWebhookReady}`,
      docs: ['https://open.feishu.cn/document/server-docs/event-subscription-guide/events-with-connection'],
    },
    {
      id: 'event_subscription_publish',
      label: '事件订阅保存与版本发布',
      status: realDeliveryEventCount > 0 ? 'ok' : 'manual_check',
      required: '开放平台事件订阅已保存为长连接或公网 HTTPS，订阅事件已添加，并且应用版本已发布到当前测试企业。',
      evidence: realDeliveryEventCount > 0
        ? '已收到真实事件投递'
        : '本地只能证明长连接已连接，不能证明开放平台订阅配置已保存并发布；如果开会后 ws_event_count=0，优先检查这一项。',
      docs: ['https://open.feishu.cn/document/server-docs/event-subscription-guide/events-with-connection'],
    },
    {
      id: 'event_delivery_evidence',
      label: '事件投递证据',
      status: realDeliveryEventCount > 0 ? 'ok' : longConnectionReady || publicWebhookReady ? 'todo' : 'blocked',
      required: '直接开启一次飞书会议后，事件日志应至少出现一条 ws_long_connection 或公网 http_webhook 事件。',
      evidence: realDeliveryEventCount > 0
        ? `recent_ws_event_count=${recentWsEventCount}, recent_public_webhook_event_count=${recentPublicWebhookEventCount}`
        : longConnectionReady
          ? '长连接已 connected，但 recent_ws_event_count=0，说明飞书云端尚未向当前应用投递事件。'
          : '事件接收通道尚未就绪，无法形成真实投递证据。',
      docs: ['https://open.feishu.cn/document/server-docs/event-subscription-guide/events-with-connection'],
    },
    {
      id: 'direct_meeting_events',
      label: '直开会议事件订阅',
      status: missingDirectEvents.length ? 'todo' : 'ok',
      required: `${meetingEventRequirements.direct_meeting.start_event} / ${meetingEventRequirements.direct_meeting.end_event}`,
      evidence: missingDirectEvents.length
        ? `本地接收器缺少 ${missingDirectEvents.join(', ')}`
        : '本地接收器已注册直开会议 start/end 事件',
      docs: [
        meetingEventRequirements.direct_meeting.start_doc_url,
        meetingEventRequirements.direct_meeting.end_doc_url,
      ],
    },
    {
      id: 'meeting_context_events',
      label: '会议上下文兜底事件订阅',
      status: missingFallbackEvents.length ? 'warn' : 'ok',
      required: `${meetingEventRequirements.meeting_context.start_event} / ${meetingEventRequirements.meeting_context.end_event}`,
      evidence: missingFallbackEvents.length
        ? `本地接收器缺少 ${missingFallbackEvents.join(', ')}`
        : '本地接收器已注册 join/leave 兜底事件',
      docs: [
        meetingEventRequirements.meeting_context.start_doc_url,
        meetingEventRequirements.meeting_context.end_doc_url,
      ],
    },
    {
      id: 'direct_meeting_permission',
      label: '直开会议权限',
      status: 'manual_check',
      required: `${meetingEventRequirements.direct_meeting.required_scope}（${meetingEventRequirements.direct_meeting.permission_label}），并发布到当前测试企业。`,
      evidence: '飞书不会通过本地 API 暴露该权限是否已发布；以真实事件是否投递为最终证据。',
      action_url: permissionUrlForScopes(meetingEventRequirements.direct_meeting.required_scope),
    },
    {
      id: 'app_availability',
      label: '应用可见范围',
      status: 'manual_check',
      required: '应用可见范围包含当前开会账号，且版本已发布/测试企业可用。',
      evidence: '如果长连接 connected 且本地已注册事件但 recent_ws_event_count=0，优先检查这一项和事件订阅是否发布。',
    },
  ];
}

function realMeetingEventAudit(config = {}, options = {}) {
  const ws = config.ws_event_receiver ?? {};
  const eventSummary = config.event_log_summary ?? larkEventLogSummary(larkEventLog);
  const parserSelfTest = options.parser_self_test ?? wsParserSelfTestPayload();
  const registeredEventTypes = ws.registered_event_types ?? [];
  const requiredDirectEvents = [
    meetingEventRequirements.direct_meeting.start_event,
    meetingEventRequirements.direct_meeting.end_event,
  ];
  const fallbackContextEvents = [
    meetingEventRequirements.meeting_context.start_event,
    meetingEventRequirements.meeting_context.end_event,
  ];
  const missingDirectEvents = requiredDirectEvents.filter((eventType) => !registeredEventTypes.includes(eventType));
  const missingFallbackEvents = fallbackContextEvents.filter((eventType) => !registeredEventTypes.includes(eventType));
  const receiverConnected = Boolean(ws.enabled && ws.state === 'connected');
  const publicWebhookReady = Boolean(config.event_callback_public_https);
  const receiverReady = receiverConnected || publicWebhookReady;
  const parserReady = Boolean(parserSelfTest.passed);
  const publicWebhookEvents = publicWebhookReady
    ? larkEventLog.filter((event) => event.transport === 'http_webhook')
    : [];
  const wsEventCount = Number(eventSummary.ws_event_count ?? config.recent_ws_event_count ?? 0);
  const realEventCount = wsEventCount + publicWebhookEvents.length;
  const candidateCount = Number(eventSummary.ws_timeline_candidate_count ?? 0)
    + publicWebhookEvents.filter((event) => event.timeline_candidate).length;
  const processedCount = Number(eventSummary.ws_timeline_processed_count ?? 0)
    + publicWebhookEvents.filter((event) => event.timeline_processed).length;
  const startedCount = Number(eventSummary.ws_timeline_started_count ?? 0)
    + publicWebhookEvents.filter((event) => event.timeline_started).length;
  const localHandlersReady = missingDirectEvents.length === 0;

  let status = 'unknown';
  let severity = 'info';
  let rootCause = 'unknown';
  let summary = '真实会议事件链路状态未知。';
  let nextAction = '查看 delivery diagnostics 和开放平台事件订阅配置。';

  if (!config.configured) {
    status = 'missing_app_credentials';
    severity = 'error';
    rootCause = 'local_configuration';
    summary = '缺少飞书应用凭证，无法连接事件接收通道。';
    nextAction = '配置 LARK_APP_ID / LARK_APP_SECRET 后重启服务。';
  } else if (!receiverReady) {
    status = 'receiver_not_ready';
    severity = 'error';
    rootCause = 'event_receiver';
    summary = `事件接收通道未就绪：长连接状态 ${ws.state ?? 'unknown'}，公网 HTTPS webhook=${publicWebhookReady}。`;
    nextAction = '优先使用长连接并确认页面显示 connected；或配置公网 HTTPS webhook。';
  } else if (!localHandlersReady) {
    status = 'local_handler_missing';
    severity = 'error';
    rootCause = 'local_event_dispatcher';
    summary = `本地接收器没有注册完整直开会议事件：缺少 ${missingDirectEvents.join(', ')}。`;
    nextAction = '修正 registered_event_types / createWsEventDispatcher。';
  } else if (!parserReady) {
    status = 'parser_self_test_failed';
    severity = 'error';
    rootCause = 'local_parser';
    summary = '本地 WS parser 自检失败，收到真实事件后也可能无法建轴。';
    nextAction = '先修复 /api/lark/ws-parser-self-test。';
  } else if (startedCount > 0) {
    status = 'event_delivery_ok';
    severity = 'ok';
    rootCause = null;
    summary = `已经收到并处理 ${startedCount} 条真实建轴事件。`;
    nextAction = '等待或写入开放标注，确认时间轴实时刷新。';
  } else if (realEventCount === 0) {
    status = 'no_event_delivery_observed';
    severity = 'info';
    rootCause = 'none_until_a_real_meeting_is_opened';
    summary = '事件接收通道已就绪、本地 handler 和 parser 自检均正常，但还没有观察到真实飞书事件；如果已经直接开过一次飞书会议，则优先怀疑开放平台事件订阅未保存/发布、权限未发布或应用可见范围。';
    nextAction = `直接开启一次飞书会议；如果真实事件计数仍为 0，在飞书开放平台确认事件订阅方式已保存为长连接，已添加并发布 ${requiredDirectEvents.join(' / ')}，同时确认 ${meetingEventRequirements.direct_meeting.required_scope} 已发布且应用可见范围包含当前开会账号。`;
  } else if (candidateCount === 0) {
    status = 'wrong_event_subscription';
    severity = 'warning';
    rootCause = 'open_platform_event_types';
    summary = `长连接已收到 ${wsEventCount} 条事件，但没有会议建轴候选；说明当前订阅/投递的不是会议开始类事件。`;
    nextAction = `订阅 ${requiredDirectEvents.join(' / ')}，兜底订阅 ${fallbackContextEvents.join(' / ')}。`;
  } else if (processedCount === 0) {
    status = 'parser_payload_gap';
    severity = 'warning';
    rootCause = 'payload_shape';
    summary = `长连接已收到 ${candidateCount} 条会议候选事件，但没有成功进入时间轴处理。`;
    nextAction = '查看最近候选事件 preview，补充 meeting_id/start_time 的 payload 路径解析。';
  } else {
    status = 'meeting_event_seen_but_no_axis';
    severity = 'warning';
    rootCause = 'meeting_start_anchor';
    summary = `已有 ${processedCount} 条会议事件被处理，但没有形成建轴事件。`;
    nextAction = '确认收到的是 start/join 类事件，且 payload 含 meeting_id 和 start_time。';
  }

  return {
    status,
    severity,
    root_cause: rootCause,
    summary,
    next_action: nextAction,
    receiver_ready: receiverReady,
    parser_self_test_passed: parserReady,
    local_handlers_ready: localHandlersReady,
    event_delivery_seen: realEventCount > 0,
    timeline_started_seen: startedCount > 0,
    evidence: {
      configured: Boolean(config.configured),
      ws_state: ws.state ?? null,
      ws_enabled: Boolean(ws.enabled),
      event_callback_public_https: publicWebhookReady,
      registered_event_types: registeredEventTypes,
      required_direct_events: requiredDirectEvents,
      missing_direct_events: missingDirectEvents,
      missing_fallback_events: missingFallbackEvents,
      real_event_count: realEventCount,
      ws_event_count: wsEventCount,
      public_webhook_event_count: publicWebhookEvents.length,
      ws_timeline_candidate_count: candidateCount,
      ws_timeline_processed_count: processedCount,
      ws_timeline_started_count: startedCount,
      last_ws_event_type: eventSummary.last_ws_event_type ?? null,
      parser_self_test: {
        passed: parserReady,
        direct_meeting_would_start: Boolean(parserSelfTest.cases?.direct_meeting_start?.would_start_timeline),
        context_join_would_start: Boolean(parserSelfTest.cases?.meeting_context_join?.would_start_timeline),
      },
    },
    required_open_platform_checks: [
      {
        id: 'event_subscription_mode',
        ok: receiverReady,
        required: '事件订阅使用长连接，或 HTTP 回调为公网 HTTPS。',
        observed: receiverConnected ? `long_connection ${ws.state}` : `webhook_public_https=${publicWebhookReady}`,
      },
      {
        id: 'direct_event_types',
        ok: localHandlersReady,
        required: requiredDirectEvents.join(' / '),
        observed: missingDirectEvents.length ? `missing ${missingDirectEvents.join(', ')}` : 'local handlers registered',
      },
      {
        id: 'parser_self_test',
        ok: parserReady,
        required: '合成 all_meeting_started 和 join_meeting 都能被解析成建轴候选。',
        observed: parserReady ? 'passed' : 'failed',
      },
      {
        id: 'platform_delivery',
        ok: realEventCount > 0,
        required: '直接开一次真实飞书会议后，事件日志至少出现一条 ws_long_connection 或公网 http_webhook 事件。',
        observed: `real_event_count=${realEventCount}, ws_event_count=${wsEventCount}, public_webhook_event_count=${publicWebhookEvents.length}`,
      },
      {
        id: 'subscription_saved_and_published',
        ok: realEventCount > 0,
        required: '开放平台事件订阅方式已保存为长连接，已添加直开会议事件，并且应用版本已发布到当前测试企业。',
        observed: realEventCount > 0 ? 'event delivery proves published subscription' : 'manual check required while ws_event_count=0',
      },
      {
        id: 'permission_and_release',
        ok: startedCount > 0,
        required: `${meetingEventRequirements.direct_meeting.required_scope} 已开通，并且应用/事件订阅已发布到当前测试企业。`,
        observed: startedCount > 0 ? `ws_timeline_started_count=${startedCount}` : 'cannot verify by local API until Feishu delivers an event',
        action_url: permissionUrlForScopes(meetingEventRequirements.direct_meeting.required_scope),
      },
      {
        id: 'app_visibility',
        ok: startedCount > 0,
        required: '应用可见范围包含当前登录并开会的用户。',
        observed: startedCount > 0 ? 'event delivered for current tenant/user context' : 'manual check required when ws_event_count=0',
      },
    ],
  };
}

function permissionUrlForScopes(scopes) {
  const value = Array.isArray(scopes) ? scopes.join(',') : String(scopes ?? '');
  return `https://open.feishu.cn/app/${process.env.LARK_APP_ID}/auth?q=${encodeURIComponent(value)}&op_from=openapi&token_type=tenant`;
}

function redactForLog(value, depth = 0) {
  if (value == null || typeof value !== 'object') return value;
  if (depth >= 4) return Array.isArray(value) ? `[array:${value.length}]` : '[object]';
  if (Array.isArray(value)) return value.slice(0, 8).map((item) => redactForLog(item, depth + 1));
  const output = {};
  for (const [key, child] of Object.entries(value).slice(0, 40)) {
    if (/token|secret|ticket|authorization|password|encrypt/i.test(key)) {
      output[key] = '[redacted]';
    } else {
      output[key] = redactForLog(child, depth + 1);
    }
  }
  return output;
}

function loadLarkEventLogState() {
  try {
    const parsed = JSON.parse(readFileSync(larkEventLogPath, 'utf8'));
    return Array.isArray(parsed) ? parsed.slice(0, 50) : [];
  } catch {
    return [];
  }
}

function persistLarkEventLog() {
  mkdirSync(dataDir, { recursive: true });
  writeFileSync(larkEventLogPath, JSON.stringify(larkEventLog.slice(0, 50), null, 2));
}

function pushLarkEventLog(entry) {
  const logged = {
    id: entry.id ?? `${Date.now()}-${larkEventLog.length}`,
    at: new Date().toISOString(),
    transport: entry.transport ?? 'unknown',
    event_type: String(entry.event_type || 'unknown'),
    timeline_candidate: Boolean(entry.timeline_candidate),
    timeline_processed: Boolean(entry.timeline_processed),
    timeline_started: Boolean(entry.timeline_started),
    ignored_reason: entry.ignored_reason ?? null,
    parsed_keys: entry.parsed && typeof entry.parsed === 'object'
      ? Object.keys(entry.parsed).filter((key) => !/token|secret|ticket|authorization/i.test(key)).slice(0, 30)
      : [],
    preview: redactForLog(entry.preview ?? entry.parsed ?? null),
  };
  larkEventLog.unshift(logged);
  larkEventLog.splice(50);
  persistLarkEventLog();
  return logged;
}

function isDirectMeetingStartEvent(eventType) {
  const text = String(eventType ?? '').toLowerCase();
  return text.includes('all_meeting') && (text.includes('start') || text.includes('begin'));
}

function isReserveMeetingStartEvent(eventType) {
  const text = String(eventType ?? '').toLowerCase();
  return !text.includes('all_meeting')
    && text.includes('meeting')
    && (text.includes('start') || text.includes('begin'));
}

function isMeetingContextBootstrapEvent(eventType) {
  const text = String(eventType ?? '').toLowerCase();
  return text.includes('vc.meeting.join_meeting');
}

function isoMs(value) {
  const ms = Date.parse(value ?? '');
  return Number.isFinite(ms) ? ms : null;
}

function realDeliveredTimelineEvents(config) {
  return larkEventLog.filter((event) => event.timeline_processed && (
    event.transport === 'ws_long_connection'
      || (event.transport === 'http_webhook' && config.event_callback_public_https)
  ));
}

function findMeetingEntryAfter(config, startedAt) {
  const startedMs = isoMs(startedAt) ?? 0;
  return realDeliveredTimelineEvents(config).find((event) => (
    event.timeline_started
      && (isDirectMeetingStartEvent(event.event_type) || isMeetingContextBootstrapEvent(event.event_type))
      && (isoMs(event.at) ?? 0) >= startedMs
  )) ?? null;
}

function eventsAtOrAfter(events = [], startedAt = null) {
  const startedMs = isoMs(startedAt);
  if (startedMs == null) return events;
  return events.filter((event) => (isoMs(event.at) ?? 0) >= startedMs);
}

function larkEventLogSummary(events = larkEventLog) {
  const wsEvents = events.filter((event) => event.transport === 'ws_long_connection');
  const wsCandidates = wsEvents.filter((event) => event.timeline_candidate);
  const wsProcessed = wsEvents.filter((event) => event.timeline_processed);
  const wsStarted = wsEvents.filter((event) => event.timeline_started);
  const wsIgnored = wsEvents.filter((event) => !event.timeline_candidate && !event.timeline_processed);
  return {
    event_count: events.length,
    ws_event_count: wsEvents.length,
    ws_timeline_candidate_count: wsCandidates.length,
    ws_timeline_processed_count: wsProcessed.length,
    ws_timeline_started_count: wsStarted.length,
    ws_ignored_count: wsIgnored.length,
    last_ws_event_type: wsEvents[0]?.event_type ?? null,
    last_ws_event_at: wsEvents[0]?.at ?? null,
    last_ws_event_status: wsEvents[0]
      ? wsEvents[0].timeline_started
        ? 'timeline_started'
        : wsEvents[0].timeline_processed
          ? 'timeline_processed'
          : wsEvents[0].timeline_candidate
            ? 'timeline_candidate'
            : 'ignored'
      : null,
    last_ws_timeline_candidate_event: wsCandidates[0] ?? null,
    last_ws_timeline_processed_event: wsProcessed[0] ?? null,
    last_ws_timeline_started_event: wsStarted[0] ?? null,
    last_ws_ignored_event: wsIgnored[0] ?? null,
  };
}

function eventLogEntryMs(entry = {}) {
  return isoMs(entry.at) ?? 0;
}

function loggedEventPayload(entry = {}) {
  return entry.preview ?? null;
}

function isLoggedMeetingStartEntry(entry = {}) {
  return Boolean(entry.timeline_started && (
    isDirectMeetingStartEvent(entry.event_type)
      || isReserveMeetingStartEvent(entry.event_type)
      || isMeetingContextBootstrapEvent(entry.event_type)
  ));
}

function isLoggedMeetingEndEntry(entry = {}) {
  const text = String(entry.event_type ?? '').toLowerCase();
  return text.includes('meeting') && (text.includes('ended') || text.includes('_end') || text.endsWith('.end'));
}

function loggedMeetingIdentity(entry = {}) {
  return explicitMeetingIdentity(loggedEventPayload(entry) ?? {});
}

function sameLoggedMeetingEntry(left = {}, right = {}) {
  const leftIdentity = loggedMeetingIdentity(left);
  const rightIdentity = loggedMeetingIdentity(right);
  return hasExplicitMeetingIdentity(leftIdentity)
    && sameExplicitMeeting(leftIdentity, rightIdentity);
}

function latestLoggedMeetingStartEvent() {
  return larkEventLog.find((entry) => isLoggedMeetingStartEntry(entry)) ?? null;
}

function latestLoggedMeetingStartBefore(cutoff) {
  const cutoffMs = isoMs(cutoff);
  if (cutoffMs == null) return null;
  return larkEventLog.find((entry) => (
    isLoggedMeetingStartEntry(entry)
      && eventLogEntryMs(entry) < cutoffMs
  )) ?? null;
}

function matchingLoggedMeetingEndAfter(startEntry = {}) {
  if (!startEntry) return null;
  const startMs = eventLogEntryMs(startEntry);
  return larkEventLog.find((entry) => (
    isLoggedMeetingEndEntry(entry)
      && eventLogEntryMs(entry) >= startMs
      && sameLoggedMeetingEntry(startEntry, entry)
  )) ?? null;
}

function missedMeetingWindowForProbe(req = null) {
  if (!realMeetingProbe.started_at) return null;
  const startedMs = isoMs(realMeetingProbe.started_at);
  const startEvent = latestLoggedMeetingStartBefore(realMeetingProbe.started_at);
  if (!startEvent || startedMs == null) return null;
  const endEvent = matchingLoggedMeetingEndAfter(startEvent);
  const endMs = endEvent ? eventLogEntryMs(endEvent) : null;
  const gapMs = startedMs - eventLogEntryMs(startEvent);
  return {
    reason: 'meeting_start_before_probe',
    message: endEvent && endMs <= startedMs
      ? '最近一次真实会议开始事件早于本次等待窗口，而且结束事件也已经到达；本次 probe 不会把它算作新会议。'
      : '最近一次真实会议开始事件早于本次等待窗口；如果这是你刚开的会议，可以从已投递事件恢复时间轴。',
    probe_started_at: realMeetingProbe.started_at,
    start_event: startEvent,
    end_event: endEvent,
    start_before_probe_ms: gapMs,
    ended_before_probe: Boolean(endEvent && endMs <= startedMs),
    restore_endpoint: req ? localUrlFor(req, '/api/lark/restore-latest-real-meeting-axis') : '/api/lark/restore-latest-real-meeting-axis',
  };
}

function realDemoAxisObservedAfterPrepare() {
  if (!realDemoSession.active) return true;
  const preparedMs = isoMs(realDemoSession.prepared_at);
  if (preparedMs == null) return true;
  const axisMs = isoMs(realDemoSession.last_real_axis_at);
  return axisMs != null && axisMs >= preparedMs;
}

const contentTypes = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
};

function corsHeaders(extra = {}) {
  return {
    'access-control-allow-origin': '*',
    'access-control-allow-methods': 'GET,POST,OPTIONS',
    'access-control-allow-headers': 'content-type, authorization, x-device-id, x-hmp-device-id, x-device-type, x-hmp-device-type',
    'access-control-max-age': '86400',
    ...extra,
  };
}

function sendNoContent(res, status = 204) {
  res.writeHead(status, corsHeaders());
  res.end();
}

function sendJson(res, status, data, headers = {}) {
  res.writeHead(status, corsHeaders({ 'content-type': 'application/json; charset=utf-8', ...headers }));
  res.end(JSON.stringify(data, null, 2));
}

function sendRedirect(res, location, status = 302) {
  res.writeHead(status, corsHeaders({ location, 'cache-control': 'no-cache' }));
  res.end();
}

function sendText(res, status, text, headers = {}) {
  res.writeHead(status, corsHeaders({ 'content-type': 'text/plain; charset=utf-8', ...headers }));
  res.end(text);
}

function sendHtml(res, status, html, headers = {}) {
  res.writeHead(status, corsHeaders({ 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-cache', ...headers }));
  res.end(html);
}

function loadAuthState() {
  try {
    const parsed = JSON.parse(readFileSync(authPath, 'utf8'));
    return {
      oauth_state: null,
      oauth_state_history: [],
      token: null,
      user: null,
      updated_at: null,
      ...parsed,
      oauth_state_history: pruneOAuthStateHistory(parsed.oauth_state_history ?? [], Date.now()),
    };
  } catch {
    return { oauth_state: null, oauth_state_history: [], token: null, user: null, updated_at: null };
  }
}

function saveAuthState(next) {
  authState = {
    oauth_state: null,
    oauth_state_history: [],
    token: null,
    user: null,
    ...next,
    oauth_state_history: pruneOAuthStateHistory(next.oauth_state_history ?? [], Date.now()),
    updated_at: new Date().toISOString(),
  };
  mkdirSync(dataDir, { recursive: true });
  writeFileSync(authPath, JSON.stringify(authState, null, 2));
  return authState;
}

function oauthStateCreatedAtMs(entry = {}) {
  const fromMs = Number(entry.created_at_ms);
  if (Number.isFinite(fromMs)) return fromMs;
  const parsed = Date.parse(entry.created_at ?? '');
  return Number.isNaN(parsed) ? null : parsed;
}

function pruneOAuthStateHistory(history = [], now = Date.now()) {
  const seen = new Set();
  return (Array.isArray(history) ? history : [])
    .map((entry) => {
      if (typeof entry === 'string') {
        return { state: entry, created_at_ms: now, created_at: new Date(now).toISOString() };
      }
      const createdAtMs = oauthStateCreatedAtMs(entry) ?? now;
      return {
        state: String(entry?.state ?? ''),
        created_at_ms: createdAtMs,
        created_at: entry?.created_at ?? new Date(createdAtMs).toISOString(),
      };
    })
    .filter((entry) => {
      if (!entry.state || seen.has(entry.state)) return false;
      if (now - entry.created_at_ms > oauthStateTtlMs) return false;
      seen.add(entry.state);
      return true;
    })
    .sort((left, right) => right.created_at_ms - left.created_at_ms)
    .slice(0, oauthStateHistoryLimit);
}

function rememberOAuthState(state, now = Date.now()) {
  const history = pruneOAuthStateHistory([
    { state, created_at_ms: now, created_at: new Date(now).toISOString() },
    ...(authState.oauth_state
      ? [{ state: authState.oauth_state, created_at: authState.updated_at ?? new Date(now).toISOString() }]
      : []),
    ...(authState.oauth_state_history ?? []),
  ], now);
  return saveAuthState({ ...authState, oauth_state: state, oauth_state_history: history });
}

function oauthStateMatches(receivedState, now = Date.now()) {
  const expected = [
    authState.oauth_state ? { state: authState.oauth_state, created_at: authState.updated_at } : null,
    ...(authState.oauth_state_history ?? []),
  ].filter(Boolean);
  if (!expected.length) return true;
  if (!receivedState) return false;
  return pruneOAuthStateHistory(expected, now).some((entry) => entry.state === receivedState);
}

function loadRealMeetingProbeState() {
  try {
    const parsed = JSON.parse(readFileSync(realMeetingProbePath, 'utf8'));
    return {
      ...defaultRealMeetingProbe,
      ...parsed,
      auto_search: {
        ...defaultRealMeetingProbe.auto_search,
        ...(parsed.auto_search ?? {}),
      },
    };
  } catch {
    return { ...defaultRealMeetingProbe };
  }
}

function saveRealMeetingProbeState(next) {
  realMeetingProbe = {
    ...defaultRealMeetingProbe,
    ...next,
    auto_search: {
      ...defaultRealMeetingProbe.auto_search,
      ...(next.auto_search ?? {}),
    },
    updated_at: new Date().toISOString(),
  };
  mkdirSync(dataDir, { recursive: true });
  writeFileSync(realMeetingProbePath, JSON.stringify(realMeetingProbe, null, 2));
  return realMeetingProbe;
}

function loadAutoAcceptanceState() {
  try {
    const parsed = JSON.parse(readFileSync(autoAcceptancePath, 'utf8'));
    return {
      ...defaultAutoAcceptance,
      ...parsed,
    };
  } catch {
    return { ...defaultAutoAcceptance };
  }
}

function saveAutoAcceptanceState(next) {
  autoAcceptance = {
    ...defaultAutoAcceptance,
    ...next,
    updated_at: new Date().toISOString(),
  };
  mkdirSync(dataDir, { recursive: true });
  writeFileSync(autoAcceptancePath, JSON.stringify(autoAcceptance, null, 2));
  return autoAcceptance;
}

function publicAutoAcceptanceStatus() {
  return {
    ...autoAcceptance,
    endpoint: '/api/acceptance/auto-annotation',
    note: '默认关闭。开启后，真实会议轴建立时会通过同一套开放标注链路自动写入一条验收标注。',
  };
}

function loadDeviceSimulatorState() {
  try {
    const parsed = JSON.parse(readFileSync(deviceSimulatorPath, 'utf8'));
    return {
      ...defaultDeviceSimulator,
      ...parsed,
    };
  } catch {
    return { ...defaultDeviceSimulator };
  }
}

function saveDeviceSimulatorState(next) {
  deviceSimulator = {
    ...defaultDeviceSimulator,
    ...next,
    updated_at: new Date().toISOString(),
  };
  mkdirSync(dataDir, { recursive: true });
  writeFileSync(deviceSimulatorPath, JSON.stringify(deviceSimulator, null, 2));
  return deviceSimulator;
}

function publicDeviceSimulatorStatus() {
  return {
    ...deviceSimulator,
    endpoint: '/api/device-simulator',
    stream_endpoint: '/api/device-simulator/stream',
    stream: publicDeviceStreamSimulatorStatus(),
    note: 'Demo 专用虚拟墨水屏。开启后，真实会议轴建立时会自动写入一条设备来源标注；生产设备仍应调用 POST /api/annotations。',
  };
}

function loadPassiveMeetingScanState() {
  try {
    const parsed = JSON.parse(readFileSync(passiveMeetingScanPath, 'utf8'));
    return {
      ...defaultPassiveMeetingScan,
      ...parsed,
    };
  } catch {
    return { ...defaultPassiveMeetingScan };
  }
}

function savePassiveMeetingScanState(next) {
  passiveMeetingScan = {
    ...defaultPassiveMeetingScan,
    ...next,
    updated_at: new Date().toISOString(),
  };
  mkdirSync(dataDir, { recursive: true });
  writeFileSync(passiveMeetingScanPath, JSON.stringify(passiveMeetingScan, null, 2));
  return passiveMeetingScan;
}

function publicPassiveMeetingScanStatus() {
  return {
    ...passiveMeetingScan,
    server_loop: {
      enabled: Boolean(passiveMeetingScan.enabled),
      scheduled: Boolean(passiveMeetingScanTimer),
      in_flight: passiveMeetingScanInFlight,
      interval_ms: passiveMeetingScanIntervalMs(),
    },
    endpoint: '/api/lark/passive-meeting-scan',
    note: passiveMeetingScan.enabled
      ? '已显式开启。当前 OAuth 含 vc:meeting.search:read 时，服务端会常驻扫描当前账号近期会议；发现正在进行的会议后自动建轴。'
      : '默认关闭。需要扫描兜底时，显式点击“扫描我的真实会议”、打开“被动扫描建轴”，或使用带 OAuth 的验收命令。',
  };
}

function loadRealDemoSessionState() {
  try {
    const parsed = JSON.parse(readFileSync(realDemoSessionPath, 'utf8'));
    return {
      ...defaultRealDemoSession,
      ...parsed,
    };
  } catch {
    return { ...defaultRealDemoSession };
  }
}

function saveRealDemoSessionState(next) {
  realDemoSession = {
    ...defaultRealDemoSession,
    ...next,
    updated_at: new Date().toISOString(),
  };
  mkdirSync(dataDir, { recursive: true });
  writeFileSync(realDemoSessionPath, JSON.stringify(realDemoSession, null, 2));
  return realDemoSession;
}

function publicRealDemoSessionStatus() {
  const autoOpenNote = realDemoSession.auto_open_session_on_annotation === false
    ? '当前已关闭首条标注自动开放会话，只等待飞书真实事件或显式会议会话入口。'
    : '首条可靠设备标注也可先建立开放会话轴。';
  return {
    ...realDemoSession,
    note: realDemoSession.active
      ? passiveMeetingScan.enabled
        ? `真实演示已准备：不会由按钮创建本地时间轴；当前已开启真实会议扫描兜底，可通过当前账号会议扫描绑定真实轴。${autoOpenNote}`
        : `真实演示已准备：不会由按钮创建时间轴；默认等待飞书真实会议事件建轴。${autoOpenNote}`
      : '真实演示未准备。',
  };
}

function tokenScopeList(value = authState.token?.scope) {
  return String(value ?? '').split(/\s+/).filter(Boolean);
}

function publicMeetingSearchAuthStatus(base = null) {
  const state = base ?? {
    authenticated: false,
    token_present: Boolean(authState.token?.access_token),
    expired: false,
    refresh_token_present: Boolean(authState.token?.refresh_token),
    scope: authState.token?.scope ?? null,
  };
  const scopePresent = tokenScopeList(state.scope).includes('vc:meeting.search:read');
  const usable = Boolean(state.authenticated && scopePresent);
  const needsReauth = !usable;
  let reason = null;
  let nextAction = null;
  if (!state.authenticated) {
    reason = state.token_present && state.expired ? 'oauth_token_expired' : 'oauth_login_required';
    nextAction = '重新登录飞书账号，并授予 vc:meeting.search:read';
  } else if (!scopePresent) {
    reason = 'missing_scope';
    nextAction = '重新授权飞书账号，追加 vc:meeting.search:read';
  }
  return {
    usable,
    required_scope: 'vc:meeting.search:read',
    scope_present: scopePresent,
    needs_reauth: needsReauth,
    reason,
    next_action: nextAction,
  };
}

function publicAuthState() {
  const expiresAtMs = authState.token?.obtained_at_ms && authState.token?.expires_in
    ? authState.token.obtained_at_ms + authState.token.expires_in * 1000
    : null;
  const refreshExpiresAtMs = authState.token?.obtained_at_ms && authState.token?.refresh_expires_in
    ? authState.token.obtained_at_ms + authState.token.refresh_expires_in * 1000
    : null;
  const expired = expiresAtMs != null && Date.now() >= expiresAtMs;
  const refreshExpired = refreshExpiresAtMs != null && Date.now() >= refreshExpiresAtMs;
  const expiresAt = expiresAtMs
    ? new Date(expiresAtMs).toISOString()
    : null;
  const refreshExpiresAt = refreshExpiresAtMs
    ? new Date(refreshExpiresAtMs).toISOString()
    : null;
  const state = {
    authenticated: Boolean(authState.token?.access_token) && !expired,
    token_present: Boolean(authState.token?.access_token),
    expired,
    refresh_token_present: Boolean(authState.token?.refresh_token),
    refresh_expired: refreshExpired,
    user: authState.user?.data ?? authState.user ?? null,
    scope: authState.token?.scope ?? null,
    scopes: tokenScopeList(),
    token_expires_at: expiresAt,
    refresh_token_expires_at: refreshExpiresAt,
    updated_at: authState.updated_at,
  };
  return {
    ...state,
    meeting_search: publicMeetingSearchAuthStatus(state),
  };
}

function validUserAccessToken() {
  return publicAuthState().authenticated ? authState.token.access_token : null;
}

async function ensureUserAccessToken() {
  const state = publicAuthState();
  if (state.authenticated) return authState.token.access_token;
  if (!authState.token?.refresh_token) return null;
  if (state.refresh_expired) return null;
  const refreshed = await lark.refreshOAuthToken(authState.token.refresh_token);
  const tokenWithTime = { ...refreshed, obtained_at_ms: Date.now() };
  saveAuthState({ ...authState, token: tokenWithTime });
  return tokenWithTime.access_token;
}

async function ensureMeetingSearchAuth() {
  const before = publicAuthState();
  if (before.meeting_search?.usable) {
    return { auth: before, refreshed: false, ok: true, reason: null, error: null };
  }
  if (!before.token_present || !before.expired || !before.refresh_token_present || before.refresh_expired) {
    return {
      auth: before,
      refreshed: false,
      ok: false,
      reason: before.meeting_search?.reason ?? 'meeting_search_unavailable',
      error: null,
    };
  }
  try {
    const token = await ensureUserAccessToken();
    const after = publicAuthState();
    return {
      auth: after,
      refreshed: Boolean(token),
      ok: Boolean(after.meeting_search?.usable),
      reason: after.meeting_search?.reason ?? null,
      error: null,
    };
  } catch (error) {
    return {
      auth: publicAuthState(),
      refreshed: false,
      ok: false,
      reason: 'oauth_refresh_failed',
      error: error.message ?? String(error),
    };
  }
}

function sendSse(res, event, data) {
  res.write(`event: ${event}\n`);
  res.write(`data: ${JSON.stringify(data)}\n\n`);
}

function stateBroadcastSummary(state = {}) {
  const binding = annotationBindingSummary(state);
  return {
    meeting_id: state.meeting?.meeting_id ?? null,
    meeting_source: state.meeting?.source ?? null,
    meeting_pending: Boolean(state.meeting?.pending_binding),
    real_meeting_axis_active: isRealMeetingAxis(state.meeting),
    sequence_count: state.sequence?.length ?? 0,
    open_annotation_count: binding.total,
    real_axis_annotation_count: binding.real_axis_count,
    last_annotation: binding.last,
  };
}

function rememberStreamBroadcast(state = {}, event = 'state') {
  streamTelemetry.broadcast_count += 1;
  streamTelemetry.last_broadcast_at = new Date().toISOString();
  streamTelemetry.last_event = event;
  streamTelemetry.last_client_count = streamClients.size;
  if (event === 'state') streamTelemetry.last_state_summary = stateBroadcastSummary(state);
}

function broadcastState(state, event = 'state') {
  const payload = event === 'state' ? publicTimelineState(state) : state;
  for (const res of streamClients) sendSse(res, event, payload);
  rememberStreamBroadcast(state, event);
}

async function saveAndBroadcast(nextState, event = 'state') {
  const saved = await store.save(nextState);
  broadcastState(saved, event);
  return saved;
}

function publicStreamStatus() {
  return {
    current_clients: streamClients.size,
    ...streamTelemetry,
  };
}

function realDemoPresentationForState(state = {}) {
  const meeting = state.meeting ?? {};
  const realAxis = isRealMeetingAxis(meeting);
  const axisObservedAfterPrepare = realDemoAxisObservedAfterPrepare();
  const realDemoWaiting = Boolean(realDemoSession.active && !axisObservedAfterPrepare);
  const pending = Boolean(meeting.pending_binding);
  const demoSample = meeting.meeting_id === 'demo-lark-meeting-001' && !meeting.source && !pending;
  const localSimulation = meeting.source === 'local_simulation' && !pending;
  const staleRealAxis = realDemoSession.active && realAxis && !axisObservedAfterPrepare;
  const hideTimeline = Boolean(
    realDemoWaiting
      && (demoSample || localSimulation || pending || staleRealAxis || !realAxis),
  );
  const hiddenReason = !hideTimeline
    ? null
    : pending
      ? 'pending_annotations_waiting_real_meeting_start'
      : demoSample
        ? 'demo_sample_axis_is_not_current_real_demo'
        : localSimulation
          ? 'local_simulation_hidden_during_real_demo'
          : staleRealAxis
            ? 'stale_real_axis_before_current_prepare'
            : 'waiting_for_real_lark_meeting_axis';
  const axisStatus = realAxis && axisObservedAfterPrepare
    ? 'real_axis_active'
    : pending
      ? 'pending_annotations_waiting_rebind'
      : realDemoWaiting
        ? 'waiting_real_meeting_start'
        : demoSample
          ? 'demo_sample'
          : localSimulation
            ? 'local_simulation'
            : 'unbound';
  return {
    hide_timeline: hideTimeline,
    hidden_reason: hiddenReason,
    axis_status: axisStatus,
    real_demo_session_active: Boolean(realDemoSession.active),
    real_demo_waiting: realDemoWaiting,
    real_axis_active: Boolean(realAxis && axisObservedAfterPrepare),
    raw_axis_real_meeting: realAxis,
    pending_binding: pending,
    pending_annotation_count: pending ? openAnnotationItems(state).length : 0,
    visible_sequence_count: hideTimeline ? 0 : state.sequence?.length ?? 0,
    note: hideTimeline
      ? '真实演示等待期间不展示样例/临时轴，避免误认为已按飞书会议开始时间建轴。'
      : null,
  };
}

function publicTimelineState(state = {}) {
  return {
    ...state,
    presentation: realDemoPresentationForState(state),
    real_demo_session: publicRealDemoSessionStatus(),
  };
}

function appendAnnotationToState(current, input, fallbackNow = Date.now()) {
  const createdPendingTimeline = shouldCreatePendingTimeline(current, input);
  const working = createdPendingTimeline
    ? buildPendingAnnotationTimeline(current, input, fallbackNow)
    : current;
  const item = normalizeAnnotationEvent(input, working.meeting, fallbackNow);
  const existing = working.sequence ?? [];
  const replacedExisting = existing.some((x) => x.id === item.id);
  const deduped = existing.filter((x) => x.id !== item.id);
  const next = mergeTimeline(working, { sequence: [...deduped, item] });
  return {
    state: next,
    item,
    options: {
      created_pending_timeline: createdPendingTimeline,
      replaced_existing: replacedExisting,
      idempotency_key_source: annotationIdempotencyKeySource(input),
    },
  };
}

function isRealtimeAnnotationInput(input = {}) {
  return Boolean(
    input.realtime === true
      || input.live === true
      || input.mode === 'realtime'
      || input.payload?.realtime === true
      || input.payload?.mode === 'realtime',
  );
}

function currentAxisIsStaleForRealtimeAnnotation(state = {}) {
  return Boolean(
    realDemoSession.active
      && isRealMeetingAxis(state.meeting)
      && !realDemoAxisObservedAfterPrepare(),
  );
}

function shouldStartFreshRealtimeAxisForAnnotation(state = {}, input = {}) {
  if (!isRealtimeAnnotationInput(input)) return false;
  if (state.meeting?.pending_binding) return false;
  if (!state.meeting?.meeting_id || state.meeting?.meeting_id === 'demo-lark-meeting-001') return true;
  if (currentAxisIsStaleForRealtimeAnnotation(state)) return true;
  if (isRealMeetingAxis(state.meeting) && state.meeting?.end_time) return true;
  return false;
}

function shouldAttemptPassiveBindForAnnotation(current = {}, input = {}, now = Date.now()) {
  if (!passiveMeetingScan.enabled) return { ok: false, reason: 'passive_scan_disabled' };
  if (annotationCapturedAbsoluteMsStrict(input) == null) return { ok: false, reason: 'missing_captured_at_ms' };
  if (
    isRealMeetingAxis(current.meeting)
    && !current.meeting?.end_time
    && !shouldStartFreshRealtimeAxisForAnnotation(current, input)
  ) {
    return { ok: false, reason: 'real_axis_already_active' };
  }
  if (isRealMeetingProbeWaiting(now) && realMeetingProbe.auto_search?.enabled) {
    return { ok: false, reason: 'probe_auto_search_active' };
  }
  const auth = publicAuthState();
  if (
    !auth.meeting_search?.usable
    && auth.token_present
    && auth.expired
    && auth.refresh_token_present
    && !auth.refresh_expired
  ) {
    return { ok: true, reason: null };
  }
  if (!auth.meeting_search?.usable) return { ok: false, reason: auth.meeting_search?.reason ?? 'meeting_search_unavailable' };
  return { ok: true, reason: null };
}

async function maybeBindCurrentMeetingForIncomingAnnotation(input = {}, current = null) {
  const now = Date.now();
  const state = current ?? await store.load();
  const decision = shouldAttemptPassiveBindForAnnotation(state, input, now);
  if (!decision.ok) return { attempted: false, reason: decision.reason };
  const result = await autoBindPassiveMeeting({
    source: 'annotation_arrival_passive_scan',
    force: true,
    lookback_seconds: passiveMeetingScan.lookback_seconds,
    lookahead_seconds: passiveMeetingScan.lookahead_seconds,
  });
  return {
    attempted: true,
    status: result.status,
    reason: result.reason ?? null,
    selected_meeting_id: result.selected_meeting_id ?? result.state?.meeting?.meeting_id ?? null,
    meeting_title: result.state?.meeting?.title ?? null,
    bound: result.status === 'bound',
  };
}

function autoOpenSessionDecisionForAnnotation(current = {}, input = {}, now = Date.now()) {
  if (!realDemoSession.active) return { ok: false, reason: 'real_demo_inactive' };
  if (realDemoSession.auto_open_session_on_annotation === false) {
    return { ok: false, reason: 'auto_open_session_on_annotation_disabled' };
  }
  if (annotationMeetingSessionBody(input)) return { ok: false, reason: 'explicit_meeting_session_context' };
  const capturedAtMs = annotationCapturedAbsoluteMsStrict(input);
  if (capturedAtMs == null) return { ok: false, reason: 'missing_captured_at_ms' };
  const shouldStartFreshRealtimeAxis = shouldStartFreshRealtimeAxisForAnnotation(current, input);
  if (isRealMeetingAxis(current.meeting) && !current.meeting?.end_time && !shouldStartFreshRealtimeAxis) {
    return {
      ok: false,
      reason: current.meeting?.source === 'open_meeting_session'
        ? 'open_meeting_session_already_active'
        : 'real_axis_already_active',
    };
  }
  const source = current.meeting?.source ?? null;
  const canReplaceCurrentAxis = Boolean(
    !hasActiveMeeting(current)
      || current.meeting?.pending_binding
      || current.meeting?.meeting_id === 'demo-lark-meeting-001'
      || source === 'annotation_fallback'
      || source === 'local_simulation'
      || source === 'lark_reserve_pending'
      || shouldStartFreshRealtimeAxis
  );
  if (!canReplaceCurrentAxis) return { ok: false, reason: 'concrete_axis_already_active' };
  return { ok: true, captured_at_ms: capturedAtMs, force: shouldStartFreshRealtimeAxis, now };
}

async function maybeAutoOpenSessionForIncomingAnnotation(input = {}, current = null) {
  const state = current ?? await store.load();
  const decision = autoOpenSessionDecisionForAnnotation(state, input);
  if (!decision.ok) return { attempted: false, reason: decision.reason };
  const result = await startOpenMeetingSession({
    platform: 'lark',
    meeting_id: firstNonEmpty(
      input.meeting_id,
      input.session_id,
      input.context?.meeting_id,
      `annotation-open-session-${decision.captured_at_ms}`,
    ),
    external_meeting_id: firstNonEmpty(input.external_meeting_id, input.meeting_no, null),
    meeting_url: firstNonEmpty(input.meeting_url, input.url, input.join_url, null),
    title: String(firstNonEmpty(
      input.meeting_title,
      input.context?.meeting_title,
      input.context?.title,
      '实时标注开放会话',
    )),
    start_time_ms: decision.captured_at_ms,
    detector_source: firstNonEmpty(input.detector_source, input.source, input.device?.type, 'first_reliable_annotation'),
    note: 'auto_open_session_on_first_reliable_annotation',
    force: decision.force,
    suppress_auto_annotations: true,
  });
  return {
    attempted: true,
    status: 'started',
    reason: null,
    meeting_id: result.meeting?.meeting_id ?? null,
    source: result.meeting?.source ?? null,
    start_time_ms: decision.captured_at_ms,
    state: result.state,
  };
}

async function appendAnnotation(input) {
  const meetingSessionBinding = await maybeStartOpenMeetingSessionFromAnnotation(input);
  let passiveBinding = { attempted: false, reason: `meeting_session_${meetingSessionBinding.status ?? 'attempted'}` };
  let autoOpenSessionBinding = { attempted: false, reason: `meeting_session_${meetingSessionBinding.status ?? 'attempted'}` };
  if (!meetingSessionBinding.attempted) {
    passiveBinding = await maybeBindCurrentMeetingForIncomingAnnotation(input);
    const currentAfterPassive = await store.load();
    autoOpenSessionBinding = passiveBinding.bound
      ? { attempted: false, reason: 'passive_scan_bound_real_axis' }
      : await maybeAutoOpenSessionForIncomingAnnotation(input, currentAfterPassive);
  }
  const current = await store.load();
  const result = appendAnnotationToState(current, input);
  const saved = await saveAndBroadcast(result.state, 'state');
  return {
    ack: annotationAck(result.item, saved, result.options),
    meeting_session_binding: meetingSessionBinding,
    passive_binding: passiveBinding,
    auto_open_session_binding: autoOpenSessionBinding,
    item: result.item,
    state: saved,
  };
}

async function appendAnnotationBatch(inputs = []) {
  if (!Array.isArray(inputs) || inputs.length === 0) {
    const error = new Error('annotations array is required');
    error.status = 400;
    throw error;
  }
  if (inputs.length > 200) {
    const error = new Error('annotations batch limit is 200');
    error.status = 413;
    throw error;
  }
  let working = await store.load();
  const firstMeetingSessionInput = inputs.find((input) => annotationMeetingSessionBody(input));
  const meetingSessionBinding = firstMeetingSessionInput
    ? await maybeStartOpenMeetingSessionFromAnnotation(firstMeetingSessionInput)
    : { attempted: false, reason: 'missing_meeting_session_context' };
  working = await store.load();
  const firstReliableInput = inputs.find((input) => annotationCapturedAbsoluteMsStrict(input) != null) ?? inputs[0];
  let passiveBinding = meetingSessionBinding.attempted
    ? { attempted: false, reason: `meeting_session_${meetingSessionBinding.status ?? 'attempted'}` }
    : { attempted: false, reason: 'empty_batch' };
  let autoOpenSessionBinding = meetingSessionBinding.attempted
    ? { attempted: false, reason: `meeting_session_${meetingSessionBinding.status ?? 'attempted'}` }
    : { attempted: false, reason: 'empty_batch' };
  if (!meetingSessionBinding.attempted && firstReliableInput) {
    passiveBinding = await maybeBindCurrentMeetingForIncomingAnnotation(firstReliableInput, working);
    working = await store.load();
    autoOpenSessionBinding = passiveBinding.bound
      ? { attempted: false, reason: 'passive_scan_bound_real_axis' }
      : await maybeAutoOpenSessionForIncomingAnnotation(firstReliableInput, working);
  }
  working = await store.load();
  const results = [];
  for (const input of inputs) {
    const result = appendAnnotationToState(working, input);
    working = result.state;
    results.push(result);
  }
  const saved = await saveAndBroadcast(working, 'state');
  const rows = results.map((result) => ({
    ack: annotationAck(result.item, saved, result.options),
    item: result.item,
  }));
  return {
    accepted: true,
    count: rows.length,
    acks: rows.map((row) => row.ack),
    items: rows.map((row) => row.item),
    meeting_session_binding: meetingSessionBinding,
    passive_binding: passiveBinding,
    auto_open_session_binding: autoOpenSessionBinding,
    state: saved,
  };
}

function autoAcceptanceMeetingKey(meeting = {}) {
  return String(
    meeting.meeting_id
      ?? meeting.external_meeting_id
      ?? meeting.minute_token
      ?? meeting.meeting_url
      ?? 'unknown-meeting',
  ).replace(/[^\w.-]+/g, '_');
}

async function maybeAppendAutoAcceptanceAnnotation(state = {}, triggerSource = 'unknown') {
  if (!autoAcceptance.enabled) return { state, annotation: null, skipped_reason: 'disabled' };
  if (!isRealMeetingAxis(state.meeting)) {
    return { state, annotation: null, skipped_reason: 'not_real_meeting_axis' };
  }
  if (state.meeting?.end_time) {
    return { state, annotation: null, skipped_reason: 'meeting_already_ended' };
  }
  const meetingKey = autoAcceptanceMeetingKey(state.meeting);
  const annotationId = `auto-acceptance-${meetingKey}`;
  if ((state.sequence ?? []).some((item) => item.id === annotationId)) {
    return { state, annotation: null, skipped_reason: 'already_exists' };
  }

  const now = Date.now();
  const input = {
    id: annotationId,
    source: 'demo_auto_acceptance',
    captured_at_ms: now,
    kind: 'acceptance_mark',
    label: autoAcceptance.label || defaultAutoAcceptance.label,
    text_candidates: [autoAcceptance.label || defaultAutoAcceptance.label, 'auto acceptance mark'],
    intent: 'acceptance_check',
    strokes: [],
    payload: {
      origin: 'auto_acceptance',
      trigger_source: triggerSource,
      meeting_id: state.meeting?.meeting_id ?? null,
    },
  };
  const result = appendAnnotationToState(state, input, now);
  const saved = await saveAndBroadcast(result.state, 'state');
  saveAutoAcceptanceState({
    ...autoAcceptance,
    count: Number(autoAcceptance.count ?? 0) + 1,
    last_annotation_id: annotationId,
    last_meeting_id: state.meeting?.meeting_id ?? null,
    last_trigger_at: new Date(now).toISOString(),
  });
  return { state: saved, annotation: result.item, skipped_reason: null };
}

async function maybeAppendDeviceSimulatorAnnotation(state = {}, triggerSource = 'unknown') {
  if (!deviceSimulator.enabled) return { state, annotation: null, skipped_reason: 'disabled' };
  if (!isRealMeetingAxis(state.meeting)) {
    return { state, annotation: null, skipped_reason: 'not_real_meeting_axis' };
  }
  if (state.meeting?.end_time) {
    return { state, annotation: null, skipped_reason: 'meeting_already_ended' };
  }
  const meetingKey = autoAcceptanceMeetingKey(state.meeting);
  const annotationId = `device-simulator-${meetingKey}`;
  if ((state.sequence ?? []).some((item) => item.id === annotationId)) {
    return { state, annotation: null, skipped_reason: 'already_exists' };
  }

  const now = Date.now();
  const label = String(deviceSimulator.label || defaultDeviceSimulator.label);
  const input = {
    id: annotationId,
    source: deviceSimulator.device_type || defaultDeviceSimulator.device_type,
    device_id: deviceSimulator.device_id || defaultDeviceSimulator.device_id,
    device: {
      id: deviceSimulator.device_id || defaultDeviceSimulator.device_id,
      type: deviceSimulator.device_type || defaultDeviceSimulator.device_type,
    },
    captured_at_ms: now,
    kind: 'handwriting_trigger',
    label,
    text_candidates: [label, 'why?', 'why'],
    intent: 'question',
    strokes: [],
    payload: {
      origin: 'real_demo_device_simulator',
      trigger_source: triggerSource,
      meeting_id: state.meeting?.meeting_id ?? null,
      page_id: 'lark-meeting-demo',
    },
  };
  const result = appendAnnotationToState(state, input, now);
  const saved = await saveAndBroadcast(result.state, 'state');
  saveDeviceSimulatorState({
    ...deviceSimulator,
    count: Number(deviceSimulator.count ?? 0) + 1,
    last_annotation_id: annotationId,
    last_meeting_id: state.meeting?.meeting_id ?? null,
    last_trigger_at: new Date(now).toISOString(),
  });
  return { state: saved, annotation: result.item, skipped_reason: null };
}

function publicDeviceStreamSimulatorStatus() {
  return {
    ...deviceStreamState,
    endpoint: '/api/device-simulator/stream',
    timer_active: Boolean(deviceStreamTimer),
    in_flight: Boolean(deviceStreamInFlight),
    note: 'Demo 专用流式设备模拟器。可在会议前启动，真实会议轴出现后会按间隔通过开放标注链路写入多条设备标注。',
  };
}

function stopDeviceStreamSimulator(reason = 'stopped') {
  if (deviceStreamTimer) {
    clearTimeout(deviceStreamTimer);
    deviceStreamTimer = null;
  }
  deviceStreamState = {
    ...deviceStreamState,
    enabled: false,
    status: reason,
    stopped_at: new Date().toISOString(),
  };
  return publicDeviceStreamSimulatorStatus();
}

function scheduleDeviceStreamSimulator(delayMs = deviceStreamState.interval_ms) {
  if (!deviceStreamState.enabled) return;
  if (deviceStreamTimer) clearTimeout(deviceStreamTimer);
  deviceStreamTimer = setTimeout(runDeviceStreamSimulatorTick, Math.max(250, Number(delayMs) || 1000));
  deviceStreamTimer.unref?.();
}

function nextDeviceStreamLabel(count) {
  const samples = [
    'why?',
    'action item',
    'follow up',
    'important',
    'check later',
  ];
  return `${deviceStreamState.label_prefix || '流式设备标注'} ${count}: ${samples[(count - 1) % samples.length]}`;
}

async function runDeviceStreamSimulatorTick() {
  if (!deviceStreamState.enabled) return;
  if (deviceStreamInFlight) {
    scheduleDeviceStreamSimulator();
    return;
  }
  deviceStreamInFlight = true;
  try {
    const current = await store.load();
	    if (!isRealMeetingAxis(current.meeting) || !realDemoAxisObservedAfterPrepare()) {
	      deviceStreamState = {
	        ...deviceStreamState,
	        status: 'waiting_for_real_axis',
        last_error: null,
      };
      scheduleDeviceStreamSimulator();
      return;
    }
    if (current.meeting?.end_time) {
      stopDeviceStreamSimulator('meeting_ended');
      return;
    }
    const nextCount = Number(deviceStreamState.count ?? 0) + 1;
    if (nextCount > Number(deviceStreamState.max_count ?? 1)) {
      stopDeviceStreamSimulator('complete');
      return;
    }
    const now = Date.now();
    const annotationId = `device-stream-${autoAcceptanceMeetingKey(current.meeting)}-${nextCount}`;
    const label = nextDeviceStreamLabel(nextCount);
    const result = await appendAnnotation({
      id: annotationId,
      source: deviceSimulator.device_type || defaultDeviceSimulator.device_type,
      device_id: deviceSimulator.device_id || defaultDeviceSimulator.device_id,
      device: {
        id: deviceSimulator.device_id || defaultDeviceSimulator.device_id,
        type: deviceSimulator.device_type || defaultDeviceSimulator.device_type,
      },
      captured_at_ms: now,
      kind: 'handwriting_trigger',
      label,
      text_candidates: [label, 'why?', 'action item'],
      intent: nextCount % 2 === 1 ? 'question' : 'attention',
      strokes: [
        [
          { x: 0.1 + nextCount * 0.01, y: 0.2, t: now - 120 },
          { x: 0.2 + nextCount * 0.01, y: 0.24, t: now },
        ],
      ],
      payload: {
        origin: 'real_demo_device_stream_simulator',
        meeting_id: current.meeting?.meeting_id ?? null,
        page_id: 'lark-meeting-demo',
      },
    });
    deviceStreamState = {
      ...deviceStreamState,
      status: nextCount >= Number(deviceStreamState.max_count ?? 1) ? 'complete' : 'streaming',
      enabled: nextCount < Number(deviceStreamState.max_count ?? 1),
      count: nextCount,
      last_annotation_id: annotationId,
      last_meeting_id: current.meeting?.meeting_id ?? null,
      last_trigger_at: new Date(now).toISOString(),
      last_error: null,
    };
    if (deviceStreamState.enabled) scheduleDeviceStreamSimulator();
    else stopDeviceStreamSimulator('complete');
    return result;
  } catch (error) {
    deviceStreamState = {
      ...deviceStreamState,
      status: 'error',
      last_error: error.message ?? String(error),
    };
    scheduleDeviceStreamSimulator(Math.max(1000, deviceStreamState.interval_ms));
  } finally {
    deviceStreamInFlight = false;
  }
}

async function startDeviceStreamSimulator(body = {}) {
  if (deviceStreamTimer) clearTimeout(deviceStreamTimer);
  const intervalMs = Math.min(Math.max(Number(body.interval_ms ?? body.interval ?? deviceStreamState.interval_ms ?? 1500), 250), 60_000);
  const maxCount = Math.min(Math.max(Number(body.max_count ?? body.count ?? deviceStreamState.max_count ?? 5), 1), 200);
  deviceStreamState = {
    ...deviceStreamState,
    enabled: true,
    status: 'starting',
    interval_ms: intervalMs,
    max_count: maxCount,
    count: 0,
    started_at: new Date().toISOString(),
    stopped_at: null,
    last_annotation_id: null,
    last_meeting_id: null,
    last_trigger_at: null,
    last_error: null,
    label_prefix: String(body.label_prefix || body.label || deviceStreamState.label_prefix || '流式设备标注'),
  };
  scheduleDeviceStreamSimulator(0);
  return publicDeviceStreamSimulatorStatus();
}

async function resumeRealDemoDeviceStreamOnStartup() {
  if (!realDemoSession.active) return { resumed: false, reason: 'real_demo_inactive' };
  if (!deviceSimulator.enabled) return { resumed: false, reason: 'device_simulator_disabled' };
  const current = await store.load();
  const preparedAt = realDemoSession.prepared_at ?? null;
  const binding = annotationBindingSummary(current, { since_at: preparedAt });
  if (isRealMeetingAxis(current.meeting) && realDemoAxisObservedAfterPrepare() && binding.since_real_axis_count > 0) {
    return { resumed: false, reason: 'real_demo_already_has_real_axis_annotation' };
  }
  const stream = await startDeviceStreamSimulator({
    interval_ms: deviceStreamState.interval_ms,
    max_count: deviceStreamState.max_count,
    label_prefix: deviceStreamState.label_prefix || '流式设备标注',
  });
  return { resumed: true, reason: 'real_demo_active', stream };
}

async function prepareRealDemoRuntime(body = {}, req = null) {
  const autoAnnotationEnabled = body.auto_annotation !== false;
  const deviceSimulatorEnabled = body.device_simulator !== false;
  const deviceStreamEnabled = body.device_stream ?? deviceSimulatorEnabled;
  const passiveScanEnabled = body.passive_scan !== false;
  const temporaryAxisReset = await resetTemporaryAxisForProbeIfRequested({
    reset_temporary_axis: body.reset_temporary_axis !== false,
  });
  saveRealDemoSessionState({
    ...realDemoSession,
    active: true,
    auto_open_session_on_annotation: body.auto_open_session_on_annotation
      ?? body.auto_open_on_annotation
      ?? realDemoSession.auto_open_session_on_annotation
      ?? defaultRealDemoSession.auto_open_session_on_annotation,
    prepared_at: new Date().toISOString(),
    last_real_axis_at: null,
    last_real_axis_source: null,
    last_note: body.note ? String(body.note) : 'real_demo_prepare',
  });
  if (autoAnnotationEnabled) {
    saveAutoAcceptanceState({
      ...autoAcceptance,
      enabled: true,
      label: String(body.auto_annotation_label || autoAcceptance.label || defaultAutoAcceptance.label),
    });
  } else {
    saveAutoAcceptanceState({
      ...autoAcceptance,
      enabled: false,
    });
  }
  if (deviceSimulatorEnabled) {
    saveDeviceSimulatorState({
      ...deviceSimulator,
      enabled: true,
      label: String(body.device_simulator_label || deviceSimulator.label || defaultDeviceSimulator.label),
      device_id: String(body.device_id || deviceSimulator.device_id || defaultDeviceSimulator.device_id),
      device_type: String(body.device_type || deviceSimulator.device_type || defaultDeviceSimulator.device_type),
    });
  } else {
    saveDeviceSimulatorState({
      ...deviceSimulator,
      enabled: false,
    });
  }
  const deviceStream = deviceStreamEnabled
    ? await startDeviceStreamSimulator({
      interval_ms: body.device_stream_interval_ms ?? body.device_stream_interval ?? 1500,
      max_count: body.device_stream_max_count ?? body.device_stream_count ?? 5,
      label_prefix: body.device_stream_label_prefix ?? '流式设备标注',
    })
    : stopDeviceStreamSimulator('disabled_by_prepare');
  if (passiveScanEnabled) {
    savePassiveMeetingScanState({
      ...passiveMeetingScan,
      enabled: true,
      tenant_fallback_enabled: body.tenant_fallback_enabled ?? passiveMeetingScan.tenant_fallback_enabled ?? defaultPassiveMeetingScan.tenant_fallback_enabled,
      interval_ms: Math.min(Math.max(Number(body.interval_ms ?? passiveMeetingScan.interval_ms ?? 10_000), 5000), 5 * 60_000),
      lookback_seconds: Math.min(Math.max(Number(body.lookback_seconds ?? passiveMeetingScan.lookback_seconds ?? 600), 60), 2 * 60 * 60),
      lookahead_seconds: Math.min(Math.max(Number(body.lookahead_seconds ?? passiveMeetingScan.lookahead_seconds ?? 120), 0), 30 * 60),
    });
    schedulePassiveMeetingScanLoop(1000);
  } else {
    savePassiveMeetingScanState({
      ...passiveMeetingScan,
      enabled: false,
      updated_at: new Date().toISOString(),
    });
    stopPassiveMeetingScanLoop();
  }

  const authAttempt = await ensureMeetingSearchAuth();
  const auth = authAttempt.auth;
  const trigger = auth.meeting_search?.usable && passiveScanEnabled
    ? publicPassiveScanResult(await runImmediatePassiveMeetingScan(body.prepare_source || 'real_demo_prepare'))
    : null;
  const authStart = req && !auth.meeting_search?.usable
    ? meetingSearchOAuthStart(req)
    : null;
  return {
    prepared: true,
    generated_at: new Date().toISOString(),
    auth_required: !auth.meeting_search?.usable,
    auth_refresh: {
      attempted: authAttempt.refreshed || authAttempt.reason === 'oauth_refresh_failed',
      refreshed: authAttempt.refreshed,
      reason: authAttempt.reason,
      error: authAttempt.error,
    },
    auth: publicAuthState(),
    auth_start: authStart,
    passive_meeting_scan: publicPassiveMeetingScanStatus(),
    auto_acceptance: publicAutoAcceptanceStatus(),
    device_simulator: publicDeviceSimulatorStatus(),
    device_stream_simulator: deviceStream,
    real_demo_session: publicRealDemoSessionStatus(),
    temporary_axis_reset: temporaryAxisReset,
    trigger,
    next_step: passiveScanEnabled
      ? auth.meeting_search?.usable
        ? '现在直接开启飞书会议；被动扫描会自动绑定正在进行的会议，真实轴建立后设备流会连续写入开放标注。'
        : '现在可直接开启飞书会议等待事件建轴；如需扫描兜底，请先完成返回的飞书授权 URL。'
      : '现在直接开启飞书会议；服务端只等待飞书真实事件建轴。扫描兜底需要单独点击“扫描我的真实会议”、打开“被动扫描建轴”，或使用带 OAuth 的验收命令。',
  };
}

function realDemoAutoArmOptionsFromEnv() {
  return {
    auto_annotation: boolFromEnv(process.env.REAL_DEMO_AUTO_ANNOTATION, false),
    device_simulator: boolFromEnv(process.env.REAL_DEMO_DEVICE_SIMULATOR, false),
    device_stream: boolFromEnv(process.env.REAL_DEMO_DEVICE_STREAM, false),
    passive_scan: boolFromEnv(process.env.REAL_DEMO_PASSIVE_SCAN, true),
    auto_open_session_on_annotation: boolFromEnv(process.env.REAL_DEMO_AUTO_OPEN_SESSION_ON_ANNOTATION, true),
    tenant_fallback_enabled: boolFromEnv(process.env.REAL_DEMO_TENANT_FALLBACK_SCAN, false),
    reset_temporary_axis: boolFromEnv(process.env.REAL_DEMO_RESET_TEMPORARY_AXIS, true),
    device_stream_interval_ms: boundedNumber(process.env.REAL_DEMO_DEVICE_STREAM_INTERVAL_MS, 1500, 200, 60_000),
    device_stream_max_count: boundedNumber(process.env.REAL_DEMO_DEVICE_STREAM_COUNT, 5, 1, 200),
    device_stream_label_prefix: process.env.REAL_DEMO_DEVICE_STREAM_LABEL_PREFIX || '流式设备标注',
    note: 'real_demo_auto_arm_startup',
    prepare_source: 'real_demo_auto_arm_startup',
  };
}

async function maybeAutoArmRealDemoOnStartup() {
  if (!boolFromEnv(process.env.REAL_DEMO_AUTO_ARM, false)) {
    return { armed: false, reason: 'REAL_DEMO_AUTO_ARM_disabled' };
  }
  const result = await prepareRealDemoRuntime(realDemoAutoArmOptionsFromEnv());
  return { armed: true, reason: 'REAL_DEMO_AUTO_ARM_enabled', result };
}

function annotationIdempotencyKeySource(input = {}) {
  for (const field of ['id', 'annotation_id', 'mark_id', 'event_id']) {
    if (input[field] != null && input[field] !== '') return field;
  }
  return null;
}

function hasActiveMeeting(state = {}) {
  return Boolean(
    state.meeting?.start_time
      && !state.meeting?.end_time
      && !state.meeting?.pending_binding
      && state.meeting?.source !== 'lark_reserve_pending',
  );
}

function hasConcreteMeetingForAnnotation(state = {}) {
  return Boolean(
    state.meeting?.start_time
      && state.meeting?.meeting_id
      && state.meeting?.meeting_id !== 'demo-lark-meeting-001'
      && !state.meeting?.pending_binding
      && state.meeting?.source !== 'lark_reserve_pending',
  );
}

function isRealMeetingProbeWaiting(now = Date.now()) {
  if (!realMeetingProbe.active || !realMeetingProbe.started_at) return false;
  const startedMs = isoMs(realMeetingProbe.started_at);
  if (startedMs == null) return false;
  return Math.max(0, now - startedMs) <= realMeetingProbe.timeout_ms;
}

function shouldCreatePendingTimeline(state = {}, input = {}) {
  const hasReliableCaptureTime = annotationCapturedAbsoluteMsStrict(input) != null;
  if (!hasReliableCaptureTime) return false;
  if (input.force_new_session || input.new_session) return true;
  if (state.meeting?.source === 'lark_reserve_pending') return true;
  if (state.meeting?.pending_binding) return false;
  if (shouldStartFreshRealtimeAxisForAnnotation(state, input)) return true;
  if (!state.meeting?.meeting_id) return true;
  if (state.meeting?.meeting_id === 'demo-lark-meeting-001') return true;
  if (isRealMeetingProbeWaiting() && !isRealMeetingAxis(state.meeting)) return true;
  if (hasConcreteMeetingForAnnotation(state)) return false;
  return !hasActiveMeeting(state);
}

function shouldCarryMeetingIdentityIntoPending(meeting = {}) {
  if (!meeting?.meeting_id) return false;
  if (meeting.meeting_id === 'demo-lark-meeting-001') return false;
  if (meeting.source === 'local_simulation') return false;
  return true;
}

function buildPendingAnnotationTimeline(current, input, fallbackNow) {
  const startMs = annotationCapturedAbsoluteMsStrict(input) ?? annotationCapturedAbsoluteMs(input, fallbackNow);
  const carryIdentity = shouldCarryMeetingIdentityIntoPending(current.meeting);
  return buildTimeline({
    meeting: {
      platform: 'lark',
      meeting_id: `pending-live-${startMs}`,
      external_meeting_id: carryIdentity ? current.meeting?.external_meeting_id ?? null : null,
      meeting_url: carryIdentity ? current.meeting?.meeting_url ?? null : null,
      minute_token: carryIdentity ? current.meeting?.minute_token ?? null : null,
      title: '等待飞书会议事件',
      start_time: new Date(startMs).toISOString(),
      end_time: null,
      timezone: current.meeting?.timezone ?? 'Asia/Shanghai',
      pending_binding: true,
      source: 'annotation_fallback',
      reserve_id: carryIdentity ? current.meeting?.reserve_id ?? null : null,
      reserve_meeting_no: carryIdentity ? current.meeting?.reserve_meeting_no ?? null : null,
      app_link: carryIdentity ? current.meeting?.app_link ?? null : null,
      live_link: carryIdentity ? current.meeting?.live_link ?? null : null,
    },
    segments: [],
    events: [{
      id: 'evt-pending-live-start',
      time_ms: 0,
      type: 'meeting_start',
      label: '等待飞书会议事件',
      source: 'local_pending',
      metadata: { raw_type: 'local.pending_annotation_session' },
    }],
    sequence: [],
  });
}

function shouldRebaseAnnotationItem(item = {}) {
  return Boolean(item.raw && item.payload?.timing);
}

function rebaseAnnotationSequence(sequence = [], meeting = {}, fallbackNow = Date.now()) {
  return sequence.map((item) => {
    if (shouldRebaseAnnotationItem(item)) return normalizeAnnotationEvent(item.raw, meeting, fallbackNow);
    return item;
  });
}

function rebasePendingSequence(sequence = [], meeting = {}) {
  return rebaseAnnotationSequence(sequence, meeting);
}

function canOpenSessionCarryIntoMeeting(current = {}, meeting = {}) {
  if (current.meeting?.source !== 'open_meeting_session') return false;
  if (current.meeting?.end_time) return false;
  if (sameMeeting(current.meeting, meeting)) return true;
  const currentStartMs = parseAbsoluteMs(current.meeting?.start_time);
  const meetingStartMs = parseAbsoluteMs(meeting?.start_time);
  if (currentStartMs == null || meetingStartMs == null) return false;
  const deltaMs = currentStartMs - meetingStartMs;
  return deltaMs >= -30_000 && deltaMs <= 30 * 60_000;
}

function shouldCarrySequenceIntoNewAxis(current = {}, meeting = {}) {
  if (current.meeting?.pending_binding) return true;
  if (current.meeting?.source === 'local_simulation') return true;
  if (current.meeting?.source === 'lark_reserve_pending') return true;
  if (current.meeting?.source === 'open_meeting_session') return canOpenSessionCarryIntoMeeting(current, meeting);
  if (isRealMeetingAxis(current.meeting)) return sameMeeting(current.meeting, meeting);
  return false;
}

function meetingStartTimeFromSessionBody(body = {}, nowMs = Date.now()) {
  const explicit = firstNonEmpty(
    body.start_time_ms,
    body.started_at_ms,
    body.meeting_start_time_ms,
    body.start_time,
    body.started_at,
    body.meeting_start_time,
  );
  const parsed = parseAbsoluteMs(explicit);
  if (parsed != null) {
    return {
      ms: parsed,
      source: 'operator_supplied_start_time',
      reliable: true,
    };
  }
  return {
    ms: nowMs,
    source: 'session_opened_at',
    reliable: true,
  };
}

function openMeetingSessionMeetingFromBody(body = {}, current = {}, nowMs = Date.now()) {
  const start = meetingStartTimeFromSessionBody(body, nowMs);
  return {
    platform: body.platform ?? 'lark',
    meeting_id: String(firstNonEmpty(
      body.meeting_id,
      body.id,
      body.session_id,
      body.external_meeting_id,
      `open-session-${start.ms}`,
    )),
    external_meeting_id: firstNonEmpty(
      body.external_meeting_id,
      body.meeting_no,
      body.lark_meeting_id,
      body.open_meeting_id,
      null,
    ),
    meeting_url: firstNonEmpty(body.meeting_url, body.url, body.join_url, current.meeting?.meeting_url, null),
    minute_token: firstNonEmpty(body.minute_token, current.meeting?.minute_token, null),
    title: String(firstNonEmpty(body.title, body.topic, body.name, '开放会议会话')),
    start_time: new Date(start.ms).toISOString(),
    start_time_source: start.source,
    start_time_reliable: start.reliable,
    end_time: body.end_time ?? null,
    timezone: body.timezone ?? current.meeting?.timezone ?? 'Asia/Shanghai',
    pending_binding: false,
    source: 'open_meeting_session',
  };
}

function annotationMeetingSessionBody(input = {}) {
  const explicit = input.meeting_session
    ?? input.meetingSession
    ?? input.payload?.meeting_session
    ?? input.payload?.meetingSession
    ?? input.context?.meeting_session
    ?? input.context?.meetingSession
    ?? null;
  if (explicit && typeof explicit === 'object' && !Array.isArray(explicit)) {
    return {
      ...explicit,
      detector_source: explicit.detector_source ?? explicit.source ?? input.source ?? input.device?.type ?? null,
      force: explicit.force ?? input.force_meeting_session ?? input.force_open_meeting_session ?? false,
    };
  }
  if (input.start_meeting_session === true || input.open_meeting_session === true) {
    return {
      platform: input.meeting_platform ?? input.platform ?? 'lark',
      meeting_id: input.meeting_id ?? input.session_id ?? null,
      external_meeting_id: input.external_meeting_id ?? input.meeting_no ?? null,
      meeting_url: input.meeting_url ?? input.url ?? input.join_url ?? null,
      title: input.meeting_title ?? input.title ?? input.topic ?? null,
      start_time_ms: input.meeting_start_time_ms ?? input.start_time_ms ?? input.captured_at_ms ?? null,
      start_time: input.meeting_start_time ?? input.start_time ?? null,
      detector_source: input.detector_source ?? input.source ?? input.device?.type ?? null,
      force: input.force_meeting_session ?? input.force_open_meeting_session ?? false,
    };
  }
  return null;
}

function hasMeetingSessionIdentity(body = {}) {
  return Boolean(
    firstNonEmpty(
      body.meeting_id,
      body.id,
      body.session_id,
      body.external_meeting_id,
      body.meeting_no,
      body.lark_meeting_id,
      body.open_meeting_id,
      body.meeting_url,
      body.url,
      body.join_url,
      null,
    ),
  );
}

async function maybeStartOpenMeetingSessionFromAnnotation(input = {}) {
  const body = annotationMeetingSessionBody(input);
  if (!body) return { attempted: false, reason: 'missing_meeting_session_context' };
  const current = await store.load();
  const nowMs = Date.now();
  const meeting = openMeetingSessionMeetingFromBody(body, current, nowMs);
  const activeRealAxis = isRealMeetingAxis(current.meeting) && !current.meeting?.end_time;
  if (activeRealAxis && !body.force) {
    if (!hasMeetingSessionIdentity(body) || sameMeeting(current.meeting, meeting)) {
      return {
        attempted: true,
        status: 'already_bound',
        reason: 'current_real_axis_matches_or_no_new_identity',
        meeting_id: current.meeting?.meeting_id ?? null,
        source: current.meeting?.source ?? null,
      };
    }
    return {
      attempted: true,
      status: 'kept_current_axis',
      reason: 'different_active_real_axis_requires_force',
      meeting_id: current.meeting?.meeting_id ?? null,
      requested_meeting_id: meeting.meeting_id,
      source: current.meeting?.source ?? null,
    };
  }
  const result = await startOpenMeetingSession(body);
  return {
    attempted: true,
    status: 'started',
    reason: null,
    meeting_id: result.meeting?.meeting_id ?? null,
    source: result.meeting?.source ?? null,
    state: result.state,
  };
}

function carriedSequenceForNewRealAxis(current = {}, meeting = {}) {
  if (shouldCarrySequenceIntoNewAxis(current, meeting)) {
    return rebaseAnnotationSequence(current.sequence ?? [], meeting);
  }
  return [];
}

async function startOpenMeetingSession(body = {}) {
  const current = await store.load();
  const nowMs = Date.now();
  const meeting = openMeetingSessionMeetingFromBody(body, current, nowMs);
  const currentIsActiveReal = hasActiveMeeting(current) && isRealMeetingAxis(current.meeting);
  if (currentIsActiveReal && !sameMeeting(current.meeting, meeting) && !body.force) {
    const error = new Error('Current real meeting axis is active; pass force=true to replace it with an open meeting session.');
    error.status = 409;
    error.current_meeting = current.meeting;
    throw error;
  }
  const next = buildTimeline({
    meeting,
    segments: body.keep_segments === true ? current.segments ?? [] : [],
    events: [{
      id: `evt-open-session-${meeting.meeting_id}-start`,
      time_ms: 0,
      type: 'meeting_start',
      label: '开放会议会话开始',
      source: 'open_meeting_session',
      metadata: {
        raw_type: 'local.open_meeting_session.started',
        detector_source: body.detector_source ?? body.source ?? null,
        note: body.note ?? null,
      },
    }],
    sequence: carriedSequenceForNewRealAxis(current, meeting),
  });
  const saved = await saveAndBroadcast(next, 'state');
  if (realDemoSession.active && isRealMeetingAxis(saved.meeting)) {
    saveRealDemoSessionState({
      ...realDemoSession,
      last_real_axis_at: new Date().toISOString(),
      last_real_axis_source: 'open_meeting_session',
    });
  }
  const shouldSuppressDemoAnnotations = Boolean(body.suppress_auto_annotations || body.suppress_demo_annotations);
  const autoAnnotation = shouldSuppressDemoAnnotations
    ? { state: saved, annotation: null, skipped_reason: 'suppressed' }
    : await maybeAppendAutoAcceptanceAnnotation(saved, 'open_meeting_session');
  const deviceAnnotation = shouldSuppressDemoAnnotations
    ? { state: autoAnnotation.state, annotation: null, skipped_reason: 'suppressed' }
    : await maybeAppendDeviceSimulatorAnnotation(autoAnnotation.state, 'open_meeting_session');
  return {
    ok: true,
    meeting: deviceAnnotation.state.meeting,
    state: deviceAnnotation.state,
    auto_acceptance_annotation: autoAnnotation.annotation,
    device_simulator_annotation: deviceAnnotation.annotation,
    contract: {
      source: 'open_meeting_session',
      strict_lark_event_axis: false,
      product_axis: true,
      note: 'This endpoint is the stable local meeting-session contract. Lark events, desktop observers, or e-ink host apps can all call it when a meeting actually starts.',
    },
  };
}

async function endOpenMeetingSession(body = {}) {
  const current = await store.load();
  if (!isRealMeetingAxis(current.meeting)) {
    const error = new Error('No active real meeting axis to end.');
    error.status = 409;
    error.current_meeting = current.meeting;
    throw error;
  }
  const now = new Date();
  const explicitEndMs = parseAbsoluteMs(firstNonEmpty(body.end_time_ms, body.ended_at_ms, body.end_time, body.ended_at));
  const startMs = Date.parse(current.meeting?.start_time ?? '');
  const offsetMs = explicitEndMs != null && Number.isFinite(startMs)
    ? Math.max(0, explicitEndMs - startMs)
    : body.time_ms != null
      ? Math.max(0, Math.round(Number(body.time_ms)))
      : Number.isFinite(startMs)
        ? Math.max(0, now.getTime() - startMs)
        : 0;
  const eventMap = new Map((current.events ?? []).map((item) => [item.id, item]));
  eventMap.set('evt-open-session-end', {
    id: 'evt-open-session-end',
    time_ms: offsetMs,
    type: 'meeting_end',
    label: '开放会议会话结束',
    source: 'open_meeting_session',
    metadata: {
      raw_type: 'local.open_meeting_session.ended',
      detector_source: body.detector_source ?? body.source ?? null,
    },
  });
  const next = mergeTimelineWithRebasedAnnotations(current, {
    meeting: {
      ...current.meeting,
      end_time: explicitEndMs != null ? new Date(explicitEndMs).toISOString() : isoAtMeetingOffset(current.meeting, offsetMs, now),
      end_time_source: explicitEndMs != null ? 'operator_supplied_end_time' : 'session_closed_at',
    },
    events: [...eventMap.values()],
  });
  return {
    ok: true,
    state: await saveAndBroadcast(next, 'state'),
  };
}

function mergeTimelineWithRebasedAnnotations(current = {}, patch = {}, fallbackNow = Date.now()) {
  const meeting = { ...(current.meeting ?? {}), ...(patch.meeting ?? {}) };
  return mergeTimeline(current, {
    ...patch,
    meeting,
    sequence: rebaseAnnotationSequence(patch.sequence ?? current.sequence ?? [], meeting, fallbackNow),
  });
}

function sameMeeting(left = {}, right = {}) {
  const leftIds = new Set([
    left.meeting_id,
    left.external_meeting_id,
    left.minute_token,
    left.meeting_url,
  ].filter(Boolean).map(String));
  return [
    right.meeting_id,
    right.external_meeting_id,
    right.minute_token,
    right.meeting_url,
  ].filter(Boolean).some((id) => leftIds.has(String(id)));
}

function explicitMeetingIdentity(payload = {}) {
  const event = payload?.event ?? payload?.data?.event ?? payload;
  return {
    meeting_id: event?.meeting?.meeting_id
      ?? event?.meeting?.id
      ?? event?.meeting_info?.meeting_id
      ?? event?.meeting_info?.id
      ?? event?.meeting_info?.open_meeting_id
      ?? event?.meeting_info?.meeting_no
      ?? event?.meeting_id
      ?? event?.meeting_no
      ?? event?.vc_meeting_id
      ?? event?.id
      ?? payload?.meeting_id
      ?? null,
    external_meeting_id: event?.open_meeting_id
      ?? event?.meeting?.open_meeting_id
      ?? event?.meeting?.meeting_no
      ?? event?.meeting_info?.open_meeting_id
      ?? event?.meeting_info?.meeting_no
      ?? event?.meeting_no
      ?? event?.vc_meeting_id
      ?? null,
    minute_token: event?.minute?.token
      ?? event?.minute?.minute_token
      ?? event?.minute_token
      ?? event?.minutes_token
      ?? payload?.minute_token
      ?? null,
    meeting_url: event?.meeting?.meeting_url
      ?? event?.meeting?.url
      ?? event?.meeting?.join_url
      ?? event?.meeting?.share_url
      ?? event?.meeting_info?.meeting_url
      ?? event?.meeting_info?.url
      ?? event?.meeting_info?.join_url
      ?? event?.meeting_info?.share_url
      ?? event?.meeting_url
      ?? event?.url
      ?? event?.join_url
      ?? event?.share_url
      ?? payload?.meeting_url
      ?? null,
  };
}

function firstPayloadPath(payload = {}, paths = []) {
  const event = payload?.event ?? payload?.data?.event ?? payload;
  const header = payload?.header ?? {};
  const roots = { payload, event, header };
  for (const path of paths) {
    const parts = path.split('.');
    let node = roots;
    for (const part of parts) node = node?.[part];
    if (node != null && node !== '') return node;
  }
  return null;
}

function explicitMeetingStartMs(payload = {}) {
  return parseAbsoluteMs(firstPayloadPath(payload, [
    'event.meeting.start_time',
    'event.meeting.start_at',
    'event.meeting.begin_time',
    'event.meeting_info.start_time',
    'event.meeting_info.start_at',
    'event.meeting_info.begin_time',
    'event.start_time',
    'event.start_at',
    'event.begin_time',
    'payload.start_time',
  ]));
}

function inferredStartEventForEnd(payload = {}, source = 'lark_event') {
  const event = payload?.event ?? payload?.data?.event ?? payload;
  const header = payload?.header ?? {};
  return {
    id: String(header.event_id ? `${header.event_id}-inferred-start` : `evt-inferred-start-${Date.now()}`),
    time_ms: 0,
    type: 'meeting_start',
    label: '会议开始',
    source,
    metadata: {
      raw_type: 'inferred_from_meeting_end_start_time',
      meeting_id: explicitMeetingIdentity(payload).meeting_id,
      source_end_event_id: header.event_id ?? event?.id ?? null,
    },
  };
}

function hasExplicitMeetingIdentity(identity = {}) {
  return Boolean(
    identity.meeting_id
      || identity.external_meeting_id
      || identity.minute_token
      || identity.meeting_url,
  );
}

function sameExplicitMeeting(left = {}, identity = {}) {
  const leftIds = new Set([
    left.meeting_id,
    left.external_meeting_id,
    left.minute_token,
    left.meeting_url,
  ].filter(Boolean).map(String));
  return [
    identity.meeting_id,
    identity.external_meeting_id,
    identity.minute_token,
    identity.meeting_url,
  ].filter(Boolean).some((id) => leftIds.has(String(id)));
}

function clearInheritedFieldsForNewMeeting(meetingPatch, identity) {
  if (!identity.external_meeting_id) meetingPatch.external_meeting_id = null;
  if (!identity.minute_token) meetingPatch.minute_token = null;
  if (!identity.meeting_url) meetingPatch.meeting_url = null;
  return meetingPatch;
}

function secondsOrMsToIso(value, fallback = null) {
  const ms = parseAbsoluteMs(value);
  if (ms == null) return fallback;
  return new Date(ms).toISOString();
}

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '') ?? null;
}

function meetingRecordTimeRaw(record = {}, kind = 'start') {
  const meeting = record.meeting ?? record;
  return firstNonEmpty(
    kind === 'end' ? meeting.end_time : meeting.start_time,
    kind === 'end' ? meeting.end_at : meeting.start_at,
    kind === 'end' ? record.end_time : record.start_time,
    kind === 'end' ? record.end_at : record.start_at,
    kind === 'start' ? meeting.begin_time : null,
    kind === 'start' ? record.begin_time : null,
  );
}

function meetingStartOverride(body = {}) {
  const raw = firstNonEmpty(
    body.meeting_start_time,
    body.meeting_started_at,
    body.axis_start_time,
  );
  return parseAbsoluteMs(raw) == null ? null : raw;
}

function requireReliableMeetingStart(record = {}, body = {}, source = 'lark_meeting_search_api') {
  if (meetingRecordTimeMs(record, 'start') != null) return;
  if (meetingStartOverride(body) != null) return;
  const error = new Error('Feishu meeting search result does not include start_time; refusing to build a real axis from request/click time.');
  error.status = 422;
  error.code = 'missing_meeting_start_time';
  error.source = source;
  error.next_action = 'Wait for a real meeting_start event, enable an API that returns meeting.start_time, or pass meeting_start_time explicitly for manual recovery.';
  throw error;
}

function meetingFromLarkRecord(record = {}, fallback = {}) {
  const meeting = record.meeting ?? record;
  const startRaw = meetingRecordTimeRaw(record, 'start');
  const endRaw = meetingRecordTimeRaw(record, 'end');
  const fallbackStartMs = parseAbsoluteMs(fallback.start_time);
  const fallbackEndMs = parseAbsoluteMs(fallback.end_time);
  const startTime = secondsOrMsToIso(
    startRaw,
    fallbackStartMs != null ? new Date(fallbackStartMs).toISOString() : new Date().toISOString(),
  );
  const endTime = secondsOrMsToIso(
    endRaw,
    fallbackEndMs != null ? new Date(fallbackEndMs).toISOString() : null,
  );
  const startTimeSource = startRaw != null
    ? 'lark_record_start_time'
    : fallbackStartMs != null
      ? fallback.start_time_source ?? 'manual_start_override'
      : 'server_now_fallback';
  const endTimeSource = endRaw != null
    ? 'lark_record_end_time'
    : fallbackEndMs != null
      ? fallback.end_time_source ?? 'manual_end_override'
      : null;
  return {
    platform: 'lark',
    meeting_id: String(firstNonEmpty(
      meeting.id,
      meeting.meeting_id,
      record.id,
      record.meeting_id,
      fallback.meeting_id,
      `lark-meeting-${Date.now()}`,
    )),
    external_meeting_id: firstNonEmpty(
      meeting.meeting_no,
      meeting.open_meeting_id,
      record.meeting_no,
      record.open_meeting_id,
      fallback.external_meeting_id,
    ),
    meeting_url: firstNonEmpty(
      meeting.url,
      meeting.meeting_url,
      meeting.join_url,
      meeting.share_url,
      record.url,
      record.meeting_url,
      record.join_url,
      fallback.meeting_url,
    ),
    minute_token: firstNonEmpty(
      meeting.minute_token,
      meeting.minutes_token,
      meeting.minute?.token,
      record.minute_token,
      record.minutes_token,
      fallback.minute_token,
    ),
    title: String(firstNonEmpty(
      meeting.topic,
      meeting.title,
      meeting.name,
      record.topic,
      record.title,
      fallback.title,
      '飞书会议',
    )),
    start_time: startTime,
    start_time_source: startTimeSource,
    start_time_reliable: startTimeSource !== 'server_now_fallback',
    end_time: endTime,
    end_time_source: endTimeSource,
    timezone: fallback.timezone ?? 'Asia/Shanghai',
    source: fallback.source ?? 'lark_meeting_lookup_api',
  };
}

function candidateMeetingsFromListResponse(raw = {}) {
  const data = raw.data ?? raw;
  const candidates = [
    data.meetings,
    data.meeting_list,
    data.items,
    data.list,
    data.meeting ? [data.meeting] : null,
  ].find(Array.isArray);
  return candidates ?? [];
}

function candidateMeetingsFromSearchResponse(raw = {}) {
  const data = raw.data ?? raw;
  return [
    data.items,
    data.meetings,
    data.meeting_list,
    data.list,
  ].find(Array.isArray) ?? [];
}

function meetingRecordTimeMs(record = {}, kind = 'start') {
  const meeting = record.meeting ?? record;
  return parseAbsoluteMs(firstNonEmpty(
    kind === 'end' ? meeting.end_time : meeting.start_time,
    kind === 'end' ? meeting.end_at : meeting.start_at,
    kind === 'end' ? record.end_time : record.start_time,
    kind === 'end' ? record.end_at : record.start_at,
    kind === 'start' ? meeting.begin_time : null,
    kind === 'start' ? record.begin_time : null,
  ));
}

function chooseCurrentMeetingCandidate(items = [], requestedId = null, nowMs = Date.now()) {
  if (!items.length) return null;
  if (requestedId) {
    return items.find((item) => String(item.id ?? item.meeting_id ?? item.meeting?.id) === String(requestedId)) ?? items[0];
  }
  return [...items].sort((left, right) => {
    const score = (item) => {
      const startMs = meetingRecordTimeMs(item, 'start');
      const endMs = meetingRecordTimeMs(item, 'end');
      if (startMs == null) return Number.MAX_SAFE_INTEGER;
      const currentishEnd = endMs ?? startMs + 3 * 60 * 60 * 1000;
      if (startMs <= nowMs && currentishEnd >= nowMs) return Math.abs(nowMs - startMs);
      if (startMs <= nowMs) return 10_000_000_000 + Math.abs(nowMs - startMs);
      return 20_000_000_000 + Math.abs(startMs - nowMs);
    };
    return score(left) - score(right);
  })[0];
}

function currentUserIdCandidates() {
  const user = publicAuthState().user ?? {};
  return [
    user.user_id,
    user.open_id,
    user.union_id,
  ].filter(Boolean).map(String);
}

function hasOAuthScope(scope) {
  return tokenScopeList().includes(scope);
}

function meetingSearchErrorPayload(error) {
  const message = error.message ?? String(error);
  const invalidToken = /invalid access token|token attached|invalid token/i.test(message);
  const permissionDenied = !invalidToken && /permission|access denied|scope|auth/i.test(message);
  capabilityStatus.meeting_search.status = invalidToken ? 'invalid_token' : permissionDenied ? 'missing_scope' : 'error';
  capabilityStatus.meeting_search.checked_at = new Date().toISOString();
  capabilityStatus.meeting_search.error = message;
  capabilityStatus.meeting_search.permission_url = permissionUrlForScopes('vc:meeting.search:read');
  return {
    status: error.status ?? (invalidToken ? 401 : permissionDenied ? 403 : 500),
    payload: {
      error: message,
      required_scope: 'vc:meeting.search:read',
      permission_url: capabilityStatus.meeting_search.permission_url,
      search_status: capabilityStatus.meeting_search.status,
      next_action: invalidToken
        ? '当前会议搜索接口拒绝应用/租户 access token；请使用当前用户 OAuth 重新授权 vc:meeting.search:read 后走“扫描我的真实会议”。'
        : permissionDenied
          ? '请在飞书开放平台开通并发布 vc:meeting.search:read，然后重新授权当前用户。'
          : '请查看飞书 API 返回错误并重试。',
    },
  };
}

async function searchCurrentUserMeetings(body = {}) {
  const userToken = await ensureUserAccessToken();
  if (!userToken) {
    const error = new Error('请先登录或重新登录飞书账号。');
    error.status = 401;
    throw error;
  }
  const idCandidates = currentUserIdCandidates();
  if (!idCandidates.length) {
    const error = new Error('当前 OAuth 用户缺少 user_id/open_id/union_id，无法按当前用户扫描会议。');
    error.status = 400;
    throw error;
  }
  if (!hasOAuthScope('vc:meeting.search:read')) {
    const error = new Error('当前飞书 OAuth token 缺少 vc:meeting.search:read；请更新 LARK_OAUTH_SCOPES 后重新登录。');
    error.status = 403;
    throw error;
  }

  const nowSeconds = Math.floor(Date.now() / 1000);
  const baseOpts = {
    query: body.query ?? body.title ?? '',
    start_time: body.start_time ?? nowSeconds - Number(body.lookback_seconds ?? 2 * 60 * 60),
    end_time: body.end_time ?? nowSeconds + Number(body.lookahead_seconds ?? 30 * 60),
    page_size: body.page_size ?? 10,
  };
  const attempts = [];
  const seen = new Set();
  const items = [];
  for (const id of idCandidates) {
    for (const filterKind of ['participant_ids', 'organizer_ids']) {
      const opts = { ...baseOpts, [filterKind]: [id] };
      try {
        const raw = await lark.searchMeetingsWithToken(userToken, opts);
        const rawItems = candidateMeetingsFromSearchResponse(raw);
        attempts.push({ filter_kind: filterKind, id, ok: true, item_count: rawItems.length, raw });
        for (const item of rawItems) {
          const itemId = String(item.id ?? item.meeting_id ?? item.meeting?.id ?? item.display_info ?? JSON.stringify(item));
          if (seen.has(itemId)) continue;
          seen.add(itemId);
          items.push(item);
        }
      } catch (error) {
        attempts.push({ filter_kind: filterKind, id, ok: false, error: error.message ?? String(error) });
      }
    }
  }
  const anyOk = attempts.some((attempt) => attempt.ok);
  if (!anyOk) {
    const firstError = attempts.find((attempt) => !attempt.ok)?.error ?? 'meeting search failed';
    throw new Error(firstError);
  }
  capabilityStatus.meeting_search.status = 'ok';
  capabilityStatus.meeting_search.checked_at = new Date().toISOString();
  capabilityStatus.meeting_search.error = null;
  capabilityStatus.meeting_search.permission_url = permissionUrlForScopes('vc:meeting.search:read');
  return {
    auth_mode: 'user_oauth',
    user_ids_tried: idCandidates,
    item_count: items.length,
    items,
    attempts,
  };
}

async function searchTenantMeetings(body = {}) {
  const nowSeconds = Math.floor(Date.now() / 1000);
  const opts = {
    query: body.query ?? body.title ?? '',
    start_time: body.start_time ?? nowSeconds - Number(body.lookback_seconds ?? 15 * 60),
    end_time: body.end_time ?? nowSeconds + Number(body.lookahead_seconds ?? 30 * 60),
    page_size: body.page_size ?? 10,
    page_token: body.page_token,
  };
  const raw = await lark.searchMeetings(opts);
  const items = candidateMeetingsFromSearchResponse(raw);
  capabilityStatus.meeting_search.status = 'ok';
  capabilityStatus.meeting_search.checked_at = new Date().toISOString();
  capabilityStatus.meeting_search.error = null;
  capabilityStatus.meeting_search.permission_url = permissionUrlForScopes('vc:meeting.search:read');
  return {
    auth_mode: 'tenant_access_token',
    item_count: items.length,
    items,
    raw,
  };
}

async function bindMeetingFromSearch(search, body = {}, source = 'lark_meeting_search_api') {
  if (!search.items.length) {
    const error = new Error('No meeting found in the requested time range');
    error.status = 404;
    throw error;
  }
  const current = await store.load();
  const selected = chooseCurrentMeetingCandidate(search.items, body.meeting_id);
  const selectedId = selected.id ?? selected.meeting_id ?? selected.meeting?.id;
  let detailRaw = null;
  let record = selected;
  const userToken = search.auth_mode === 'user_oauth' ? await ensureUserAccessToken() : null;
  if (selectedId && userToken) {
    try {
      detailRaw = await lark.fetchMeetingDetailWithToken(selectedId, userToken);
      record = detailRaw?.data?.meeting ?? detailRaw?.data ?? detailRaw?.meeting ?? selected;
    } catch {
      record = selected;
    }
  }
  requireReliableMeetingStart(record, body, source);
  const manualStart = meetingStartOverride(body);
  const meeting = meetingFromLarkRecord(record, {
    meeting_id: selectedId ?? undefined,
    title: selected.display_info ?? body.title ?? undefined,
    meeting_url: selected.meta_data?.app_link ?? body.meeting_url ?? body.url ?? undefined,
    end_time: body.end_time ?? null,
    source,
    start_time: manualStart,
    start_time_source: manualStart != null ? 'manual_start_override' : undefined,
    timezone: current.meeting?.timezone ?? 'Asia/Shanghai',
  });
  const carriedSequence = carriedSequenceForNewRealAxis(current, meeting);
  const next = buildTimeline({
    meeting,
    segments: current.segments ?? [],
    events: [{
      id: `evt-search-${meeting.meeting_id}-start`,
      time_ms: 0,
      type: 'meeting_start',
      label: source === 'lark_probe_auto_search'
        ? 'probe 自动扫描建轴'
        : source === 'lark_passive_meeting_scan'
          ? '被动扫描建轴'
          : source === 'lark_tenant_passive_meeting_scan'
            ? '租户被动扫描建轴'
            : source === 'lark_tenant_meeting_search_api'
              ? '应用身份会议搜索建轴'
              : '会议搜索建轴',
      source,
      metadata: { raw_type: 'vc.meeting.search', selected_id: selectedId ?? null, auth_mode: search.auth_mode },
    }],
    sequence: carriedSequence,
  });
  const saved = await saveAndBroadcast(next, 'state');
  if (realDemoSession.active && isRealMeetingAxis(saved.meeting)) {
    saveRealDemoSessionState({
      ...realDemoSession,
      last_real_axis_at: new Date().toISOString(),
      last_real_axis_source: source,
    });
  }
  const autoAnnotation = await maybeAppendAutoAcceptanceAnnotation(saved, source);
  const deviceAnnotation = await maybeAppendDeviceSimulatorAnnotation(autoAnnotation.state, source);
  return {
    search: {
      auth_mode: search.auth_mode,
      item_count: search.item_count,
      user_ids_tried: search.user_ids_tried,
    },
    selected,
    detail_raw: detailRaw,
    auto_acceptance_annotation: autoAnnotation.annotation,
    device_simulator_annotation: deviceAnnotation.annotation,
    state: deviceAnnotation.state,
  };
}

async function bindCurrentUserMeeting(body = {}) {
  return bindMeetingFromSearch(await searchCurrentUserMeetings(body), body, 'lark_meeting_search_api');
}

async function bindTenantMeeting(body = {}) {
  return bindMeetingFromSearch(await searchTenantMeetings(body), body, 'lark_tenant_meeting_search_api');
}

function meetingRecordMatchesProbeWindow(record = {}, startedMs = 0, nowMs = Date.now()) {
  const startMs = meetingRecordTimeMs(record, 'start');
  if (startMs == null) return false;
  const endMs = meetingRecordTimeMs(record, 'end');
  const startedSlackMs = 5 * 60 * 1000;
  const futureSlackMs = 2 * 60 * 1000;
  if (startMs < startedMs - startedSlackMs) return false;
  if (startMs > nowMs + futureSlackMs) return false;
  if (endMs != null && endMs < startedMs - 30_000) return false;
  return true;
}

function meetingRecordMatchesPassiveScanWindow(record = {}, nowMs = Date.now(), options = {}) {
  const startMs = meetingRecordTimeMs(record, 'start');
  if (startMs == null) return false;
  const endMs = meetingRecordTimeMs(record, 'end');
  const lookbackMs = Number(options.lookback_seconds ?? passiveMeetingScan.lookback_seconds ?? 10 * 60) * 1000;
  const lookaheadMs = Number(options.lookahead_seconds ?? passiveMeetingScan.lookahead_seconds ?? 2 * 60) * 1000;
  if (startMs < nowMs - lookbackMs) return false;
  if (startMs > nowMs + lookaheadMs) return false;
  if (endMs != null && endMs < nowMs - 30_000) return false;
  return true;
}

function passiveMeetingScanResultPatch(result = {}) {
  const slim = {
    status: result.status,
    reason: result.reason ?? null,
    at: result.at ?? new Date().toISOString(),
    item_count: result.item_count ?? null,
    candidate_count: result.candidate_count ?? null,
    selected_meeting_id: result.selected_meeting_id ?? result.state?.meeting?.meeting_id ?? null,
    meeting_title: result.state?.meeting?.title ?? null,
    error: result.error ?? null,
    auth_mode: result.auth_mode ?? result.search?.auth_mode ?? null,
    tenant_fallback: result.tenant_fallback ?? null,
    user_oauth_reason: result.user_oauth_reason ?? null,
    required_scope: result.required_scope ?? null,
    permission_url: result.permission_url ?? null,
    next_allowed_at: result.next_allowed_at ?? null,
  };
  return {
    ...passiveMeetingScan,
    last_attempt_at: slim.at,
    last_result: slim,
  };
}

function rememberPassiveMeetingScanResult(result = {}) {
  savePassiveMeetingScanState(passiveMeetingScanResultPatch(result));
  return result;
}

function passiveTenantFallbackCooldownMs() {
  return Math.min(Math.max(Number(
    passiveMeetingScan.tenant_fallback_cooldown_ms ?? defaultPassiveMeetingScan.tenant_fallback_cooldown_ms,
  ), 30_000), 30 * 60_000);
}

function tenantFallbackCooldownResult(context = {}) {
  const last = passiveMeetingScan.last_result ?? {};
  let nextAllowedMs = null;
  if (last.reason === 'tenant_search_cooldown' && last.next_allowed_at) {
    nextAllowedMs = isoMs(last.next_allowed_at);
  } else if (last.reason === 'tenant_search_failed' && last.tenant_fallback === 'failed') {
    const lastMs = isoMs(last.at ?? passiveMeetingScan.last_attempt_at);
    if (lastMs == null) return null;
    nextAllowedMs = lastMs + passiveTenantFallbackCooldownMs();
  } else {
    return null;
  }
  const now = context.now ?? Date.now();
  if (nextAllowedMs == null || now >= nextAllowedMs) return null;
  return {
    status: 'skipped',
    reason: 'tenant_search_cooldown',
    at: context.at ?? new Date(now).toISOString(),
    auth_mode: 'tenant_access_token',
    tenant_fallback: 'cooldown',
    user_oauth_reason: context.user_oauth_reason ?? last.user_oauth_reason ?? null,
    error: last.error ?? 'tenant passive scan is in cooldown after a previous failure',
    required_scope: last.required_scope ?? 'vc:meeting.search:read',
    permission_url: last.permission_url ?? permissionUrlForScopes('vc:meeting.search:read'),
    next_allowed_at: new Date(nextAllowedMs).toISOString(),
  };
}

async function tryBindTenantMeetingForPassiveScan(body = {}, context = {}) {
  const now = context.now ?? Date.now();
  const at = context.at ?? new Date(now).toISOString();
  if (body.tenant_fallback_enabled === false || passiveMeetingScan.tenant_fallback_enabled === false) {
    return rememberPassiveMeetingScanResult({
      status: 'skipped',
      reason: 'tenant_fallback_disabled',
      at,
      auth_mode: 'tenant_access_token',
      tenant_fallback: 'disabled',
      user_oauth_reason: context.user_oauth_reason ?? null,
      required_scope: 'vc:meeting.search:read',
      permission_url: permissionUrlForScopes('vc:meeting.search:read'),
    });
  }
  if (!body.force) {
    const cooldown = tenantFallbackCooldownResult({ ...context, now, at });
    if (cooldown) return rememberPassiveMeetingScanResult(cooldown);
  }
  try {
    const search = await searchTenantMeetings({
      ...body,
      lookback_seconds: body.lookback_seconds ?? passiveMeetingScan.lookback_seconds,
      lookahead_seconds: body.lookahead_seconds ?? passiveMeetingScan.lookahead_seconds,
      page_size: body.page_size ?? 10,
    });
    const candidates = (search.items ?? []).filter((item) => (
      meetingRecordMatchesPassiveScanWindow(item, now, {
        lookback_seconds: body.lookback_seconds,
        lookahead_seconds: body.lookahead_seconds,
      })
    ));
    if (!candidates.length) {
      return rememberPassiveMeetingScanResult({
        status: 'no_match',
        reason: 'no_current_meeting_in_tenant_window',
        at,
        auth_mode: 'tenant_access_token',
        tenant_fallback: 'attempted',
        user_oauth_reason: context.user_oauth_reason ?? null,
        item_count: search.item_count,
        candidate_count: 0,
      });
    }
    const bound = await bindMeetingFromSearch(
      { ...search, items: candidates, item_count: candidates.length },
      body,
      'lark_tenant_passive_meeting_scan',
    );
    return rememberPassiveMeetingScanResult({
      status: 'bound',
      reason: 'tenant_passive_scan_matched_current_meeting',
      at,
      auth_mode: 'tenant_access_token',
      tenant_fallback: 'bound',
      user_oauth_reason: context.user_oauth_reason ?? null,
      item_count: search.item_count,
      candidate_count: candidates.length,
      selected_meeting_id: bound.state.meeting.meeting_id,
      search: bound.search,
      selected: bound.selected,
      state: bound.state,
    });
  } catch (error) {
    const response = meetingSearchErrorPayload(error);
    return rememberPassiveMeetingScanResult({
      status: 'error',
      reason: 'tenant_search_failed',
      at,
      auth_mode: 'tenant_access_token',
      tenant_fallback: 'failed',
      user_oauth_reason: context.user_oauth_reason ?? null,
      error: response.payload?.error ?? error.message ?? String(error),
      required_scope: response.payload?.required_scope,
      permission_url: response.payload?.permission_url,
    });
  }
}

async function autoBindPassiveMeeting(body = {}) {
  const now = Date.now();
  const at = new Date(now).toISOString();
  if (!passiveMeetingScan.enabled && !body.force) {
    return rememberPassiveMeetingScanResult({ status: 'skipped', reason: 'disabled', at });
  }
  const intervalMs = passiveMeetingScanIntervalMs();
  const lastAttemptMs = isoMs(passiveMeetingScan.last_attempt_at);
  if (!body.force && lastAttemptMs != null && now - lastAttemptMs < intervalMs) {
    return {
      status: 'skipped',
      reason: 'throttled',
      at,
      next_allowed_at: new Date(lastAttemptMs + intervalMs).toISOString(),
      last_result: passiveMeetingScan.last_result ?? null,
    };
  }
  const current = await store.load();
  if (isRealMeetingAxis(current.meeting) && !current.meeting?.end_time) {
    return rememberPassiveMeetingScanResult({
      status: 'skipped',
      reason: 'real_axis_already_active',
      at,
      selected_meeting_id: current.meeting.meeting_id,
    });
  }
  if (isRealMeetingProbeWaiting(now) && realMeetingProbe.auto_search?.enabled && !body.force) {
    return rememberPassiveMeetingScanResult({
      status: 'skipped',
      reason: 'probe_auto_search_active',
      at,
    });
  }
  const authAttempt = await ensureMeetingSearchAuth();
  const auth = authAttempt.auth;
  if (!auth.authenticated) {
    const userReason = authAttempt.reason === 'oauth_refresh_failed'
      ? 'oauth_refresh_failed'
      : auth.token_present && auth.expired ? 'oauth_token_expired' : 'oauth_login_required';
    return tryBindTenantMeetingForPassiveScan(body, {
      now,
      at,
      user_oauth_reason: userReason,
      user_oauth_error: authAttempt.error
        ?? (auth.token_present && auth.expired ? '飞书 OAuth token 已过期' : '尚未登录飞书账号'),
    });
  }
  if (!hasOAuthScope('vc:meeting.search:read')) {
    return tryBindTenantMeetingForPassiveScan(body, {
      now,
      at,
      user_oauth_reason: 'missing_oauth_scope',
      user_oauth_error: '当前飞书 OAuth token 缺少 vc:meeting.search:read',
    });
  }
  try {
    const search = await searchCurrentUserMeetings({
      ...body,
      lookback_seconds: body.lookback_seconds ?? passiveMeetingScan.lookback_seconds,
      lookahead_seconds: body.lookahead_seconds ?? passiveMeetingScan.lookahead_seconds,
      page_size: body.page_size ?? 10,
    });
    const candidates = (search.items ?? []).filter((item) => (
      meetingRecordMatchesPassiveScanWindow(item, now, {
        lookback_seconds: body.lookback_seconds,
        lookahead_seconds: body.lookahead_seconds,
      })
    ));
    if (!candidates.length) {
      return rememberPassiveMeetingScanResult({
        status: 'no_match',
        reason: 'no_current_meeting_in_passive_window',
        at,
        item_count: search.item_count,
        candidate_count: 0,
      });
    }
    const bound = await bindMeetingFromSearch(
      { ...search, items: candidates, item_count: candidates.length },
      body,
      'lark_passive_meeting_scan',
    );
    return rememberPassiveMeetingScanResult({
      status: 'bound',
      reason: 'passive_scan_matched_current_user_meeting',
      at,
      item_count: search.item_count,
      candidate_count: candidates.length,
      selected_meeting_id: bound.state.meeting.meeting_id,
      search: bound.search,
      selected: bound.selected,
      state: bound.state,
    });
  } catch (error) {
    const response = meetingSearchErrorPayload(error);
    return rememberPassiveMeetingScanResult({
      status: 'error',
      reason: 'search_failed',
      at,
      error: response.payload?.error ?? error.message ?? String(error),
      required_scope: response.payload?.required_scope,
      permission_url: response.payload?.permission_url,
    });
  }
}

function publicPassiveScanResult(result = null) {
  if (!result) return null;
  return {
    status: result.status ?? null,
    reason: result.reason ?? null,
    at: result.at ?? null,
    selected_meeting_id: result.selected_meeting_id ?? result.state?.meeting?.meeting_id ?? null,
    meeting_title: result.state?.meeting?.title ?? null,
    error: result.error ?? null,
    auth_mode: result.auth_mode ?? result.search?.auth_mode ?? null,
    tenant_fallback: result.tenant_fallback ?? null,
    user_oauth_reason: result.user_oauth_reason ?? null,
    required_scope: result.required_scope ?? null,
    permission_url: result.permission_url ?? null,
    next_allowed_at: result.next_allowed_at ?? null,
  };
}

async function runImmediatePassiveMeetingScan(source = 'manual') {
  if (!passiveMeetingScan.enabled) {
    return { status: 'skipped', reason: 'disabled', at: new Date().toISOString() };
  }
  try {
    return await autoBindPassiveMeeting({ source, force: true });
  } catch (error) {
    return rememberPassiveMeetingScanResult({
      status: 'error',
      reason: `${source}_failed`,
      at: new Date().toISOString(),
      error: error.message ?? String(error),
    });
  }
}

function triggerImmediatePassiveMeetingScan(source = 'manual') {
  runImmediatePassiveMeetingScan(source).catch((error) => {
    rememberPassiveMeetingScanResult({
      status: 'error',
      reason: `${source}_failed`,
      at: new Date().toISOString(),
      error: error.message ?? String(error),
    });
  });
}

function passiveMeetingScanIntervalMs() {
  return Math.min(Math.max(Number(
    passiveMeetingScan.interval_ms ?? 10_000,
  ), 5000), 5 * 60_000);
}

function stopPassiveMeetingScanLoop() {
  if (passiveMeetingScanTimer) {
    clearTimeout(passiveMeetingScanTimer);
    passiveMeetingScanTimer = null;
  }
}

function schedulePassiveMeetingScanLoop(delayMs = passiveMeetingScanIntervalMs()) {
  stopPassiveMeetingScanLoop();
  if (!passiveMeetingScan.enabled) return;
  passiveMeetingScanTimer = setTimeout(runPassiveMeetingScanLoopOnce, Math.max(500, delayMs));
  passiveMeetingScanTimer.unref?.();
}

async function runPassiveMeetingScanLoopOnce() {
  if (passiveMeetingScanInFlight) {
    schedulePassiveMeetingScanLoop();
    return;
  }
  passiveMeetingScanInFlight = true;
  try {
    await autoBindPassiveMeeting({ source: 'server_passive_meeting_scan' });
  } catch (error) {
    rememberPassiveMeetingScanResult({
      status: 'error',
      reason: 'server_passive_meeting_scan_failed',
      at: new Date().toISOString(),
      error: error.message ?? String(error),
    });
  } finally {
    passiveMeetingScanInFlight = false;
    schedulePassiveMeetingScanLoop();
  }
}

function probeAutoSearchResultPatch(result = {}) {
  const slim = {
    status: result.status,
    reason: result.reason ?? null,
    at: result.at ?? new Date().toISOString(),
    item_count: result.item_count ?? null,
    candidate_count: result.candidate_count ?? null,
    selected_meeting_id: result.selected_meeting_id ?? result.state?.meeting?.meeting_id ?? null,
    meeting_title: result.state?.meeting?.title ?? null,
    error: result.error ?? null,
  };
  return {
    ...realMeetingProbe,
    auto_search: {
      ...realMeetingProbe.auto_search,
      last_attempt_at: slim.at,
      last_result: slim,
    },
  };
}

function rememberProbeAutoSearchResult(result = {}) {
  saveRealMeetingProbeState(probeAutoSearchResultPatch(result));
  return result;
}

async function autoBindProbeMeeting(body = {}) {
  const now = Date.now();
  const at = new Date(now).toISOString();
  if (!isRealMeetingProbeWaiting(now) && !body.force) {
    return rememberProbeAutoSearchResult({ status: 'skipped', reason: 'probe_not_waiting', at });
  }
  const current = await store.load();
  if (isRealMeetingAxis(current.meeting)) {
    return rememberProbeAutoSearchResult({
      status: 'skipped',
      reason: 'real_axis_already_active',
      at,
      selected_meeting_id: current.meeting.meeting_id,
    });
  }
  if (!realMeetingProbe.auto_search?.enabled && !body.force) {
    return rememberProbeAutoSearchResult({ status: 'skipped', reason: 'auto_search_disabled', at });
  }
  const intervalMs = Math.min(Math.max(Number(
    body.interval_ms ?? realMeetingProbe.auto_search?.interval_ms ?? 5000,
  ), 1000), 60_000);
  const lastAttemptMs = isoMs(realMeetingProbe.auto_search?.last_attempt_at);
  if (!body.force && lastAttemptMs != null && now - lastAttemptMs < intervalMs) {
    return {
      status: 'skipped',
      reason: 'throttled',
      at,
      next_allowed_at: new Date(lastAttemptMs + intervalMs).toISOString(),
      last_result: realMeetingProbe.auto_search?.last_result ?? null,
    };
  }
  const authAttempt = await ensureMeetingSearchAuth();
  if (!authAttempt.auth?.authenticated) {
    return rememberProbeAutoSearchResult({
      status: 'skipped',
      reason: authAttempt.reason === 'oauth_refresh_failed'
        ? 'oauth_refresh_failed'
        : authAttempt.auth?.token_present && authAttempt.auth?.expired ? 'oauth_token_expired' : 'oauth_login_required',
      at,
      error: authAttempt.error
        ?? (authAttempt.auth?.token_present && authAttempt.auth?.expired ? '飞书 OAuth token 已过期' : '尚未登录飞书账号'),
    });
  }
  if (!hasOAuthScope('vc:meeting.search:read')) {
    return rememberProbeAutoSearchResult({
      status: 'skipped',
      reason: 'missing_oauth_scope',
      at,
      error: '当前飞书 OAuth token 缺少 vc:meeting.search:read',
    });
  }
  const probeStartedMs = isoMs(realMeetingProbe.started_at) ?? now;
  try {
    const search = await searchCurrentUserMeetings({
      ...body,
      start_time: Math.floor((probeStartedMs - 5 * 60 * 1000) / 1000),
      end_time: Math.floor((now + Number(body.lookahead_seconds ?? 30 * 60) * 1000) / 1000),
      page_size: body.page_size ?? 10,
    });
    const candidates = (search.items ?? []).filter((item) => (
      meetingRecordMatchesProbeWindow(item, probeStartedMs, now)
    ));
    if (!candidates.length) {
      return rememberProbeAutoSearchResult({
        status: 'no_match',
        reason: 'no_current_meeting_in_probe_window',
        at,
        item_count: search.item_count,
        candidate_count: 0,
      });
    }
    const bound = await bindMeetingFromSearch(
      { ...search, items: candidates, item_count: candidates.length },
      body,
      'lark_probe_auto_search',
    );
    return rememberProbeAutoSearchResult({
      status: 'bound',
      reason: 'probe_auto_search_matched_real_meeting',
      at,
      item_count: search.item_count,
      candidate_count: candidates.length,
      selected_meeting_id: bound.state.meeting.meeting_id,
      search: bound.search,
      selected: bound.selected,
      state: bound.state,
    });
  } catch (error) {
    const response = meetingSearchErrorPayload(error);
    return rememberProbeAutoSearchResult({
      status: 'error',
      reason: 'search_failed',
      at,
      error: response.payload?.error ?? error.message ?? String(error),
      required_scope: response.payload?.required_scope,
      permission_url: response.payload?.permission_url,
    });
  }
}

function probeAutoSearchIntervalMs() {
  return Math.min(Math.max(Number(
    realMeetingProbe.auto_search?.interval_ms ?? 5000,
  ), 1000), 60_000);
}

function stopProbeAutoSearchLoop() {
  if (probeAutoSearchTimer) {
    clearTimeout(probeAutoSearchTimer);
    probeAutoSearchTimer = null;
  }
}

function shouldRunProbeAutoSearchLoop(now = Date.now()) {
  return Boolean(
    realMeetingProbe.active
      && realMeetingProbe.auto_search?.enabled
      && isRealMeetingProbeWaiting(now),
  );
}

function scheduleProbeAutoSearchLoop(delayMs = probeAutoSearchIntervalMs()) {
  stopProbeAutoSearchLoop();
  if (!shouldRunProbeAutoSearchLoop()) return;
  probeAutoSearchTimer = setTimeout(runProbeAutoSearchLoopOnce, Math.max(100, delayMs));
  probeAutoSearchTimer.unref?.();
}

async function runProbeAutoSearchLoopOnce() {
  if (probeAutoSearchInFlight) {
    scheduleProbeAutoSearchLoop();
    return;
  }
  if (!shouldRunProbeAutoSearchLoop()) {
    stopProbeAutoSearchLoop();
    return;
  }
  probeAutoSearchInFlight = true;
  try {
    const result = await autoBindProbeMeeting({ source: 'server_probe_auto_search' });
    const terminal = result.status === 'bound'
      || result.reason === 'real_axis_already_active'
      || result.reason === 'probe_not_waiting';
    if (terminal) stopProbeAutoSearchLoop();
    else scheduleProbeAutoSearchLoop();
  } catch (error) {
    rememberProbeAutoSearchResult({
      status: 'error',
      reason: 'server_probe_auto_search_failed',
      at: new Date().toISOString(),
      error: error.message ?? String(error),
    });
    scheduleProbeAutoSearchLoop();
  } finally {
    probeAutoSearchInFlight = false;
  }
}

async function searchMeetingsByNoInput(body = {}) {
  const meetingNo = extractMeetingNo(body.meeting_no ?? body.meeting_url ?? body.url ?? body.text);
  if (!meetingNo) {
    const error = new Error('meeting_no or meeting_url is required');
    error.status = 400;
    throw error;
  }
  const opts = {
    start_time: body.start_time ?? Math.floor((Date.now() - 24 * 60 * 60 * 1000) / 1000),
    end_time: body.end_time ?? Math.floor((Date.now() + 24 * 60 * 60 * 1000) / 1000),
    page_size: body.page_size ?? 10,
    page_token: body.page_token,
  };
  let raw;
  let authMode = 'tenant_access_token';
  const userToken = await ensureUserAccessToken();
  try {
    raw = userToken
      ? await lark.listMeetingsByNoWithToken(meetingNo, userToken, opts)
      : await lark.listMeetingsByNo(meetingNo, opts);
    authMode = userToken ? 'user_oauth' : 'tenant_access_token';
  } catch (userError) {
    if (!userToken) throw userError;
    try {
      raw = await lark.listMeetingsByNo(meetingNo, opts);
      authMode = 'tenant_access_token';
    } catch (tenantError) {
      throw new Error(`${userError.message ?? String(userError)}; tenant fallback failed: ${tenantError.message ?? String(tenantError)}`);
    }
  }
  const items = candidateMeetingsFromListResponse(raw);
  capabilityStatus.meeting_lookup.status = 'ok';
  capabilityStatus.meeting_lookup.checked_at = new Date().toISOString();
  capabilityStatus.meeting_lookup.error = null;
  capabilityStatus.meeting_lookup.permission_url = permissionUrlForScopes(['vc:meeting.search:read', 'vc:meeting.meetingid:read']);
  return {
    raw,
    meeting_no: meetingNo,
    auth_mode: authMode,
    item_count: items.length,
    items,
  };
}

function meetingLookupErrorPayload(error, meetingNo = null) {
  const message = error.message ?? String(error);
  const permissionDenied = /permission|access denied|scope|auth/i.test(message);
  const permissionUrl = message.match(/https:\/\/open\.feishu\.cn\/app\/[^\s，]+/)?.[0] ?? null;
  capabilityStatus.meeting_lookup.status = permissionDenied ? 'missing_scope' : 'error';
  capabilityStatus.meeting_lookup.checked_at = new Date().toISOString();
  capabilityStatus.meeting_lookup.error = message;
  capabilityStatus.meeting_lookup.permission_url = permissionUrl ?? permissionUrlForScopes(['vc:meeting.search:read', 'vc:meeting.meetingid:read']);
  return {
    status: error.status ?? (permissionDenied ? 403 : 500),
    payload: {
      error: message,
      meeting_no: meetingNo,
      required_scope: 'vc:meeting.search:read or vc:meeting.meetingid:read',
      permission_url: capabilityStatus.meeting_lookup.permission_url,
    },
  };
}

function isLocalSimulationMeeting(meeting = {}) {
  return meeting.source === 'local_simulation' || meeting.source === 'annotation_fallback';
}

function isRealMeetingAxis(meeting = {}) {
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
  ].includes(meeting.source)
    && !meeting.pending_binding
    && meeting.start_time_reliable !== false
    && meeting.start_time_source !== 'server_now_fallback';
}

function isTemporaryMeetingAxis(meeting = {}) {
  return Boolean(
    meeting.pending_binding
      || meeting.meeting_id === 'demo-lark-meeting-001'
      || [
        'local_simulation',
        'annotation_fallback',
        'lark_reserve_pending',
      ].includes(meeting.source),
  );
}

async function resetTemporaryAxisForProbeIfRequested(body = {}) {
  if (!body.reset_temporary_axis) return { reset: false, reason: 'not_requested' };
  const current = await store.load();
  if (isRealMeetingAxis(current.meeting)) {
    return { reset: false, reason: 'real_meeting_axis_active' };
  }
  if (!isTemporaryMeetingAxis(current.meeting)) {
    return { reset: false, reason: 'current_axis_not_temporary' };
  }
  const state = await store.resetDemo();
  broadcastState(state, 'state');
  return { reset: true, reason: 'temporary_axis_reset' };
}

function sourceForLoggedEvent(entry = {}) {
  if (entry.transport === 'ws_long_connection') return 'lark_ws_event';
  if (entry.transport === 'http_webhook') return 'lark_http_event';
  if (entry.transport === 'http_local_event') return 'lark_http_local_event';
  return 'lark_event_log_restore';
}

async function restoreLatestRealMeetingAxisFromEventLog(body = {}) {
  const startEntry = body.event_id
    ? larkEventLog.find((entry) => entry.id === body.event_id || entry.preview?.event_id === body.event_id)
    : latestLoggedMeetingStartEvent();
  if (!startEntry || !isLoggedMeetingStartEntry(startEntry)) {
    const error = new Error('No logged real meeting start event is available to restore');
    error.status = 404;
    throw error;
  }
  const startPayload = loggedEventPayload(startEntry);
  if (!startPayload) {
    const error = new Error('Logged meeting start event has no payload preview to restore');
    error.status = 409;
    throw error;
  }
  const startResult = await processLarkEventPayload(startPayload, {
    source: sourceForLoggedEvent(startEntry),
    suppress_auto_annotations: true,
  });
  let endEntry = null;
  let endResult = null;
  if (body.apply_end !== false) {
    endEntry = matchingLoggedMeetingEndAfter(startEntry);
    if (endEntry) {
      const endPayload = loggedEventPayload(endEntry);
      if (endPayload) {
        endResult = await processLarkEventPayload(endPayload, {
          source: sourceForLoggedEvent(endEntry),
          suppress_auto_annotations: true,
        });
      }
    }
  }
  return {
    restored: true,
    restored_at: new Date().toISOString(),
    start_event: startEntry,
    end_event: endEntry,
    start_result: {
      ok: startResult.ok,
      timeline_started: startResult.timeline_started,
      ignored_reason: startResult.ignored_reason ?? null,
    },
    end_result: endResult ? {
      ok: endResult.ok,
      timeline_started: endResult.timeline_started,
      ignored_reason: endResult.ignored_reason ?? null,
    } : null,
    state: endResult?.state ?? startResult.state,
  };
}

function openAnnotationItems(state = {}) {
  return (state.sequence ?? []).filter((item) => item.payload?.timing?.server_received_at_ms != null);
}

function annotationObservedMs(item = {}) {
  return item.payload?.timing?.server_received_at_ms
    ?? item.payload?.timing?.captured_at_ms
    ?? null;
}

function meetingEndOffsetMs(meeting = {}) {
  const startMs = parseAbsoluteMs(meeting.start_time);
  const endMs = parseAbsoluteMs(meeting.end_time);
  if (startMs == null || endMs == null) return null;
  return Math.max(0, endMs - startMs);
}

function meetingEndEventOffsetMs(state = {}) {
  const eventOffsets = (state.events ?? [])
    .filter((event) => event?.type === 'meeting_end')
    .map((event) => Number(event.time_ms))
    .filter((value) => Number.isFinite(value) && value >= 0);
  return eventOffsets.length ? Math.min(...eventOffsets) : null;
}

function knownMeetingEndOffsetMs(state = {}) {
  return meetingEndOffsetMs(state.meeting) ?? meetingEndEventOffsetMs(state);
}

function annotationAfterMeetingEndMs(item = {}, state = {}) {
  const stored = Number(item.payload?.timing?.after_meeting_end_ms ?? 0);
  const endOffset = knownMeetingEndOffsetMs(state);
  const computed = endOffset == null ? 0 : Math.max(0, Number(item.time_ms ?? 0) - endOffset);
  return Math.max(Number.isFinite(stored) ? stored : 0, computed);
}

function annotationAxisBindingState(state = {}) {
  if (isRealMeetingAxis(state.meeting)) return 'real_meeting_bound';
  if (state.meeting?.pending_binding) return 'pending_real_meeting';
  if (state.meeting?.source === 'local_simulation') return 'local_simulation';
  if (state.meeting?.meeting_id === 'demo-lark-meeting-001') return 'demo_ignored';
  return 'unbound';
}

function annotationBindingsForState(state = {}) {
  const annotations = openAnnotationItems(state);
  const bindingState = annotationAxisBindingState(state);
  const pending = bindingState === 'pending_real_meeting';
  return annotations.map((item) => {
    const timingReliable = item.time_source !== 'server_received_at';
    const afterMeetingEndMs = annotationAfterMeetingEndMs(item, state);
    const onRealAxis = bindingState === 'real_meeting_bound'
      && timingReliable
      && afterMeetingEndMs <= 0;
    return {
      id: item.id,
      label: item.label,
      kind: item.kind,
      source: item.source,
      time_ms: item.time_ms,
      time_source: item.time_source,
      timing_reliable: timingReliable,
      after_meeting_end_ms: afterMeetingEndMs,
      captured_at_ms: item.payload?.timing?.captured_at_ms ?? null,
      server_received_at_ms: item.payload?.timing?.server_received_at_ms ?? null,
      observed_at_ms: annotationObservedMs(item),
      meeting_id: state.meeting?.meeting_id ?? null,
      meeting_source: state.meeting?.source ?? null,
      pending_binding: Boolean(state.meeting?.pending_binding),
      binding_state: bindingState,
      on_real_axis: onRealAxis,
      pending_real_meeting: pending,
    };
  });
}

function annotationBindingSummary(state = {}, options = {}) {
  const probeStartedMs = parseAbsoluteMs(options.probe_started_at);
  const sinceMs = parseAbsoluteMs(options.since_at ?? options.since_ms);
  const allItems = annotationBindingsForState(state);
  const queryId = options.id == null || options.id === '' ? null : String(options.id);
  const items = queryId == null
    ? allItems
    : allItems.filter((item) => String(item.id) === queryId);
  const sinceProbeItems = probeStartedMs == null
    ? items
    : items.filter((item) => {
      const observedMs = item.observed_at_ms;
      return observedMs != null && observedMs >= probeStartedMs;
    });
  const sinceItems = sinceMs == null
    ? items
    : items.filter((item) => {
      const observedMs = item.observed_at_ms;
      return observedMs != null && observedMs >= sinceMs;
    });
  const countItems = (rows, predicate) => rows.filter(predicate).length;
  return {
    meeting_id: state.meeting?.meeting_id ?? null,
    meeting_source: state.meeting?.source ?? null,
    binding_state: annotationAxisBindingState(state),
    total: items.length,
    real_axis_count: countItems(items, (item) => item.on_real_axis),
    pending_count: countItems(items, (item) => item.pending_real_meeting),
    since_probe_count: sinceProbeItems.length,
    since_probe_real_axis_count: countItems(sinceProbeItems, (item) => item.on_real_axis),
    since_at: options.since_at ?? null,
    since_count: sinceItems.length,
    since_real_axis_count: countItems(sinceItems, (item) => item.on_real_axis),
    last: items[items.length - 1] ?? null,
    query: queryId == null ? null : { id: queryId },
    found: queryId == null ? null : items.length > 0,
    item: queryId == null ? null : items[0] ?? null,
    items,
  };
}

function annotationStatusForState(state = {}, options = {}) {
  const id = options.id == null || options.id === '' ? null : String(options.id);
  const summary = annotationBindingSummary(state, {
    probe_started_at: options.probe_started_at,
    since_at: options.since_at,
    id,
  });
  const item = summary.item;
  const currentMeeting = {
    meeting_id: state.meeting?.meeting_id ?? null,
    title: state.meeting?.title ?? null,
    source: state.meeting?.source ?? null,
    pending_binding: Boolean(state.meeting?.pending_binding),
    start_time: state.meeting?.start_time ?? null,
    end_time: state.meeting?.end_time ?? null,
  };
  if (!item) {
    return {
      query: { id },
      found: false,
      accepted: false,
      annotation_id: id,
      status: 'not_found',
      next_action: 'verify annotation_id, or POST the annotation with a stable id before polling',
      poll_after_ms: null,
      warnings: ['annotation_not_found'],
      current_meeting: currentMeeting,
      binding: null,
      item: null,
      summary,
    };
  }

  const warnings = [];
  if (!item.timing_reliable) warnings.push('missing_captured_at_ms');
  if (Number(item.after_meeting_end_ms ?? 0) > 0) warnings.push('normalized_after_meeting_end');
  if (item.pending_real_meeting) warnings.push('pending_real_meeting');
  if (!item.on_real_axis && !item.pending_real_meeting && item.timing_reliable && Number(item.after_meeting_end_ms ?? 0) <= 0) {
    warnings.push('not_on_real_axis');
  }

  let status = 'stored_not_aligned';
  let nextAction = 'stored, but not accepted as a real-axis aligned mark yet';
  let pollAfterMs = null;
  if (!item.timing_reliable) {
    status = 'needs_device_captured_at';
    nextAction = 'resend the same stable id with captured_at_ms/captured_at or absolute stroke timestamps';
  } else if (Number(item.after_meeting_end_ms ?? 0) > 0) {
    status = 'after_meeting_end';
    nextAction = 'keep the stored content, but do not count it as realtime meeting annotation; verify device clock and captured_at_ms';
  } else if (item.on_real_axis) {
    status = 'real_axis_bound';
    nextAction = 'done';
  } else if (item.pending_real_meeting) {
    status = 'pending_real_meeting';
    nextAction = 'wait for the real Lark meeting axis, then poll again or listen to the SSE stream';
    pollAfterMs = 1000;
  }

  return {
    query: { id },
    found: true,
    accepted: true,
    annotation_id: item.id,
    status,
    next_action: nextAction,
    poll_after_ms: pollAfterMs,
    warnings,
    current_meeting: currentMeeting,
    binding_state: item.binding_state,
    on_real_axis: item.on_real_axis,
    pending_real_meeting: item.pending_real_meeting,
    timing_reliable: item.timing_reliable,
    requires_device_captured_at: !item.timing_reliable,
    after_meeting_end_ms: item.after_meeting_end_ms,
    normalized_time_ms: item.time_ms,
    time_source: item.time_source,
    captured_at_ms: item.captured_at_ms,
    server_received_at_ms: item.server_received_at_ms,
    meeting_id: item.meeting_id,
    meeting_source: item.meeting_source,
    binding: item,
    item,
    summary,
  };
}

function annotationAck(item = {}, state = {}, options = {}) {
  const binding = annotationBindingsForState(state).find((entry) => entry.id === item.id) ?? null;
  const replacedExisting = Boolean(options.replaced_existing);
  const idempotent = Boolean(options.idempotency_key_source);
  const usesServerReceiveTime = item.time_source === 'server_received_at';
  const warnings = [];
  if (!idempotent) warnings.push('missing_stable_id');
  if (usesServerReceiveTime) warnings.push('missing_captured_at_ms');
  if (usesServerReceiveTime && state.meeting?.pending_binding) warnings.push('pending_time_uses_server_receive_time');
  if (usesServerReceiveTime && !state.meeting?.pending_binding && !isRealMeetingAxis(state.meeting)) {
    warnings.push('unbound_time_uses_server_receive_time');
  }
  if (annotationAfterMeetingEndMs(item, state) > 0) warnings.push('normalized_after_meeting_end');
  return {
    accepted: true,
    annotation_id: item.id,
    operation: replacedExisting ? 'updated' : 'created',
    idempotent,
    idempotency_key_source: options.idempotency_key_source ?? null,
    warnings,
    label: item.label,
    kind: item.kind,
    source: item.source,
    normalized_time_ms: item.time_ms,
    time_source: item.time_source,
    timing_reliable: !usesServerReceiveTime,
    requires_device_captured_at: usesServerReceiveTime,
    binding_state: binding?.binding_state ?? annotationAxisBindingState(state),
    pending_binding: binding?.pending_binding ?? Boolean(state.meeting?.pending_binding),
    on_real_axis: Boolean(binding?.on_real_axis),
    pending_real_meeting: Boolean(binding?.pending_real_meeting),
    meeting_id: binding?.meeting_id ?? state.meeting?.meeting_id ?? null,
    meeting_source: binding?.meeting_source ?? state.meeting?.source ?? null,
    meeting_title: state.meeting?.title ?? null,
    captured_at_ms: binding?.captured_at_ms ?? item.payload?.timing?.captured_at_ms ?? null,
    server_received_at_ms: binding?.server_received_at_ms ?? item.payload?.timing?.server_received_at_ms ?? null,
    after_meeting_end_ms: binding?.after_meeting_end_ms ?? annotationAfterMeetingEndMs(item, state),
    created_pending_timeline: Boolean(options.created_pending_timeline),
    replaced_existing: replacedExisting,
    stream_event: 'state',
  };
}

function transcriptStatusForState(state = {}) {
  const segments = state.segments ?? [];
  const meeting = state.meeting ?? {};
  const waitingForRealDemoAxis = Boolean(
    realDemoSession.active
      && !isRealMeetingAxis(meeting)
      && !meeting.pending_binding
      && meeting.meeting_id === 'demo-lark-meeting-001'
      && !meeting.source,
  );
  const effectiveSegments = waitingForRealDemoAxis ? [] : segments;
  const segmentSources = effectiveSegments.reduce((acc, segment) => {
    const source = segment.source ?? 'unknown';
    acc[source] = (acc[source] ?? 0) + 1;
    return acc;
  }, {});
  const hasTranscript = effectiveSegments.length > 0;
  const meetingEnded = Boolean(meeting.end_time);
  const minuteToken = meeting.minute_token ?? null;
  const status = waitingForRealDemoAxis
    ? 'waiting_real_meeting'
    : hasTranscript
    ? 'synced'
    : meetingEnded
      ? minuteToken ? 'ready_to_sync_minute' : 'ended_no_minute_token'
      : 'live_no_transcript_expected';
  const nextAction = waitingForRealDemoAxis
    ? '本次真实会议尚未建轴；样例转写不计入本次验收。实时标注先等待真实会议轴，转写仍在会议结束后处理。'
    : hasTranscript
    ? '转写已导入，标注可对齐到附近发言片段。'
    : meetingEnded
      ? minuteToken
        ? '会议已结束，可调用 POST /api/lark/sync-minute 拉取飞书妙记转写。'
        : '会议已结束但没有 minute_token；可粘贴妙记 token 调用 POST /api/lark/sync-minute，或用 POST /api/import/lark-transcript 手动导入。'
      : '会议中不要求实时转写；标注会先按会议时间轴显示，会议结束后再补转写。';
  return {
    status,
    realtime_blocking: false,
    post_meeting: true,
    has_transcript: hasTranscript,
    segment_count: effectiveSegments.length,
    ignored_sample_segment_count: waitingForRealDemoAxis ? segments.length : 0,
    segment_sources: segmentSources,
    meeting_ended: meetingEnded,
    meeting_end_time: meeting.end_time ?? null,
    minute_token_present: Boolean(minuteToken),
    minute_token: minuteToken,
    sync_endpoint: '/api/lark/sync-minute',
    import_endpoint: '/api/import/lark-transcript',
    next_action: nextAction,
  };
}

function isoAtMeetingOffset(meeting = {}, timeMs = 0, fallback = new Date()) {
  const startMs = Date.parse(meeting.start_time ?? '');
  if (!Number.isFinite(startMs)) return fallback.toISOString();
  return new Date(startMs + Math.max(0, timeMs)).toISOString();
}

function pathValue(root, path) {
  if (!root || !path) return undefined;
  let node = root;
  for (const part of path.split('.')) {
    if (node == null) return undefined;
    node = node[part];
  }
  return node;
}

function eventPayloadFromParsed(parsed = {}) {
  const envelope = parsed?.event ?? parsed?.data ?? parsed;
  return envelope?.data?.event
    ?? envelope?.event
    ?? parsed?.data?.event
    ?? parsed?.data?.payload?.event
    ?? envelope;
}

function eventHeaderFromParsed(parsed = {}) {
  const envelope = parsed?.event ?? parsed?.data ?? parsed;
  return parsed?.header
    ?? parsed?.data?.header
    ?? envelope?.header
    ?? envelope?.data?.header
    ?? parsed?.raw_ws?.header
    ?? parsed?.raw_ws?.data?.header
    ?? {};
}

function eventTypeFromParsed(parsed = {}) {
  for (const symbol of Object.getOwnPropertySymbols(parsed)) {
    if (symbol.description === 'event-type') return parsed[symbol];
  }
  const envelope = parsed?.event ?? parsed?.data ?? parsed;
  const event = eventPayloadFromParsed(parsed);
  const header = eventHeaderFromParsed(parsed);
  return firstNonEmpty(
    header.event_type,
    parsed.header?.event_type,
    parsed.event_type,
    parsed.type,
    parsed.data?.header?.event_type,
    parsed.data?.event_type,
    parsed.data?.type,
    envelope?.header?.event_type,
    envelope?.event_type,
    envelope?.type,
    envelope?.data?.header?.event_type,
    envelope?.data?.event_type,
    envelope?.data?.type,
    event?.event_type,
    event?.type,
    '',
  );
}

function timelineEventCandidate(eventType, parsed = {}) {
  const resolvedEventType = eventType || eventTypeFromParsed(parsed);
  const text = String(resolvedEventType ?? '').toLowerCase();
  if (text.includes('vc.') || text.includes('meeting') || text.includes('minute') || text.includes('record')) return true;
  const event = eventPayloadFromParsed(parsed);
  return Boolean(
    parsed.meeting_id
      || parsed.meeting_no
      || parsed.open_meeting_id
      || parsed.vc_meeting_id
      || parsed.meeting
      || parsed.meeting_info
      || parsed.minute_token
      || parsed.minutes_token
      || parsed.join_url
      || parsed.meeting_url
      || event?.meeting_id
      || event?.meeting_no
      || event?.open_meeting_id
      || event?.vc_meeting_id
      || event?.meeting
      || event?.meeting_info
      || event?.minute_token
      || event?.minutes_token
      || event?.join_url
      || event?.meeting_url
      || pathValue(parsed, 'data.event.meeting')
      || pathValue(parsed, 'data.event.meeting_info')
      || pathValue(parsed, 'event.data.event.meeting')
      || pathValue(parsed, 'event.data.event.meeting_info'),
  );
}

function buildWsPayload(eventType, parsed = {}, raw = {}) {
  const event = eventPayloadFromParsed(parsed);
  const parsedHeader = eventHeaderFromParsed(parsed);
  const rawHeader = eventHeaderFromParsed(raw);
  const resolvedEventType = firstNonEmpty(eventType, eventTypeFromParsed(parsed), eventTypeFromParsed(raw), parsedHeader.event_type, rawHeader.event_type, '');
  return {
    schema: raw.schema ?? '2.0',
    header: {
      event_id: firstNonEmpty(parsed.event_id, parsed.uuid, event?.event_id, event?.id, parsedHeader.event_id, rawHeader.event_id),
      event_type: resolvedEventType,
      create_time: firstNonEmpty(parsed.create_time, parsed.event_ts, event?.create_time, event?.event_ts, event?.time, parsedHeader.create_time, rawHeader.create_time),
    },
    event,
    raw_ws: raw,
  };
}

function syntheticWsParserPayload(eventType, event = {}) {
  const startSeconds = 1_782_442_800;
  return {
    schema: '2.0',
    header: {
      event_id: `self-test-${eventType}`,
      event_type: eventType,
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'ws-parser-self-test-meeting',
        topic: 'WS parser self test',
        start_time: String(startSeconds),
        url: 'https://vc.feishu.cn/j/ws-parser-self-test',
      },
      ...event,
    },
  };
}

function wsParserSelfTestCase(raw) {
  const dispatcher = new Lark.EventDispatcher({
    loggerLevel: Lark.LoggerLevel.warn,
  });
  const parsed = dispatcher.requestHandle?.parse(raw) ?? raw;
  const eventType = eventTypeFromParsed(parsed);
  const candidate = timelineEventCandidate(eventType, parsed);
  const wsPayload = buildWsPayload(eventType, parsed, raw);
  const meetingPatch = extractLarkMeetingPatch(wsPayload, {});
  const normalizedEvent = normalizeLarkEventPayload(wsPayload, meetingPatch);
  const startLike = normalizedEvent.type === 'meeting_start'
    || (normalizedEvent.type === 'participant_join' && isMeetingContextBootstrapEvent(normalizedEvent.metadata?.raw_type));
  return {
    input_event_type: raw.header?.event_type ?? null,
    parsed_event_type: eventType,
    parsed_keys: parsed && typeof parsed === 'object'
      ? Object.keys(parsed).filter((key) => !/token|secret|ticket|authorization/i.test(key)).slice(0, 30)
      : [],
    timeline_candidate: candidate,
    normalized_event: {
      type: normalizedEvent.type,
      label: normalizedEvent.label,
      time_ms: normalizedEvent.time_ms,
      raw_type: normalizedEvent.metadata?.raw_type ?? null,
      meeting_id: normalizedEvent.metadata?.meeting_id ?? null,
    },
    meeting_patch: {
      meeting_id: meetingPatch.meeting_id,
      title: meetingPatch.title,
      start_time: meetingPatch.start_time,
      source: meetingPatch.source ?? null,
    },
    would_start_timeline: Boolean(candidate && startLike && meetingPatch.meeting_id && meetingPatch.start_time),
  };
}

function wsParserSelfTestPayload() {
  const direct = wsParserSelfTestCase(syntheticWsParserPayload(
    meetingEventRequirements.direct_meeting.start_event,
  ));
  const context = wsParserSelfTestCase(syntheticWsParserPayload(
    meetingEventRequirements.meeting_context.start_event,
    {
      participant: {
        open_id: 'ou_self_test',
        name: 'Self Test User',
      },
    },
  ));
  const passed = direct.parsed_event_type === meetingEventRequirements.direct_meeting.start_event
    && direct.would_start_timeline
    && context.parsed_event_type === meetingEventRequirements.meeting_context.start_event
    && context.would_start_timeline;
  return {
    generated_at: new Date().toISOString(),
    passed,
    mutates_state: false,
    note: 'This uses the same SDK RequestHandle.parse plus local event normalization, but does not write timeline state or event logs.',
    cases: {
      direct_meeting_start: direct,
      meeting_context_join: context,
    },
  };
}

async function processLarkEventPayload(payload, opts = {}) {
  const current = await store.load();
  const meetingPatch = extractLarkMeetingPatch(payload, current.meeting);
  const event = normalizeLarkEventPayload(payload, { ...current.meeting, ...meetingPatch });
  const source = opts.source ?? 'lark_event';
  event.source = source;
  const payloadStartMs = explicitMeetingStartMs(payload);
  const incomingIdentity = explicitMeetingIdentity(payload);
  const hasIncomingIdentity = hasExplicitMeetingIdentity(incomingIdentity);
  const incomingMatchesCurrent = hasIncomingIdentity
    ? sameExplicitMeeting(current.meeting, incomingIdentity)
    : sameMeeting(current.meeting, meetingPatch);
  const rawEventType = event.metadata?.raw_type;
  const currentIsTemporaryAxis = Boolean(
    current.meeting?.pending_binding
      || current.meeting?.source === 'local_simulation'
      || current.meeting?.source === 'lark_reserve_pending',
  );
  const currentIsCarryableFallbackAxis = Boolean(
    currentIsTemporaryAxis
      || current.meeting?.source === 'open_meeting_session',
  );
  const canBootstrapFromContext = event.type === 'participant_join'
    && isMeetingContextBootstrapEvent(rawEventType)
    && (currentIsTemporaryAxis || shouldCreatePendingTimeline(current));
  const currentIsRealAxis = isRealMeetingAxis(current.meeting);
  const endMatchesKnownRealAxis = event.type === 'meeting_end' && currentIsRealAxis && incomingMatchesCurrent;
  const endCanReconstructFromPayloadStart = event.type === 'meeting_end'
    && payloadStartMs != null
    && (currentIsTemporaryAxis || incomingMatchesCurrent || shouldCreatePendingTimeline(current));
  if (realDemoSession.active && event.type === 'meeting_end' && !endMatchesKnownRealAxis) {
    return {
      ok: false,
      ignored_reason: 'real_demo_requires_meeting_start_event',
      event,
      timeline_started: false,
      state: current,
    };
  }
  if (event.type === 'meeting_end' && !endMatchesKnownRealAxis && !endCanReconstructFromPayloadStart) {
    return {
      ok: false,
      ignored_reason: 'meeting_end_without_known_start_axis',
      event,
      timeline_started: false,
      state: current,
    };
  }
  const isStartLikeEvent = event.type === 'meeting_start' || canBootstrapFromContext || endCanReconstructFromPayloadStart;
  const shouldStartTimeline = isStartLikeEvent && (!incomingMatchesCurrent || currentIsTemporaryAxis || shouldCreatePendingTimeline(current));
  if (event.type === 'meeting_start' || event.type === 'meeting_end' || shouldStartTimeline) meetingPatch.source = source;
  if (isStartLikeEvent) meetingPatch.pending_binding = false;
  if (shouldStartTimeline && hasIncomingIdentity) clearInheritedFieldsForNewMeeting(meetingPatch, incomingIdentity);
  const shouldCarryPendingSequence = shouldStartTimeline
    && currentIsCarryableFallbackAxis
    && shouldCarrySequenceIntoNewAxis(current, meetingPatch);
  const rebasedPendingSequence = shouldCarryPendingSequence
    ? rebasePendingSequence(current.sequence ?? [], meetingPatch)
    : [];
  const initialEvents = shouldStartTimeline && event.type === 'meeting_end'
    ? [inferredStartEventForEnd(payload, source)]
    : [];
  const base = shouldStartTimeline
    ? buildTimeline({
      meeting: meetingPatch,
      segments: [],
      events: initialEvents,
      sequence: rebasedPendingSequence,
    })
    : current;
  if (event.type === 'meeting_end') {
    meetingPatch.end_time = isoAtMeetingOffset({ ...base.meeting, ...meetingPatch }, event.time_ms);
  }
  const eventMap = new Map((base.events ?? []).map((x) => [x.id, x]));
  eventMap.set(event.id, event);
  const next = mergeTimelineWithRebasedAnnotations(base, {
    meeting: event.type === 'meeting_start' || event.type === 'meeting_end' || meetingPatch.minute_token
      ? meetingPatch
      : base.meeting,
    events: [...eventMap.values()],
  });
  const savedState = await saveAndBroadcast(next, 'state');
  if (realDemoSession.active && isRealMeetingAxis(savedState.meeting)) {
    saveRealDemoSessionState({
      ...realDemoSession,
      last_real_axis_at: new Date().toISOString(),
      last_real_axis_source: source,
    });
  }
  const shouldAppendSyntheticAnnotations = !opts.suppress_auto_annotations
    && (event.type === 'meeting_start' || shouldStartTimeline);
  const autoAnnotation = shouldAppendSyntheticAnnotations
    ? await maybeAppendAutoAcceptanceAnnotation(savedState, source)
    : { state: savedState, annotation: null };
  const deviceAnnotation = shouldAppendSyntheticAnnotations
    ? await maybeAppendDeviceSimulatorAnnotation(autoAnnotation.state, source)
    : { state: autoAnnotation.state, annotation: null };
  return {
    ok: true,
    event,
    timeline_started: shouldStartTimeline,
    auto_acceptance_annotation: autoAnnotation.annotation,
    device_simulator_annotation: deviceAnnotation.annotation,
    state: deviceAnnotation.state,
  };
}

function larkDomain() {
  return lark.baseUrl.includes('larksuite') ? Lark.Domain.Lark : Lark.Domain.Feishu;
}

const explicitWsTimelineEventTypes = new Set([
  meetingEventRequirements.direct_meeting.start_event,
  meetingEventRequirements.direct_meeting.end_event,
  meetingEventRequirements.reserve_meeting.start_event,
  meetingEventRequirements.reserve_meeting.end_event,
  meetingEventRequirements.meeting_context.start_event,
  meetingEventRequirements.meeting_context.end_event,
]);

async function handleWsTimelineEvent(parsed = {}, raw = {}, forcedEventType = null) {
  const eventType = forcedEventType ?? eventTypeFromParsed(parsed);
  larkWsStatus.last_event_at = new Date().toISOString();
  larkWsStatus.last_event_type = String(eventType || 'unknown');
  const candidate = timelineEventCandidate(eventType, parsed);
  const logEntry = pushLarkEventLog({
    transport: 'ws_long_connection',
    event_type: eventType,
    timeline_candidate: candidate,
    parsed,
    preview: parsed,
  });

  if (!candidate) {
    logEntry.ignored_reason = 'not_meeting_related';
    persistLarkEventLog();
    return `ignored ${eventType || 'unknown'} event`;
  }

  const result = await processLarkEventPayload(buildWsPayload(eventType, parsed, raw), { source: 'lark_ws_event' });
  logEntry.timeline_processed = result.ok !== false;
  logEntry.timeline_started = result.timeline_started;
  if (result.ignored_reason) logEntry.ignored_reason = result.ignored_reason;
  persistLarkEventLog();
  larkWsStatus.last_timeline_event_at = new Date().toISOString();
  larkWsStatus.last_timeline_event_type = String(eventType || result.event.type);
  return result;
}

function createWsEventDispatcher() {
  const dispatcher = new Lark.EventDispatcher({
    loggerLevel: Lark.LoggerLevel.warn,
  }).register(Object.fromEntries(
    [...explicitWsTimelineEventTypes].map((eventType) => [
      eventType,
      async (data) => handleWsTimelineEvent(data, data, eventType),
    ]),
  ));
  const defaultInvoke = dispatcher.invoke.bind(dispatcher);
  dispatcher.invoke = async (raw, params) => {
    let parsed = raw;
    try {
      parsed = dispatcher.requestHandle?.parse(raw) ?? raw;
    } catch {
      parsed = raw;
    }
    const eventType = eventTypeFromParsed(parsed);

    if (eventType === 'app_ticket') {
      const logEntry = pushLarkEventLog({
        transport: 'ws_long_connection',
        event_type: eventType,
        timeline_candidate: false,
        parsed,
        preview: parsed,
      });
      logEntry.ignored_reason = 'sdk_app_ticket';
      persistLarkEventLog();
      return defaultInvoke(raw, params);
    }
    if (explicitWsTimelineEventTypes.has(eventType)) {
      return defaultInvoke(raw, params);
    }

    return handleWsTimelineEvent(parsed, raw, eventType);
  };
  return dispatcher;
}

function startLarkWsReceiver() {
  if (process.env.LARK_WS_EVENTS === '0') {
    larkWsStatus.enabled = false;
    larkWsStatus.state = 'disabled';
    return;
  }
  if (!lark.isConfigured) {
    larkWsStatus.enabled = false;
    larkWsStatus.state = 'missing_credentials';
    return;
  }

  larkWsStatus.enabled = true;
  larkWsStatus.state = 'connecting';
  larkWsStatus.error = null;
  larkWsClient = new Lark.WSClient({
    appId: process.env.LARK_APP_ID,
    appSecret: process.env.LARK_APP_SECRET,
    domain: larkDomain(),
    loggerLevel: Lark.LoggerLevel.warn,
    autoReconnect: true,
    handshakeTimeoutMs: 15_000,
    source: 'meeting-timeline-demo',
    onReady: () => {
      larkWsStatus.state = 'connected';
      larkWsStatus.error = null;
      console.log('Lark WS receiver connected');
    },
    onReconnecting: () => {
      larkWsStatus.state = 'reconnecting';
    },
    onReconnected: () => {
      larkWsStatus.state = 'connected';
      larkWsStatus.error = null;
    },
    onError: (error) => {
      larkWsStatus.state = 'failed';
      larkWsStatus.error = error.message ?? String(error);
      console.error('Lark WS receiver error:', error);
    },
  });
  larkWsClient.start({ eventDispatcher: createWsEventDispatcher() }).catch((error) => {
    larkWsStatus.state = 'failed';
    larkWsStatus.error = error.message ?? String(error);
    console.error('Lark WS receiver start failed:', error);
  });
}

function publicLarkWsStatus() {
  const connection = larkWsClient?.getConnectionStatus?.();
  return {
    ...larkWsStatus,
    registered_event_types: [...explicitWsTimelineEventTypes],
    state: connection?.state ?? larkWsStatus.state,
    reconnect_attempts: connection?.reconnectAttempts ?? 0,
    last_connect_time: connection?.lastConnectTime ? new Date(connection.lastConnectTime).toISOString() : null,
    next_connect_time: connection?.nextConnectTime ? new Date(connection.nextConnectTime).toISOString() : null,
  };
}

function publicCapabilityStatus() {
  return {
    meeting_events: {
      direct_meeting: {
        ...meetingEventRequirements.direct_meeting,
        permission_url: permissionUrlForScopes(meetingEventRequirements.direct_meeting.required_scope),
      },
      reserve_meeting: {
        ...meetingEventRequirements.reserve_meeting,
        permission_url: permissionUrlForScopes(meetingEventRequirements.reserve_meeting.required_scope),
      },
      meeting_context: {
        ...meetingEventRequirements.meeting_context,
        permission_url: permissionUrlForScopes(meetingEventRequirements.meeting_context.required_scope),
      },
      meeting_lookup: {
        ...meetingEventRequirements.meeting_lookup,
        permission_url: permissionUrlForScopes(['vc:meeting.search:read', 'vc:meeting.meetingid:read']),
      },
    },
    reserve: {
      ...capabilityStatus.reserve,
      permission_url: permissionUrlForScopes('vc:reserve'),
    },
    meeting_lookup: {
      ...capabilityStatus.meeting_lookup,
      permission_url: capabilityStatus.meeting_lookup.permission_url
        ?? permissionUrlForScopes(['vc:meeting.search:read', 'vc:meeting.meetingid:read']),
    },
    meeting_search: {
      ...capabilityStatus.meeting_search,
      permission_url: capabilityStatus.meeting_search.permission_url
        ?? permissionUrlForScopes('vc:meeting.search:read'),
    },
  };
}

function publicRealMeetingProbeStatus(req) {
  const config = larkConfigPayload(req);
  const observed = realMeetingProbe.started_at
    ? findMeetingEntryAfter(config, realMeetingProbe.started_at)
    : null;
  const missedEvent = !observed ? missedMeetingWindowForProbe(req) : null;
  const now = Date.now();
  const startedMs = isoMs(realMeetingProbe.started_at);
  const elapsedMs = startedMs == null ? 0 : Math.max(0, now - startedMs);
  const timedOut = Boolean(realMeetingProbe.active && !observed && elapsedMs > realMeetingProbe.timeout_ms);
  const status = observed
    ? 'passed'
    : timedOut
      ? 'timeout'
      : realMeetingProbe.active
        ? 'waiting'
        : 'idle';
  return {
    status,
    active: Boolean(realMeetingProbe.active && status === 'waiting'),
    started_at: realMeetingProbe.started_at,
    timeout_ms: realMeetingProbe.timeout_ms,
    elapsed_ms: elapsedMs,
    note: realMeetingProbe.note,
    auto_search: {
      ...realMeetingProbe.auto_search,
      server_loop: {
        enabled: Boolean(realMeetingProbe.auto_search?.enabled),
        scheduled: Boolean(probeAutoSearchTimer),
        in_flight: probeAutoSearchInFlight,
        interval_ms: probeAutoSearchIntervalMs(),
      },
    },
    required_event: meetingEventRequirements.direct_meeting.start_event,
    accepted_events: [
      meetingEventRequirements.direct_meeting.start_event,
      meetingEventRequirements.meeting_context.start_event,
    ],
    required_scope: meetingEventRequirements.direct_meeting.required_scope,
    observed_event: observed,
    missed_event: missedEvent,
    next_step: status === 'idle'
      ? 'probe 是可选验收窗口；即使不点击，服务端长连接也会被动监听真实会议事件。需要严格验收时可点击开始后再开会。'
      : status === 'waiting'
        ? missedEvent
          ? missedEvent.message
          : '现在直接开启飞书会议，等待长连接事件进入 demo。'
        : status === 'passed'
          ? '真实直开会议事件已进入 demo，时间轴应自动建轴。'
          : missedEvent
            ? missedEvent.message
            : `未收到真实 ${meetingEventRequirements.direct_meeting.start_event}；检查事件订阅、应用可见范围和 ${meetingEventRequirements.direct_meeting.required_scope} 权限。`,
  };
}

async function acceptanceReportPayload(req) {
  const [state, readiness, diagnostics] = await Promise.all([
    store.load(),
    readinessPayload(req),
    publicDeliveryDiagnostics(req),
  ]);
  const config = larkConfigPayload(req);
  const direct = {
    ...meetingEventRequirements.direct_meeting,
    permission_url: permissionUrlForScopes(meetingEventRequirements.direct_meeting.required_scope),
  };
  const context = {
    ...meetingEventRequirements.meeting_context,
    permission_url: permissionUrlForScopes(meetingEventRequirements.meeting_context.required_scope),
  };
  const oauthScopes = String(authState.token?.scope ?? '').split(/\s+/).filter(Boolean);
  const realAxis = isRealMeetingAxis(state.meeting);
  const annotationCount = openAnnotationItems(state).length;
  const probeStatus = readiness.real_meeting_probe;
  const probeObservedEvent = probeStatus?.observed_event ?? null;
  const hasProbeWindow = Boolean(probeStatus?.started_at);
  const probeStartedMs = isoMs(probeStatus?.started_at);
  const autoSearchLastAttemptMs = isoMs(probeStatus?.auto_search?.last_attempt_at);
  const autoSearchLastResult = probeStatus?.auto_search?.last_result ?? null;
  const autoSearchBindingAfterProbe = Boolean(
    hasProbeWindow
      && realAxis
      && state.meeting?.source === 'lark_probe_auto_search'
      && autoSearchLastResult?.status === 'bound'
      && autoSearchLastAttemptMs != null
      && probeStartedMs != null
      && autoSearchLastAttemptMs >= probeStartedMs
      && String(autoSearchLastResult.selected_meeting_id ?? '') === String(state.meeting?.meeting_id ?? ''),
  );
  const annotationBinding = annotationBindingSummary(state, {
    probe_started_at: hasProbeWindow ? probeStatus.started_at : null,
  });
  const transcriptStatus = transcriptStatusForState(state);
  const currentValidationRealEvent = hasProbeWindow
    ? Boolean(probeObservedEvent)
    : diagnostics.evidence.real_event_count > 0;
  const currentValidationRealEntry = hasProbeWindow
    ? (currentValidationRealEvent || autoSearchBindingAfterProbe)
    : (diagnostics.evidence.real_event_count > 0 || realAxis);
  const currentValidationAnnotation = realAxis && (hasProbeWindow
    ? annotationBinding.since_probe_real_axis_count > 0
    : annotationBinding.real_axis_count > 0);
  const currentValidationReady = currentValidationRealEntry && currentValidationAnnotation;
  return {
    generated_at: new Date().toISOString(),
    ready: readiness.ready,
    status: readiness.ready
      ? 'ready'
      : realAxis
        ? annotationCount > 0 ? 'real_axis_with_annotation' : 'real_axis_waiting_annotation'
        : diagnostics.receiver.ws_state === 'connected'
          ? 'waiting_for_lark_event_delivery'
          : 'receiver_not_ready',
    objective: '用户直接开启飞书会议后自动建轴，开放标注实时进入时间轴；转写会后处理。',
    current_evidence: {
      ws_state: diagnostics.receiver.ws_state,
      registered_event_types: diagnostics.receiver.registered_event_types,
      real_event_count: diagnostics.evidence.real_event_count,
      ws_event_count: diagnostics.evidence.ws_event_count,
      ws_timeline_candidate_count: diagnostics.evidence.ws_timeline_candidate_count,
      ws_timeline_processed_count: diagnostics.evidence.ws_timeline_processed_count,
      ws_timeline_started_count: diagnostics.evidence.ws_timeline_started_count,
      ws_ignored_count: diagnostics.evidence.ws_ignored_count,
      current_meeting_source: diagnostics.evidence.current_meeting_source,
      current_meeting_pending: diagnostics.evidence.current_meeting_pending,
      real_meeting_axis_active: realAxis,
      open_annotation_count: annotationCount,
      annotation_binding: annotationBinding,
      transcript_status: transcriptStatus,
      probe: readiness.real_meeting_probe,
      auto_acceptance: publicAutoAcceptanceStatus(),
      device_simulator: publicDeviceSimulatorStatus(),
      passive_meeting_scan: publicPassiveMeetingScanStatus(),
      real_demo_session: publicRealDemoSessionStatus(),
      oauth_scopes: oauthScopes,
      meeting_search_oauth_scope_present: hasOAuthScope('vc:meeting.search:read'),
    },
    current_validation: {
      ready: currentValidationReady,
      scoped_by_probe: hasProbeWindow,
      probe_started_at: probeStatus?.started_at ?? null,
      probe_status: probeStatus?.status ?? 'idle',
      observed_event: probeObservedEvent,
      real_event_after_probe: hasProbeWindow ? Boolean(probeObservedEvent) : null,
      auto_search_binding_after_probe: hasProbeWindow ? autoSearchBindingAfterProbe : null,
      real_entry_after_probe: hasProbeWindow ? currentValidationRealEntry : null,
      real_meeting_axis_active: realAxis,
      open_annotation_count: annotationCount,
      annotation_count_after_probe: hasProbeWindow ? annotationBinding.since_probe_count : null,
      real_axis_annotation_count: annotationBinding.real_axis_count,
      real_axis_annotation_count_after_probe: hasProbeWindow ? annotationBinding.since_probe_real_axis_count : null,
      annotation_on_real_axis: currentValidationAnnotation,
      annotation_binding_state: annotationBinding.binding_state,
      last_open_annotation_binding: annotationBinding.last,
      note: hasProbeWindow
        ? '本区块只把 probe started_at 之后的真实会议事件和标注计入本次验收。'
        : '未启动 probe 时给出全局证据；服务端长连接会常驻被动监听真实会议事件，probe 只是可选验收辅助。',
    },
    required_main_path: {
      transport: config.ws_event_receiver?.enabled
        ? '飞书长连接事件'
        : config.event_callback_public_https ? '公网 HTTPS webhook' : '未就绪',
      event_subscription: [
        direct.start_event,
        direct.end_event,
      ],
      app_permission: direct.required_scope,
      permission_label: direct.permission_label,
      docs: [
        direct.start_doc_url,
        direct.end_doc_url,
      ],
      permission_url: direct.permission_url,
    },
    open_platform_checklist: diagnostics.open_platform_checklist,
    optional_fallbacks: [
      {
        name: 'join/leave 会议上下文事件建轴',
        events: [context.start_event, context.end_event],
        app_permission: context.required_scope,
        permission_label: context.permission_label,
        docs: [context.start_doc_url, context.end_doc_url],
        permission_url: context.permission_url,
      },
      {
        name: '当前 OAuth 用户会议扫描',
        required_scope: 'vc:meeting.search:read',
        scope_present: hasOAuthScope('vc:meeting.search:read'),
        passive_scan: publicPassiveMeetingScanStatus(),
        permission_url: permissionUrlForScopes('vc:meeting.search:read'),
      },
      {
        name: '应用身份扫描租户近期会议（实验）',
        required_scope: 'vc:meeting.search:read',
        endpoint: '/api/lark/bind-tenant-latest-meeting',
        status: capabilityStatus.meeting_search.status,
        error: capabilityStatus.meeting_search.status === 'invalid_token'
          ? capabilityStatus.meeting_search.error
          : null,
        note: capabilityStatus.meeting_search.status === 'invalid_token'
          ? '当前飞书接口拒绝应用/租户 access token，不作为本 demo 可靠兜底；请优先使用当前 OAuth 用户会议扫描。'
          : '仅作为实验入口；事件投递失败时，优先使用当前 OAuth 用户会议扫描或会议号绑定。',
        permission_url: permissionUrlForScopes('vc:meeting.search:read'),
      },
      {
        name: '按会议号/链接手动绑定',
        required_scope: 'vc:meeting.search:read or vc:meeting.meetingid:read',
        permission_url: permissionUrlForScopes(['vc:meeting.search:read', 'vc:meeting.meetingid:read']),
      },
    ],
    acceptance_steps: [
      {
        id: 'open_platform_events',
        done: diagnostics.receiver.missing_registered_events.length === 0 && diagnostics.receiver.missing_fallback_events.length === 0,
        text: `在开放平台事件订阅中启用长连接，并订阅 ${direct.start_event} / ${direct.end_event}；可同时订阅 ${context.start_event} / ${context.end_event}`,
      },
      {
        id: 'open_platform_permission',
        done: diagnostics.evidence.real_event_count > 0 || realAxis,
        text: `开通并发布应用权限 ${direct.required_scope}（${direct.permission_label}），确认应用可见范围包含开会账号`,
        action_url: direct.permission_url,
      },
      {
        id: 'passive_meeting_listener',
        done: diagnostics.receiver.ws_state === 'connected' || Boolean(config.event_callback_public_https),
        text: '保持服务运行；飞书长连接或公网 webhook 会被动接收真实会议事件，直接开会即可触发建轴',
      },
      {
        id: 'optional_probe',
        optional: true,
        done: readiness.real_meeting_probe.status === 'waiting' || readiness.real_meeting_probe.status === 'passed',
        text: '可选：点击“启动验收探针”或 POST /api/lark/real-meeting-probe/start，用于缩小验收时间窗并启用自动扫描兜底；它不会创建会议轴',
      },
      {
        id: 'direct_start_meeting',
        done: currentValidationRealEntry,
        text: `直接在飞书客户端开启会议，等待 ${direct.start_event} / ${context.start_event} 投递；只有显式开启扫描兜底后，才会由当前账号会议扫描绑定真实会议`,
        evidence: currentValidationRealEvent
          ? 'lark_event'
          : autoSearchBindingAfterProbe
            ? 'probe_auto_search'
            : state.meeting?.source === 'lark_passive_meeting_scan'
              ? 'passive_meeting_scan'
            : null,
      },
      {
        id: 'write_open_annotation',
        done: currentValidationAnnotation,
        text: autoAcceptance.enabled
          ? '自动验收标注已开启；真实会议轴建立后会通过开放标注接口自动写入一条标注，也可手动 POST /api/annotations'
          : '通过 POST /api/annotations 或页面“写入验收标注”写入一条标注，确认 SSE 实时刷新时间轴',
      },
    ],
    diagnostics,
  };
}

async function realDemoStatusPayload(req) {
  const [state, readiness, diagnostics] = await Promise.all([
    store.load(),
    readinessPayload(req),
    publicDeliveryDiagnostics(req),
  ]);
  const config = larkConfigPayload(req);
  const preparedAt = realDemoSession.prepared_at ?? null;
  const deliveredSincePrepare = eventsAtOrAfter(realDeliveredTimelineEvents(config), preparedAt);
  const auth = publicAuthState();
  const annotationBinding = readiness.current?.annotation_binding
    ?? annotationBindingSummary(state, { since_at: preparedAt });
  const preparedAnnotationBinding = annotationBindingSummary(state, { since_at: preparedAt });
  const receiverReady = Boolean(
    diagnostics.receiver?.ws_state === 'connected'
      || diagnostics.receiver?.http_callback_public_https,
  );
  const openMeetingSessionReady = true;
  const openAnnotationReady = Boolean(readiness.checks?.find((check) => check.id === 'open_annotation_api')?.ok);
  const realtimeReady = Boolean(readiness.checks?.find((check) => check.id === 'realtime_timeline')?.ok);
  const axisObservedAfterPrepare = realDemoAxisObservedAfterPrepare();
  const realAxisReady = Boolean(isRealMeetingAxis(state.meeting) && axisObservedAfterPrepare);
  const realAnnotationReady = Boolean(realAxisReady && preparedAnnotationBinding.since_real_axis_count > 0);
  const scanReady = Boolean(auth.meeting_search?.usable && passiveMeetingScan.enabled);
  const futureMeetingEntryReady = Boolean(receiverReady || scanReady || openMeetingSessionReady);
  const directEventSeen = Boolean(
    deliveredSincePrepare.some((event) => (
      isDirectMeetingStartEvent(event.event_type)
        || isMeetingContextBootstrapEvent(event.event_type)
        || event.timeline_processed
    )),
  );
  const readyToOpenMeeting = Boolean(
    futureMeetingEntryReady
      && openAnnotationReady
      && realtimeReady
      && deviceSimulator.enabled
  );
  const deviceStream = publicDeviceStreamSimulatorStatus();
  const eventAudit = diagnostics.real_meeting_event_audit ?? null;

  const blockers = [];
  if (!realDemoSession.active) {
    blockers.push({
      id: 'prepare_real_demo',
      severity: 'error',
      text: '尚未进入真实等待状态；需要点击“进入真实等待状态”。',
    });
  }
  if (!receiverReady && !scanReady && !openMeetingSessionReady && !realAxisReady) {
    blockers.push({
      id: 'event_receiver',
      severity: 'error',
      text: '飞书事件接收通道未就绪，且扫描兜底不可用；直接开会无法自动建轴。',
    });
  } else if (!receiverReady && scanReady && !realAxisReady) {
    blockers.push({
      id: 'event_receiver',
      severity: 'warning',
      text: '飞书事件接收通道未就绪；但当前账号会议扫描兜底可用，直接开会后仍可建轴。',
    });
  } else if (!receiverReady && !scanReady && openMeetingSessionReady && !realAxisReady) {
    blockers.push({
      id: 'event_receiver',
      severity: 'info',
      text: '飞书事件/扫描未就绪；开放会议会话 API 可作为当前产品主路径建轴。',
    });
  }
  if (!deviceSimulator.enabled) {
    blockers.push({
      id: 'device_simulator',
      severity: 'warning',
      text: '虚拟墨水屏标注未开启；真实会议建轴后不会自动写入 demo 设备标注。',
    });
  }
  if (!openAnnotationReady || !realtimeReady) {
    blockers.push({
      id: 'annotation_realtime',
      severity: 'error',
      text: '开放标注接口或 SSE 实时刷新未就绪。',
    });
  }
  if (!auth.meeting_search?.usable) {
    const authStart = meetingSearchOAuthStart(req);
    blockers.push({
      id: 'meeting_scan_oauth',
      severity: receiverReady ? 'warning' : 'error',
      text: auth.meeting_search?.reason === 'oauth_token_expired'
        ? '当前用户会议扫描兜底不可用：OAuth 已过期，需要重新授权 vc:meeting.search:read。'
        : auth.meeting_search?.reason === 'missing_scope'
          ? '当前用户会议扫描兜底不可用：缺少 vc:meeting.search:read。'
          : '当前用户会议扫描兜底不可用，需要登录并授权 vc:meeting.search:read。',
      action_url: authStart.redirect_url,
      oauth_redirect_url: authStart.redirect_url,
      oauth_json_url: authStart.json_url,
      permission_url: permissionUrlForScopes('vc:meeting.search:read'),
    });
  }
  if (!directEventSeen && !realAxisReady) {
    blockers.push({
      id: 'no_real_meeting_yet',
      severity: 'info',
      text: `尚未收到真实飞书会议事件，也未绑定真实会议轴；现在需要直接开启飞书会议。`,
    });
    if (eventAudit?.status && eventAudit.status !== 'event_delivery_ok') {
      blockers.push({
        id: 'meeting_event_audit',
        severity: eventAudit.severity === 'error' ? 'error' : eventAudit.severity === 'warning' ? 'warning' : 'info',
        text: eventAudit.next_action,
      });
    }
  }
  if (realAxisReady && !realAnnotationReady) {
    blockers.push({
      id: 'no_real_annotation_yet',
      severity: 'warning',
      text: '真实会议轴已建立，但还没有标注落到真实轴。',
    });
  }

  const requiredBlockers = blockers.filter((item) => item.severity === 'error');
  const stateName = realAnnotationReady
    ? 'complete'
    : realAxisReady
      ? 'real_axis_waiting_annotation'
      : requiredBlockers.length
        ? 'needs_setup'
        : 'ready_to_open_meeting';
  const canOpenMeetingNow = Boolean(futureMeetingEntryReady && openAnnotationReady && realtimeReady);
  const scanNeedsReauth = Boolean(!auth.meeting_search?.usable);
  const operatorPhase = realAnnotationReady
    ? 'complete'
    : realAxisReady
      ? 'write_open_annotation'
      : canOpenMeetingNow
        ? 'open_lark_meeting'
        : 'fix_event_receiver_or_realtime';
  const operatorNextAction = operatorPhase === 'complete'
    ? '已验证真实会议建轴和开放标注实时落轴；会议结束后再处理转写。'
    : operatorPhase === 'write_open_annotation'
      ? '真实会议轴已建立；通过 POST /api/annotations 写入一条带 captured_at_ms 的标注，或等待设备流自动写入。'
      : operatorPhase === 'open_lark_meeting'
        ? scanNeedsReauth
          ? '可以直接开启飞书会议等待事件投递；同时建议重新授权 vc:meeting.search:read，以便事件未投递时用当前账号会议扫描兜底。'
          : '直接开启飞书会议；服务端会等待飞书事件投递，必要时用当前账号会议扫描兜底建轴。'
        : blockers.find((item) => item.severity === 'error')?.text ?? '先修复事件接收或实时标注通道。';

  return {
    generated_at: new Date().toISOString(),
    status: stateName,
    ready_to_open_meeting: readyToOpenMeeting && requiredBlockers.length === 0 && !realAxisReady,
    complete: realAnnotationReady,
    summary: stateName === 'complete'
      ? '真实会议轴和实时标注均已验证。'
      : stateName === 'real_axis_waiting_annotation'
        ? '真实会议轴已建立，等待设备标注进入时间轴。'
        : stateName === 'ready_to_open_meeting'
          ? '服务端已准备好；现在直接开启飞书会议即可观察时间轴自动建轴和设备流标注。'
          : '真实演示尚有必要条件未满足。',
	    gates: {
	      real_demo_prepared: Boolean(realDemoSession.active),
	      event_receiver_ready: receiverReady,
	      ws_connected: diagnostics.receiver?.ws_state === 'connected',
	      real_event_seen: directEventSeen,
      event_audit_status: eventAudit?.status ?? null,
      event_audit_root_cause: eventAudit?.root_cause ?? null,
      open_meeting_session_ready: openMeetingSessionReady,
      ws_event_count: diagnostics.evidence?.ws_event_count ?? 0,
	      ws_timeline_candidate_count: diagnostics.evidence?.ws_timeline_candidate_count ?? 0,
	      ws_timeline_started_count: diagnostics.evidence?.ws_timeline_started_count ?? 0,
	      scan_fallback_ready: scanReady,
	      future_meeting_entry_ready: futureMeetingEntryReady,
	      open_annotation_ready: openAnnotationReady,
      realtime_stream_ready: realtimeReady,
      device_simulator_enabled: Boolean(deviceSimulator.enabled),
	      device_stream_enabled: Boolean(deviceStream.enabled || deviceStream.timer_active),
	      device_stream_status: deviceStream.status,
	      device_stream_count: deviceStream.count,
	      real_meeting_axis_active: realAxisReady,
	      real_axis_annotation_count: preparedAnnotationBinding.since_real_axis_count,
	      real_axis_annotation_count_total: annotationBinding.real_axis_count,
	    },
    blockers,
    operator_runbook: {
      phase: operatorPhase,
      primary_next_action: operatorNextAction,
      can_open_meeting_now: canOpenMeetingNow,
      oauth_scan_recommended: Boolean(scanNeedsReauth && receiverReady),
      oauth_scan_required_for_main_path: false,
      oauth_scan_required_if_event_delivery_fails: scanNeedsReauth,
      meeting_session_start_endpoint: localUrlFor(req, '/api/meeting-session/start'),
      meeting_session_end_endpoint: localUrlFor(req, '/api/meeting-session/end'),
      annotation_endpoint: localUrlFor(req, '/api/annotations'),
      annotation_stream_url: localUrlFor(req, '/api/stream'),
      progress_stream_url: localUrlFor(req, '/api/lark/real-demo/progress-stream'),
      success_condition: {
        real_meeting_axis_active: true,
        real_axis_annotation_count_min: 1,
        transcript_required_realtime: false,
      },
      evidence_to_watch: [
        'gates.real_meeting_axis_active',
        'gates.real_axis_annotation_count',
        'gates.ws_event_count',
        'realtime_stream.last_broadcast_real_axis_annotation_count',
      ],
    },
    next_action: blockers.find((item) => item.severity === 'error')?.text
      ?? operatorNextAction
      ?? blockers.find((item) => item.id === 'no_real_meeting_yet')?.text
      ?? blockers[0]?.text
      ?? '直接开启飞书会议。',
	    auth: {
	      meeting_search: auth.meeting_search,
	      authenticated: auth.authenticated,
	      token_present: auth.token_present,
	      expired: auth.expired,
	      refresh_token_present: auth.refresh_token_present,
	    },
	    auth_start: auth.meeting_search?.usable ? null : meetingSearchOAuthStart(req),
	    real_demo_session: publicRealDemoSessionStatus(),
    device_simulator: publicDeviceSimulatorStatus(),
    device_stream_simulator: deviceStream,
    passive_meeting_scan: publicPassiveMeetingScanStatus(),
    realtime_stream: publicStreamStatus(),
	    evidence: {
	      meeting: state.meeting,
	      annotation_binding: preparedAnnotationBinding,
	      real_event_count: deliveredSincePrepare.length,
	      real_event_count_total: diagnostics.evidence?.real_event_count ?? 0,
	      ws_event_count: diagnostics.evidence?.ws_event_count ?? 0,
	      ws_timeline_candidate_count: diagnostics.evidence?.ws_timeline_candidate_count ?? 0,
	      ws_timeline_started_count: diagnostics.evidence?.ws_timeline_started_count ?? 0,
	      last_real_event: deliveredSincePrepare[0] ?? null,
	      last_real_event_total: diagnostics.evidence?.last_real_event ?? null,
	      last_open_annotation: diagnostics.evidence?.last_open_annotation ?? null,
	      real_meeting_event_audit: eventAudit,
	    },
  };
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitForRealDemoStatus(req, options = {}) {
  const timeoutMs = Math.min(Math.max(Number(options.timeout_ms ?? 120_000), 1_000), 10 * 60_000);
  const intervalMs = Math.min(Math.max(Number(options.interval_ms ?? 500), 100), 5000);
  const startedAt = Date.now();
  let last = await realDemoStatusPayload(req);
  while (!last.complete && Date.now() - startedAt < timeoutMs) {
    await delay(intervalMs);
    last = await realDemoStatusPayload(req);
  }
  const elapsedMs = Date.now() - startedAt;
  return {
    observed: Boolean(last.complete),
    timed_out: !last.complete,
    elapsed_ms: elapsedMs,
    timeout_ms: timeoutMs,
    interval_ms: intervalMs,
    completed_at: last.complete ? new Date().toISOString() : null,
    status: last.status,
    summary: last.summary,
    result: last,
  };
}

function realDemoCompletionEvidence(observed = {}, streamStatus = publicStreamStatus()) {
  const result = observed.result ?? observed;
  const gates = result?.gates ?? {};
  const evidence = result?.evidence ?? {};
  const eventAudit = evidence.real_meeting_event_audit ?? null;
  const deviceStream = result?.device_stream_simulator ?? publicDeviceStreamSimulatorStatus();
  const meeting = evidence.meeting ?? {};
  const meetingSource = meeting.source ?? null;
  const eventAxisBuilt = ['lark_ws_event', 'lark_http_event', 'lark_http_local_event'].includes(meetingSource);
  const meetingEntryReady = Boolean(
    gates.event_receiver_ready
      || gates.scan_fallback_ready
      || gates.open_meeting_session_ready
      || gates.real_meeting_axis_active,
  );
  const annotationCount = Number(gates.real_axis_annotation_count ?? 0);
  const broadcastAnnotationCount = Number(streamStatus.last_state_summary?.real_axis_annotation_count ?? 0);
  const requirements = [
    {
      id: 'event_receiver_ready',
      ok: Boolean(meetingEntryReady && gates.open_annotation_ready && gates.realtime_stream_ready),
      required: '至少一个会议入口、开放标注接口和实时 SSE 刷新通道都已就绪。',
      observed: `event_receiver=${Boolean(gates.event_receiver_ready)}, scan=${Boolean(gates.scan_fallback_ready)}, open_session=${Boolean(gates.open_meeting_session_ready)}, open_annotation=${Boolean(gates.open_annotation_ready)}, realtime_stream=${Boolean(gates.realtime_stream_ready)}`,
    },
    {
      id: 'real_meeting_axis_active',
      ok: Boolean(gates.real_meeting_axis_active),
      required: '直接开启飞书会议后，服务端建立真实会议时间轴。',
      observed: `real_axis=${Boolean(gates.real_meeting_axis_active)}, source=${meeting.source ?? 'none'}, meeting_id=${meeting.meeting_id ?? 'none'}`,
    },
    {
      id: 'event_meeting_axis_built',
      ok: eventAxisBuilt,
      required: '严格事件验收：真实会议轴由飞书会议开始事件建立，而不是扫描兜底或手动绑定。',
      observed: `event_axis=${eventAxisBuilt}, axis_creation_mode=${eventAxisBuilt ? 'meeting_start_event' : gates.real_meeting_axis_active ? meetingSource ?? 'unknown' : 'not_built'}, source=${meetingSource ?? 'none'}`,
    },
    {
      id: 'open_annotation_on_real_axis',
      ok: annotationCount > 0,
      required: '开放标注接口写入的标注已经落到真实会议轴。',
      observed: `real_axis_annotation_count=${annotationCount}`,
    },
    {
      id: 'realtime_state_broadcast',
      ok: broadcastAnnotationCount > 0,
      required: '标注落轴后通过 SSE 广播，页面可实时刷新。',
      observed: `last_broadcast_real_axis_annotation_count=${broadcastAnnotationCount}, last_broadcast_at=${streamStatus.last_broadcast_at ?? 'none'}`,
    },
    {
      id: 'transcript_post_meeting_only',
      ok: true,
      required: '文字转写不阻塞实时标注；会议结束后再同步或导入。',
      observed: 'transcript_required_realtime=false',
    },
  ];
  return {
    real_demo_complete: Boolean(observed.observed ?? result?.complete),
    requirements,
    real_meeting_axis_active: Boolean(gates.real_meeting_axis_active),
    meeting_source: meetingSource,
    event_axis_built: eventAxisBuilt,
    strict_event_axis_complete: Boolean(eventAxisBuilt && annotationCount > 0 && broadcastAnnotationCount > 0),
    axis_creation_mode: eventAxisBuilt
      ? 'meeting_start_event'
      : gates.real_meeting_axis_active
        ? meetingSource ?? 'unknown_real_axis_source'
        : 'not_built',
    real_axis_annotation_count: gates.real_axis_annotation_count ?? 0,
    real_axis_annotation_count_total: gates.real_axis_annotation_count_total ?? 0,
    real_event_seen: Boolean(gates.real_event_seen),
    ws_event_count: gates.ws_event_count ?? evidence.ws_event_count ?? 0,
    ws_timeline_started_count: gates.ws_timeline_started_count ?? evidence.ws_timeline_started_count ?? 0,
    event_audit_status: gates.event_audit_status ?? eventAudit?.status ?? null,
    event_audit_root_cause: gates.event_audit_root_cause ?? eventAudit?.root_cause ?? null,
    event_audit_next_action: eventAudit?.next_action ?? null,
    device_stream_status: deviceStream?.status ?? null,
    device_stream_enabled: Boolean(deviceStream?.enabled || deviceStream?.timer_active),
    device_stream_count: deviceStream?.count ?? 0,
    device_stream_max_count: deviceStream?.max_count ?? null,
    last_broadcast_at: streamStatus.last_broadcast_at,
    last_broadcast_client_count: streamStatus.last_client_count,
    last_broadcast_real_axis_annotation_count: streamStatus.last_state_summary?.real_axis_annotation_count ?? 0,
    current_sse_clients: streamStatus.current_clients,
    blockers: result?.blockers ?? [],
    next_action: result?.next_action ?? null,
  };
}

async function realDemoProgressPayload(req) {
  const result = await realDemoStatusPayload(req);
  const streamStatus = publicStreamStatus();
  return {
    generated_at: new Date().toISOString(),
    observed: Boolean(result.complete),
    timed_out: false,
    status: result.status,
    summary: result.summary,
    result,
    stream_status: streamStatus,
    completion_evidence: realDemoCompletionEvidence({ observed: result.complete, result }, streamStatus),
  };
}

async function realDemoAcceptancePayload(req) {
  const progress = await realDemoProgressPayload(req);
  const evidence = progress.completion_evidence ?? {};
  const result = progress.result ?? {};
  const runbook = result.operator_runbook ?? {};
  const requirements = Array.isArray(evidence.requirements) ? evidence.requirements : [];
  const productRequirements = requirements.filter((item) => item.id !== 'event_meeting_axis_built');
  const missingProductRequirements = productRequirements.filter((item) => !item.ok);
  const eventRequirement = requirements.find((item) => item.id === 'event_meeting_axis_built') ?? null;
  const productComplete = missingProductRequirements.length === 0;
  const realAxisActive = Boolean(evidence.real_meeting_axis_active);
  const hasRealAnnotation = Number(evidence.real_axis_annotation_count ?? 0) > 0;
  const hasRealtimeBroadcast = Number(evidence.last_broadcast_real_axis_annotation_count ?? 0) > 0;
  const oauthScanNeedsReauth = Boolean(runbook.oauth_scan_required_if_event_delivery_fails);
  const recommendedCommand = productComplete
    ? null
    : oauthScanNeedsReauth
      ? 'npm run accept:real-meeting:auth'
      : 'npm run accept:real-meeting';
  const verdict = productComplete
    ? 'pass'
    : !runbook.can_open_meeting_now
      ? 'needs_setup'
      : realAxisActive && !hasRealAnnotation
        ? 'waiting_for_open_annotation'
        : realAxisActive && hasRealAnnotation && !hasRealtimeBroadcast
          ? 'waiting_for_realtime_broadcast'
          : 'ready_to_open_meeting';

  return {
    generated_at: progress.generated_at,
    verdict,
    product_acceptance_complete: productComplete,
    strict_event_acceptance_complete: Boolean(productComplete && evidence.event_axis_built),
    event_delivery_verified: Boolean(evidence.event_axis_built),
    event_delivery_required_for_product_acceptance: false,
    can_open_meeting_now: Boolean(runbook.can_open_meeting_now),
    recommended_command: recommendedCommand,
    recommended_auth_command: 'npm run accept:real-meeting:auth',
    recommended_plain_command: 'npm run accept:real-meeting',
    next_action: productComplete
      ? '真实会议轴、开放标注落轴和实时广播均已验证；文字转写可在会议结束后处理。'
      : result.next_action ?? evidence.next_action ?? '保持服务运行，直接开启飞书会议。',
    oauth_scan_needs_reauth: oauthScanNeedsReauth,
    scan_fallback_available: Boolean(result.gates?.scan_fallback_ready),
    annotation_endpoint: runbook.annotation_endpoint ?? localUrlFor(req, '/api/annotations'),
    progress_stream_url: runbook.progress_stream_url ?? localUrlFor(req, '/api/lark/real-demo/progress-stream'),
    meeting_axis: {
      active: realAxisActive,
      source: evidence.meeting_source,
      axis_creation_mode: evidence.axis_creation_mode,
      event_axis: Boolean(evidence.event_axis_built),
    },
    realtime_annotation: {
      real_axis_annotation_count: evidence.real_axis_annotation_count ?? 0,
      last_broadcast_real_axis_annotation_count: evidence.last_broadcast_real_axis_annotation_count ?? 0,
      last_broadcast_at: evidence.last_broadcast_at ?? null,
    },
    missing_product_requirements: missingProductRequirements.map((item) => ({
      id: item.id,
      required: item.required,
      observed: item.observed,
    })),
    strict_event_requirement: eventRequirement
      ? {
          ok: Boolean(eventRequirement.ok),
          required: eventRequirement.required,
          observed: eventRequirement.observed,
        }
      : null,
    blockers: evidence.blockers ?? [],
    evidence,
  };
}

async function publicDeliveryDiagnostics(req) {
  const config = larkConfigPayload(req);
  const state = await store.load();
  const auth = publicAuthState();
  const ws = config.ws_event_receiver ?? {};
  const registeredEventTypes = ws.registered_event_types ?? [];
  const requiredDirectEvents = [
    meetingEventRequirements.direct_meeting.start_event,
    meetingEventRequirements.direct_meeting.end_event,
  ];
  const fallbackContextEvents = [
    meetingEventRequirements.meeting_context.start_event,
    meetingEventRequirements.meeting_context.end_event,
  ];
  const missingRegisteredEvents = requiredDirectEvents.filter((eventType) => !registeredEventTypes.includes(eventType));
  const missingFallbackEvents = fallbackContextEvents.filter((eventType) => !registeredEventTypes.includes(eventType));
  const checklist = openPlatformChecklist(config, {
    direct_events: missingRegisteredEvents,
    fallback_events: missingFallbackEvents,
  });
  const realEvents = realDeliveredTimelineEvents(config);
  const wsEvents = larkEventLog.filter((event) => event.transport === 'ws_long_connection');
  const eventSummary = larkEventLogSummary(larkEventLog);
  const localHttpEvents = larkEventLog.filter((event) => event.transport === 'http_webhook' && !config.event_callback_public_https);
  const localParserVerified = localHttpEvents.some((event) => (
    event.timeline_processed
    && (isDirectMeetingStartEvent(event.event_type) || isMeetingContextBootstrapEvent(event.event_type))
  ));
  const directStartEvent = realEvents.find((event) => event.timeline_started && isDirectMeetingStartEvent(event.event_type)) ?? null;
  const reserveStartEvent = realEvents.find((event) => event.timeline_started && isReserveMeetingStartEvent(event.event_type)) ?? null;
  const contextStartEvent = realEvents.find((event) => event.timeline_started && isMeetingContextBootstrapEvent(event.event_type)) ?? null;
  const anyStartEvent = directStartEvent ?? reserveStartEvent ?? contextStartEvent;
  const anyRealMeetingEvent = realEvents.find((event) => event.timeline_processed) ?? null;
  const annotations = openAnnotationItems(state);
  const annotationBinding = annotationBindingSummary(state, {
    probe_started_at: realMeetingProbe.started_at,
  });
  const parserSelfTest = wsParserSelfTestPayload();
  const eventAudit = realMeetingEventAudit(config, { parser_self_test: parserSelfTest });

  const rawWsSummary = `recent_ws_event_count=${eventSummary.ws_event_count}, ws_timeline_candidate_count=${eventSummary.ws_timeline_candidate_count}, ws_timeline_started_count=${eventSummary.ws_timeline_started_count}`;
  let status = 'waiting_for_lark_delivery';
  let summary = `长连接已连接并注册直开会议事件，但尚未收到飞书云端投递的 ${meetingEventRequirements.direct_meeting.start_event}；${rawWsSummary}`;
  const nextActions = [];

  if (!config.configured) {
    status = 'missing_app_credentials';
    summary = '缺少飞书应用凭证，无法建立真实事件接收通道';
    nextActions.push('在 .env 配置 LARK_APP_ID / LARK_APP_SECRET 后重启服务');
  } else if (ws.state !== 'connected' && !config.event_callback_public_https) {
    status = 'receiver_not_ready';
    summary = `事件接收通道未就绪：长连接状态 ${ws.state}，HTTP webhook 不是公网 HTTPS`;
    nextActions.push('优先把开放平台事件订阅模式切到长连接，并确认本服务显示 connected');
    nextActions.push('如不用长连接，则用公网 HTTPS 配置 LARK_EVENT_CALLBACK_URL');
  } else if (missingRegisteredEvents.length) {
    status = 'missing_local_event_handlers';
    summary = `本地接收器未注册完整直开会议事件：缺少 ${missingRegisteredEvents.join(', ')}`;
    nextActions.push('检查 createWsEventDispatcher 的 registered_event_types');
  } else if (directStartEvent) {
    status = 'direct_meeting_event_received';
    summary = `已收到真实直开会议开始事件：${directStartEvent.event_type}；${rawWsSummary}`;
  } else if (reserveStartEvent) {
    status = 'meeting_start_event_received';
    summary = `已收到真实会议开始事件并建轴：${reserveStartEvent.event_type}；${rawWsSummary}`;
  } else if (contextStartEvent) {
    status = 'meeting_context_event_received';
    summary = `已收到真实会议上下文兜底事件并建轴：${contextStartEvent.event_type}；${rawWsSummary}`;
  } else if (anyRealMeetingEvent) {
    status = 'real_events_seen_but_no_direct_start';
    summary = `已收到真实飞书事件，但还没有直开会议开始事件；最近事件是 ${anyRealMeetingEvent.event_type}；${rawWsSummary}`;
    nextActions.push(`直接在飞书客户端开启会议，等待 ${meetingEventRequirements.direct_meeting.start_event}`);
  } else if (wsEvents.length) {
    status = 'ws_events_seen_but_no_meeting_start';
    summary = eventSummary.ws_timeline_candidate_count > 0
      ? `长连接已收到 ${eventSummary.ws_event_count} 条原始事件，其中 ${eventSummary.ws_timeline_candidate_count} 条像会议/妙记候选事件，但没有成功建轴；最近事件 ${eventSummary.last_ws_event_type}`
      : `长连接已收到 ${eventSummary.ws_event_count} 条原始事件，但都不是会议建轴候选；最近事件 ${eventSummary.last_ws_event_type}`;
    nextActions.push(eventSummary.ws_timeline_candidate_count > 0
      ? '检查最近候选事件的 payload 字段，确认 meeting_id/start_time 是否在当前解析器支持的路径里'
      : '检查开放平台订阅的事件类型是否包含企业会议开始/结束，而不是只投递了非会议事件');
  } else {
    nextActions.push('在开放平台确认事件接收方式是长连接，且当前应用版本已发布到测试企业');
    nextActions.push(`订阅 ${requiredDirectEvents.join(' / ')}`);
    nextActions.push(`兜底可同时订阅 ${fallbackContextEvents.join(' / ')}，首次加入会议事件也能建轴`);
    nextActions.push(`开通应用身份权限 ${meetingEventRequirements.direct_meeting.required_scope} 并发布到当前测试企业`);
    nextActions.push('确认应用可见范围包含当前登录并开会的用户');
    nextActions.push('保持服务运行后，直接在飞书客户端开启一次会议；页面 probe 只是辅助验收窗口，不负责建轴');
  }

  return {
    status,
    summary,
    generated_at: new Date().toISOString(),
    app: {
      configured: config.configured,
      app_id: process.env.LARK_APP_ID ?? null,
      base_url: config.base_url,
    },
    parser_self_test: parserSelfTest,
    real_meeting_event_audit: eventAudit,
    receiver: {
      ws_state: ws.state,
      ws_enabled: ws.enabled,
      registered_event_types: registeredEventTypes,
      required_direct_events: requiredDirectEvents,
      fallback_context_events: fallbackContextEvents,
      missing_registered_events: missingRegisteredEvents,
      missing_fallback_events: missingFallbackEvents,
      http_callback_url: config.event_callback_url,
      http_callback_public_https: config.event_callback_public_https,
    },
    open_platform_checklist: checklist,
    active_search: {
      status: auth.meeting_search.usable ? 'ready' : auth.meeting_search.reason,
      checked_at: capabilityStatus.meeting_search.checked_at,
      required_scope: capabilityStatus.meeting_search.required_scope,
      has_oauth_scope: hasOAuthScope('vc:meeting.search:read'),
      auth_mode: 'user_oauth',
      oauth: auth.meeting_search,
      api_status: capabilityStatus.meeting_search.status,
      tenant_search_supported: capabilityStatus.meeting_search.status !== 'invalid_token',
      error: capabilityStatus.meeting_search.error,
      permission_url: capabilityStatus.meeting_search.permission_url ?? permissionUrlForScopes('vc:meeting.search:read'),
      next_action: auth.meeting_search.next_action
        ?? (capabilityStatus.meeting_search.status === 'invalid_token'
        ? '应用/租户 access token 当前不能用于会议扫描；请重新授权当前用户 OAuth 后使用“扫描我的真实会议”或 probe 自动扫描。'
        : null),
      passive_scan: publicPassiveMeetingScanStatus(),
    },
    evidence: {
      real_event_count: realEvents.length,
      ws_event_count: wsEvents.length,
      ws_timeline_candidate_count: eventSummary.ws_timeline_candidate_count,
      ws_timeline_processed_count: eventSummary.ws_timeline_processed_count,
      ws_timeline_started_count: eventSummary.ws_timeline_started_count,
      ws_ignored_count: eventSummary.ws_ignored_count,
      local_http_event_count: localHttpEvents.length,
      local_parser_verified: localParserVerified,
      open_annotation_count: annotations.length,
      annotation_binding: annotationBinding,
      transcript_status: transcriptStatusForState(state),
      current_meeting_source: state.meeting?.source ?? null,
      current_meeting_pending: Boolean(state.meeting?.pending_binding),
      real_meeting_axis_active: isRealMeetingAxis(state.meeting),
      last_open_annotation: annotations[annotations.length - 1] ?? null,
      start_event: anyStartEvent,
      direct_start_event: directStartEvent,
      reserve_start_event: reserveStartEvent,
      context_start_event: contextStartEvent,
      last_real_event: realEvents[0] ?? null,
      last_ws_event: wsEvents[0] ?? null,
      last_ws_timeline_candidate_event: eventSummary.last_ws_timeline_candidate_event,
      last_ws_timeline_processed_event: eventSummary.last_ws_timeline_processed_event,
      last_ws_timeline_started_event: eventSummary.last_ws_timeline_started_event,
      last_ws_ignored_event: eventSummary.last_ws_ignored_event,
      last_local_http_event: localHttpEvents[0] ?? null,
    },
    root_cause: eventAudit.root_cause,
    audit_status: eventAudit.status,
    audit_next_action: eventAudit.next_action,
    required_scope: meetingEventRequirements.direct_meeting.required_scope,
    permission_url: permissionUrlForScopes(meetingEventRequirements.direct_meeting.required_scope),
    probe: publicRealMeetingProbeStatus(req),
    next_actions: nextActions,
  };
}

async function readinessPayload(req) {
  const state = await store.load();
  const config = larkConfigPayload(req);
  const ws = config.ws_event_receiver;
  const deliveredTimelineEvents = realDeliveredTimelineEvents(config);
  const hasRealTimelineEvent = deliveredTimelineEvents.length > 0;
  const lastRealTimelineEvent = deliveredTimelineEvents[0] ?? null;
  const directMeetingStarted = deliveredTimelineEvents.some((event) => event.timeline_started && isDirectMeetingStartEvent(event.event_type));
  const contextMeetingStarted = deliveredTimelineEvents.some((event) => event.timeline_started && isMeetingContextBootstrapEvent(event.event_type));
  const reserveMeetingStarted = deliveredTimelineEvents.some((event) => event.timeline_started && isReserveMeetingStartEvent(event.event_type));
  const auth = publicAuthState();
  const reserveReady = capabilityStatus.reserve.status === 'ok';
  const meetingLookupReady = capabilityStatus.meeting_lookup.status === 'ok';
  const meetingSearchReady = Boolean(auth.meeting_search.usable);
  const passiveScanStatus = publicPassiveMeetingScanStatus();
  const realMeetingEntryReady = directMeetingStarted || reserveMeetingStarted || contextMeetingStarted || isRealMeetingAxis(state.meeting);
  const annotations = openAnnotationItems(state);
  const annotationBinding = annotationBindingSummary(state);
  const transcriptStatus = transcriptStatusForState(state);
  const realAnnotationReady = isRealMeetingAxis(state.meeting) && annotationBinding.real_axis_count > 0;
  const checks = [
    {
      id: 'lark_app_configured',
      label: '飞书应用凭证',
      ok: Boolean(config.configured),
      detail: config.configured ? 'LARK_APP_ID / LARK_APP_SECRET 已配置' : '缺少 LARK_APP_ID / LARK_APP_SECRET',
    },
    {
      id: 'lark_user_oauth',
      label: '飞书用户 OAuth',
      ok: Boolean(config.oauth?.authenticated),
      detail: config.oauth?.authenticated
        ? '已登录飞书账号，可用于会后妙记搜索/同步'
        : config.oauth?.token_present && config.oauth?.expired
          ? '飞书 OAuth token 已过期，需要重新登录'
          : '未登录飞书账号，会后妙记同步不可用',
    },
    {
      id: 'lark_ws_connected',
      label: '飞书长连接',
      ok: ws.state === 'connected',
      detail: ws.state === 'connected' ? '长连接已连接' : `长连接状态：${ws.state}${ws.error ? `，${ws.error}` : ''}`,
    },
    {
      id: 'passive_real_meeting_listener',
      label: '真实会议被动监听',
      ok: ws.state === 'connected' || Boolean(config.event_callback_public_https),
      detail: ws.state === 'connected'
        ? '服务端长连接常驻监听；用户直接在飞书客户端开会时，真实会议事件会自动建轴，不需要先点页面按钮'
        : config.event_callback_public_https
          ? '已配置公网 HTTPS webhook；用户直接在飞书客户端开会时，真实会议事件会自动建轴'
          : '未连接长连接且没有公网 webhook，无法被动接收用户直接开会事件',
    },
    {
      id: 'direct_meeting_event_seen',
      label: '真实会议建轴事件',
      ok: directMeetingStarted || reserveMeetingStarted || contextMeetingStarted,
      detail: directMeetingStarted
        ? `已收到 ${meetingEventRequirements.direct_meeting.start_event}`
        : reserveMeetingStarted
          ? `已收到 ${meetingEventRequirements.reserve_meeting.start_event} 并创建真实会议轴`
          : contextMeetingStarted
            ? `已收到兜底事件 ${meetingEventRequirements.meeting_context.start_event} 并创建真实会议轴`
            : `尚未收到真实投递的 ${meetingEventRequirements.direct_meeting.start_event} / ${meetingEventRequirements.reserve_meeting.start_event}。可同时订阅 ${meetingEventRequirements.meeting_context.start_event} 作为建轴兜底；需要开通 ${meetingEventRequirements.direct_meeting.required_scope}`,
    },
    {
      id: 'real_lark_event_seen',
      label: '任意真实飞书会议事件',
      ok: hasRealTimelineEvent,
      detail: hasRealTimelineEvent
        ? `最近真实事件：${lastRealTimelineEvent.event_type} via ${lastRealTimelineEvent.transport}`
        : '尚未收到飞书真实投递的会议事件；本机 curl 或 localhost HTTP 回调只算模拟',
    },
    {
      id: 'reserve_permission',
      label: '预约会议 API 权限（可选）',
      ok: reserveReady,
      detail: reserveReady
        ? 'vc:reserve 可用，可由 demo 创建真实飞书会议并建轴'
        : capabilityStatus.reserve.status === 'missing_scope'
          ? '缺少 vc:reserve，无法由 demo 创建预约会议；这不影响“用户直接开会”主路径'
          : '尚未验证 vc:reserve；点击“创建飞书会议”会触发验证',
      action_url: reserveReady ? null : capabilityStatus.reserve.permission_url,
    },
    {
      id: 'meeting_lookup_permission',
      label: '会议号查询兜底权限',
      ok: meetingLookupReady,
      detail: meetingLookupReady
        ? '会议号/会议链接查询 API 可用，可作为事件订阅未打通时的手动兜底'
        : capabilityStatus.meeting_lookup.status === 'missing_scope'
          ? '缺少 vc:meeting.search:read 或 vc:meeting.meetingid:read，会议号兜底不可用'
          : '尚未验证会议号查询权限；点击“按会议号/链接绑定真实会议”会触发验证',
      action_url: meetingLookupReady ? null : publicCapabilityStatus().meeting_lookup.permission_url,
    },
    {
      id: 'meeting_search_permission',
      label: '当前用户会议扫描兜底',
      ok: meetingSearchReady,
      detail: meetingSearchReady
        ? capabilityStatus.meeting_search.status === 'ok'
          ? '当前用户会议搜索 API 可用，可在事件订阅未投递时主动扫描真实会议'
          : '当前 OAuth 已具备 vc:meeting.search:read；点击“扫描我的真实会议”可验证 API 是否可用'
        : auth.meeting_search.reason === 'oauth_token_expired'
          ? '飞书 OAuth token 已过期；重新授权 vc:meeting.search:read 后可自动扫描当前账号真实会议'
          : auth.meeting_search.reason === 'missing_scope'
          ? '当前 OAuth token 缺少 vc:meeting.search:read；需要重新登录获取新 scope'
          : '尚未登录飞书账号；自动扫描兜底需要当前账号 OAuth',
      action_url: meetingSearchReady ? null : publicCapabilityStatus().meeting_search.permission_url,
    },
    {
      id: 'passive_meeting_scan',
      label: '当前账号被动扫描',
      ok: Boolean(passiveScanStatus.enabled),
      detail: passiveScanStatus.enabled
        ? meetingSearchReady
          ? `已开启；服务端每 ${Math.round(passiveScanStatus.server_loop.interval_ms / 1000)} 秒扫描当前账号近期会议，发现进行中会议会自动建轴`
          : auth.meeting_search.reason === 'oauth_token_expired'
            ? '已开启，但飞书 OAuth 已过期；重新授权 vc:meeting.search:read 后才会扫描'
            : auth.meeting_search.reason === 'missing_scope'
              ? '已开启，但当前 OAuth 缺少 vc:meeting.search:read；重新授权后才会扫描'
              : '已开启，但需要先登录飞书账号并授权 vc:meeting.search:read'
        : '已关闭；服务端不会主动扫描当前账号会议，只等待飞书事件投递',
      action_url: meetingSearchReady ? null : publicCapabilityStatus().meeting_search.permission_url,
    },
    {
      id: 'real_meeting_entry',
      label: '真实会议建轴入口',
      ok: realMeetingEntryReady,
      detail: realMeetingEntryReady
        ? '已具备至少一条真实会议入口'
        : `未证明真实飞书会议入口；主路径需要 ${meetingEventRequirements.direct_meeting.start_event} 投递，兜底可用 ${meetingEventRequirements.meeting_context.start_event} 或 vc:reserve`,
    },
    {
      id: 'real_annotation_seen',
      label: '真实会议实时标注',
      ok: realAnnotationReady,
      detail: realAnnotationReady
        ? `当前真实会议轴已有 ${annotationBinding.real_axis_count} 条通过开放标注接口写入的实时标注`
        : realMeetingEntryReady
          ? '真实会议轴已建立，但还没有通过 POST /api/annotations 写入标注'
          : '等待真实会议轴建立后，再写入一条实时标注完成端到端验收',
    },
    {
      id: 'open_annotation_api',
      label: '开放标注接口',
      ok: true,
      detail: 'POST /api/annotations 可接收未来墨水屏手写/标注事件；/api/annotation-ingest-info 提供机器可读接入信息并支持 CORS/OPTIONS',
    },
    {
      id: 'realtime_timeline',
      label: '实时刷新',
      ok: true,
      detail: 'GET /api/stream SSE 会在标注、会议事件、转写更新时推送最新时间轴',
    },
    {
      id: 'pending_rebind',
      label: '标注先到兜底',
      ok: true,
      detail: '带 captured_at_ms 或 stroke 绝对时间的标注先到时会创建 pending 时间轴，真实会议开始事件后会重绑定；缺时间戳的标注只保存不建轴',
    },
    {
      id: 'post_meeting_transcript',
      label: '会后转写',
      ok: true,
      detail: transcriptStatus.next_action,
    },
  ];
  const blockers = checks.filter((check) => !check.ok && [
    'lark_app_configured',
    'real_meeting_entry',
    'real_annotation_seen',
  ].includes(check.id));
  if (!realMeetingEntryReady) {
    blockers.push({
      id: 'external_lark_setup',
      label: '飞书开放平台配置',
      ok: false,
      detail: `需要订阅 ${meetingEventRequirements.direct_meeting.start_event} / ${meetingEventRequirements.direct_meeting.end_event} 并开通 ${meetingEventRequirements.direct_meeting.required_scope}；vc:reserve 只是预约会议备选路径`,
      action_url: permissionUrlForScopes(meetingEventRequirements.direct_meeting.required_scope),
    });
  }
  return {
    ready: checks
      .filter((check) => ![
        'lark_user_oauth',
        'lark_ws_connected',
        'passive_real_meeting_listener',
        'direct_meeting_event_seen',
        'real_lark_event_seen',
        'reserve_permission',
        'meeting_lookup_permission',
        'meeting_search_permission',
        'passive_meeting_scan',
      ].includes(check.id))
      .every((check) => check.ok),
    ready_for_local_demo: checks
      .filter((check) => !['lark_user_oauth', 'passive_real_meeting_listener', 'direct_meeting_event_seen', 'real_lark_event_seen', 'reserve_permission', 'meeting_lookup_permission', 'meeting_search_permission', 'passive_meeting_scan', 'real_meeting_entry', 'real_annotation_seen'].includes(check.id))
      .every((check) => check.ok),
    generated_at: new Date().toISOString(),
    checks,
    blockers,
    current: {
      meeting: state.meeting,
      sequence_count: state.sequence?.length ?? 0,
      open_annotation_count: annotations.length,
      annotation_binding: annotationBinding,
      transcript_status: transcriptStatus,
      event_count: state.events?.length ?? 0,
      alignment_count: state.alignments?.length ?? 0,
    },
    real_meeting_probe: publicRealMeetingProbeStatus(req),
    auto_acceptance: publicAutoAcceptanceStatus(),
    device_simulator: publicDeviceSimulatorStatus(),
    passive_meeting_scan: passiveScanStatus,
    passive_scan: passiveScanStatus,
    real_demo_session: publicRealDemoSessionStatus(),
    capabilities: publicCapabilityStatus(),
  };
}

async function readJson(req) {
  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  if (!chunks.length) return {};
  const text = Buffer.concat(chunks).toString('utf8');
  if (!text.trim()) return {};
  return JSON.parse(text);
}

function safePublicPath(urlPath) {
  const pathname = decodeURIComponent(new URL(urlPath, 'http://localhost').pathname);
  const requested = pathname === '/' ? '/index.html' : pathname;
  const filePath = normalize(join(publicDir, requested));
  if (!filePath.startsWith(publicDir)) return null;
  return filePath;
}

function verifyLarkToken(payload) {
  const expected = process.env.LARK_VERIFICATION_TOKEN;
  if (!expected) return;
  const actual = payload?.token ?? payload?.header?.token;
  if (actual && actual !== expected) throw new Error('Lark verification token mismatch');
}

function localUrlFor(req, pathname) {
  const host = req.headers.host ?? `localhost:${port}`;
  const proto = req.headers['x-forwarded-proto'] ?? 'http';
  return new URL(pathname, `${proto}://${host}`).toString();
}

function publicWebhookStatus(callbackUrl) {
  let parsed = null;
  try {
    parsed = new URL(callbackUrl);
  } catch {
    return false;
  }
  const localHosts = new Set(['localhost', '127.0.0.1', '0.0.0.0', '::1']);
  return parsed.protocol === 'https:' && !localHosts.has(parsed.hostname);
}

function oauthAuthorizeDetails(authUrl) {
  try {
    const parsed = new URL(authUrl);
    const scopes = parsed.searchParams.get('scope')?.split(/\s+/).filter(Boolean) ?? [];
    return {
      authorize_origin: parsed.origin,
      redirect_uri: parsed.searchParams.get('redirect_uri') ?? null,
      scopes,
      scope_present: {
        'vc:meeting.search:read': scopes.includes('vc:meeting.search:read'),
      },
    };
  } catch {
    return {
      authorize_origin: null,
      redirect_uri: null,
      scopes: [],
      scope_present: {
        'vc:meeting.search:read': false,
      },
    };
  }
}

function oauthStartPayload(scopes = [], options = {}) {
  const state = randomUUID();
  const savedAuth = rememberOAuthState(state);
  const stateEntry = savedAuth.oauth_state_history?.find((entry) => entry.state === state);
  const requestedScope = Array.isArray(scopes) ? scopes.filter(Boolean).join(' ') : String(scopes ?? '');
  const authUrl = lark.createAuthorizeUrl(state, {
    scope: requestedScope,
    ignoreDefaultScopes: Boolean(options.ignoreDefaultScopes),
  });
  return {
    auth_url: authUrl,
    redirect_uri: lark.redirectUri,
    callback_url: lark.redirectUri,
    requested_scope: requestedScope || null,
    scope_mode: options.scopeMode ?? null,
    state,
    state_created_at: stateEntry?.created_at ?? savedAuth.updated_at ?? null,
    state_expires_at: stateEntry?.created_at_ms
      ? new Date(stateEntry.created_at_ms + oauthStateTtlMs).toISOString()
      : null,
    ...oauthAuthorizeDetails(authUrl),
  };
}

function oauthRedirectUrlFor(req, scopes = []) {
  const params = new URLSearchParams();
  const requestedScope = Array.isArray(scopes) ? scopes.filter(Boolean).join(' ') : String(scopes ?? '');
  if (requestedScope) params.set('scope', requestedScope);
  params.set('redirect', '1');
  return localUrlFor(req, `/api/auth/lark/start?${params.toString()}`);
}

function meetingSearchOAuthStart(req) {
  return {
    method: 'GET',
    scopes: ['vc:meeting.search:read'],
    redirect_url: oauthRedirectUrlFor(req, ['vc:meeting.search:read']),
    json_url: localUrlFor(req, '/api/auth/lark/start?scope=vc%3Ameeting.search%3Aread'),
    note: '打开 redirect_url 会跳转到飞书 OAuth；授权成功后被动扫描兜底可绑定当前账号正在进行的真实会议。',
  };
}

function shellSingleQuote(value) {
  return `'${String(value).replaceAll("'", "'\\''")}'`;
}

function scriptJson(value) {
  return JSON.stringify(value).replaceAll('</', '<\\/');
}

function larkConfigPayload(req) {
  const eventCallbackUrl = process.env.LARK_EVENT_CALLBACK_URL || localUrlFor(req, '/api/lark/events');
  const recentWsEvents = larkEventLog.filter((event) => event.transport === 'ws_long_connection');
  const eventSummary = larkEventLogSummary(larkEventLog);
  const payload = {
    base_url: lark.baseUrl,
    redirect_uri: lark.redirectUri,
    configured: lark.isConfigured,
    oauth: publicAuthState(),
    event_callback_url: eventCallbackUrl,
    event_callback_public_https: publicWebhookStatus(eventCallbackUrl),
    verification_token_configured: Boolean(process.env.LARK_VERIFICATION_TOKEN),
    ws_event_receiver: publicLarkWsStatus(),
    meeting_event_requirements: publicCapabilityStatus().meeting_events,
    event_log_summary: eventSummary,
    recent_event_count: larkEventLog.length,
    recent_ws_event_count: recentWsEvents.length,
    recent_events: larkEventLog.slice(0, 8),
    recent_ws_events: recentWsEvents.slice(0, 8),
    plaintext_event_callback: true,
    encrypted_event_callback: false,
  };
  return {
    ...payload,
    real_meeting_event_audit: realMeetingEventAudit(payload),
  };
}

async function annotationIngestInfoPayload(req) {
  const state = await store.load();
  const now = Date.now();
  const route = annotationIngestRoute(state);
  const endpoint = localUrlFor(req, '/api/annotations');
  const batchEndpoint = localUrlFor(req, '/api/annotations/batch');
  const meetingSessionStartEndpoint = localUrlFor(req, '/api/meeting-session/start');
  const meetingSessionEndEndpoint = localUrlFor(req, '/api/meeting-session/end');
  const annotationStatusUrl = `${localUrlFor(req, '/api/annotation-status')}?id={annotation_id}`;
  const timeSyncUrl = localUrlFor(req, '/api/time');
  const minimalPayload = {
    id: 'device-mark-001',
    source: 'hanwang_epaper',
    captured_at_ms: now,
    kind: 'handwriting_trigger',
    label: 'why?',
    text_candidates: ['why?', 'why'],
    intent: 'question',
    strokes: [],
    payload: {
      page_id: 'optional-page-id',
    },
  };
  const minimalBatchPayload = {
    annotations: [
      {
        id: 'device-mark-001',
        source: 'hanwang_epaper',
        captured_at_ms: now,
        kind: 'handwriting_trigger',
        label: 'why?',
        text_candidates: ['why?', 'why'],
        intent: 'question',
        strokes: [],
      },
      {
        id: 'device-mark-002',
        source: 'hanwang_epaper',
        captured_at_ms: now + 5000,
        kind: 'attention',
        label: '重点',
        text_candidates: ['重点'],
        intent: 'attention',
        strokes: [],
      },
    ],
  };
  const minimalPayloadWithMeetingSession = {
    ...minimalPayload,
    id: 'device-mark-with-session-001',
    meeting_session: {
      platform: 'lark',
      title: '真实飞书会议',
      meeting_url: 'https://vc.feishu.cn/j/example',
      start_time_ms: now,
      detector_source: 'hanwang_host_app',
    },
  };
  return {
    version: 1,
    generated_at: new Date(now).toISOString(),
    endpoint,
    batch_endpoint: batchEndpoint,
    method: 'POST',
    batch_method: 'POST',
    content_type: 'application/json; charset=utf-8',
    schema_url: localUrlFor(req, '/annotation-schema.json'),
    binding_lookup_url: `${localUrlFor(req, '/api/annotation-bindings')}?id={annotation_id}`,
    annotation_status_url: annotationStatusUrl,
    time_sync_url: timeSyncUrl,
    server_time_url: timeSyncUrl,
    time_url: timeSyncUrl,
    stream_url: localUrlFor(req, '/api/stream'),
    stream_status_url: localUrlFor(req, '/api/stream/status'),
    stream_status_alias_url: localUrlFor(req, '/api/stream-status'),
    meeting_session_start_endpoint: meetingSessionStartEndpoint,
    meeting_session_end_endpoint: meetingSessionEndEndpoint,
    meeting_session_inline_annotation: {
      supported: true,
      description: 'A single POST /api/annotations can start an open meeting session and append the mark when the payload includes meeting_session or start_meeting_session=true.',
      accepted_fields: [
        'meeting_session.platform',
        'meeting_session.meeting_id',
        'meeting_session.title',
        'meeting_session.meeting_url',
        'meeting_session.start_time_ms',
        'meeting_session.start_time',
        'meeting_session.detector_source',
        'start_meeting_session=true with top-level meeting_* fields',
      ],
      overwrite_rule: 'If a different real meeting axis is already active, inline meeting_session will not replace it unless force_meeting_session=true.',
    },
    real_demo_status_url: localUrlFor(req, '/api/lark/real-demo/status'),
    real_demo_progress_url: localUrlFor(req, '/api/lark/real-demo/progress'),
    real_demo_acceptance_url: localUrlFor(req, '/api/lark/real-demo/acceptance'),
    real_demo_progress_stream_url: localUrlFor(req, '/api/lark/real-demo/progress-stream'),
    real_demo_monitor_url: localUrlFor(req, '/api/lark/real-demo/monitor'),
    cors: {
      enabled: true,
      allow_origin: '*',
      allow_methods: ['GET', 'POST', 'OPTIONS'],
      allow_headers: ['content-type', 'authorization', 'x-device-id', 'x-hmp-device-id', 'x-device-type', 'x-hmp-device-type'],
    },
    timing: {
      preferred_field: 'captured_at_ms',
      accepted_absolute_fields: [
        'captured_at_ms',
        'captured_at',
        'timing.captured_at_ms',
        'payload.timing.captured_at_ms',
        'ink_end_at_ms',
        'ink_end_time_ms',
        'stroke_end_at_ms',
        'stroke_end_time_ms',
        'timestamp_ms',
        'timestamp',
        'created_at_ms',
        'device_time_ms',
      ],
      inferred_absolute_fields: [
        'strokes[*].t',
        'strokes[*].ts',
        'strokes[*].timestamp_ms',
        'stroke_points[*].t',
      ],
      relative_override_field: 'time_ms',
      server_time_ms: now,
    },
    clock_sync: {
      required: true,
      endpoint: timeSyncUrl,
      query_param: 'client_send_at_ms',
      recommended_algorithm: 'Before sending marks, call GET /api/time?client_send_at_ms={device_now_ms}; when the response returns, set midpoint_ms=(client_send_at_ms+client_receive_at_ms)/2 and clock_offset_ms=server_time_ms-midpoint_ms. Send captured_at_ms=device_mark_end_ms+clock_offset_ms.',
      max_recommended_skew_ms: 500,
      warning: 'If device clock is not corrected, marks uploaded after the meeting can appear after meeting_end and will not count as reliable realtime annotations.',
    },
    annotation_route: route,
    real_demo_session: publicRealDemoSessionStatus(),
    device_client_contract: {
      can_send_before_real_axis: true,
      can_send_during_meeting: true,
      can_send_after_meeting: true,
      stable_id_required_for_idempotency: true,
      clock_sync_required: true,
      preferred_time_field: 'captured_at_ms',
      retry_safe_when_id_stable: true,
      pending_ack_means: 'accepted, stored on a temporary pending timeline, and will be rebound when a real Lark meeting start event or passive scan binding arrives',
      real_axis_ack_means: 'accepted and normalized onto the current real meeting timeline',
      production_success_condition: 'ack.accepted === true && (ack.on_real_axis === true || ack.pending_real_meeting === true)',
      final_success_condition: 'ack.on_real_axis === true, or annotation_status_url later reports status=real_axis_bound; legacy /api/annotation-bindings?id={annotation_id} should then report on_real_axis=true',
      product_acceptance_condition: 'real_demo_acceptance_url reports product_acceptance_complete === true; this requires a real meeting axis, at least one open annotation on that axis, and a realtime SSE broadcast. strict_event_acceptance_complete additionally proves the axis came from a Lark meeting_start event instead of scan fallback.',
      status_polling_endpoint: annotationStatusUrl,
      realtime_observation: 'Subscribe to stream_url for state updates, or poll annotation_status_url for a specific annotation id.',
      meeting_session_start_condition: 'When the host app, desktop observer, or meeting detector knows the user has opened a real meeting, either call meeting_session_start_endpoint once or include meeting_session in the first annotation. Pass start_time_ms/start_time if the detector did not call exactly at meeting start.',
      should_alert_operator_when: [
        'ack.requires_device_captured_at === true',
        'ack.warnings includes pending_time_uses_server_receive_time',
        'ack.warnings includes unbound_time_uses_server_receive_time',
        'annotation status.status === needs_device_captured_at',
        'annotation status.status === after_meeting_end',
      ],
    },
    device_simulator: publicDeviceSimulatorStatus(),
    meeting_search_oauth: publicAuthState().meeting_search,
    response_ack: {
      object_path: 'ack',
      fields: {
        accepted: 'boolean',
        annotation_id: 'string',
        operation: 'created | updated',
        idempotent: 'boolean',
        idempotency_key_source: 'id | annotation_id | mark_id | event_id | null',
        warnings: 'string[]',
        binding_state: 'real_meeting_bound | pending_real_meeting | local_simulation | demo_ignored | unbound',
        normalized_time_ms: 'number',
        time_source: 'captured_at | stroke_point_time | explicit_time | server_received_at',
        timing_reliable: 'boolean; false means the server had to use receive time because the device did not send captured_at_ms/captured_at',
        requires_device_captured_at: 'boolean; true means production sender should add captured_at_ms or absolute stroke timestamps',
        after_meeting_end_ms: 'number; >0 means the normalized annotation lands after the known meeting end and must not count as a real-axis aligned mark',
        meeting_id: 'string|null',
        pending_binding: 'boolean',
        on_real_axis: 'boolean',
        created_pending_timeline: 'boolean',
        replaced_existing: 'boolean',
      },
      sibling_fields: {
        meeting_session_binding: 'object; when the annotation includes meeting_session/start_meeting_session, this reports whether an open meeting session was started before normalizing the mark',
        passive_binding: 'object; if passive current-user scan is enabled and no real axis exists, the server may bind a real meeting before normalizing this annotation',
      },
    },
    response_batch: {
      endpoint: batchEndpoint,
      accepted_shapes: [
        'MeetingAnnotationEvent[]',
        '{ "annotations": MeetingAnnotationEvent[] }',
        '{ "items": MeetingAnnotationEvent[] }',
      ],
      object_path: 'acks',
      max_items: 200,
      note: '批量上传会按数组顺序归一化，最终只广播一次 SSE state；每条标注仍返回独立 ACK。',
    },
    current_meeting: {
      meeting_id: route.current_meeting_visible ? state.meeting?.meeting_id ?? null : null,
      title: route.current_meeting_visible ? state.meeting?.title ?? null : null,
      source: route.current_meeting_visible ? state.meeting?.source ?? null : null,
      pending_binding: route.current_meeting_visible ? Boolean(state.meeting?.pending_binding) : false,
      start_time: route.current_meeting_visible ? state.meeting?.start_time ?? null : null,
      end_time: route.current_meeting_visible ? state.meeting?.end_time ?? null : null,
    },
    ignored_current_meeting: route.current_meeting_visible ? null : {
      meeting_id: state.meeting?.meeting_id ?? null,
      title: state.meeting?.title ?? null,
      source: state.meeting?.source ?? null,
      reason: route.reason,
    },
    minimal_payload: minimalPayload,
    minimal_payload_with_meeting_session: minimalPayloadWithMeetingSession,
    minimal_batch_payload: minimalBatchPayload,
    curl_examples: {
      single: [
        `curl -sS -X POST ${shellSingleQuote(meetingSessionStartEndpoint)} \\`,
        `  -H ${shellSingleQuote('content-type: application/json; charset=utf-8')} \\`,
        `  -d ${shellSingleQuote(JSON.stringify({ platform: 'lark', title: '真实飞书会议', start_time_ms: now }))}`,
        '',
        `curl -sS -X POST ${shellSingleQuote(endpoint)} \\`,
        `  -H ${shellSingleQuote('content-type: application/json; charset=utf-8')} \\`,
        `  -H ${shellSingleQuote('x-hmp-device-id: hanwang-alpha-001')} \\`,
        `  -H ${shellSingleQuote('x-hmp-device-type: hanwang_epaper')} \\`,
        `  --data-binary ${shellSingleQuote(JSON.stringify(minimalPayload))}`,
      ].join('\n'),
      verify_status: `curl -sS ${shellSingleQuote(annotationStatusUrl.replace('{annotation_id}', encodeURIComponent(minimalPayload.id)))}`,
      watch_stream: `curl -N ${shellSingleQuote(localUrlFor(req, '/api/stream'))}`,
      full_roundtrip: [
        '# 1. 写入开放标注；真实会议轴未出现时会进入 pending，真实轴出现后自动回绑。',
        `ANNOTATION_ID=${shellSingleQuote(minimalPayload.id)}`,
        `curl -sS -X POST ${shellSingleQuote(endpoint)} \\`,
        `  -H ${shellSingleQuote('content-type: application/json; charset=utf-8')} \\`,
        `  -H ${shellSingleQuote('x-hmp-device-id: hanwang-alpha-001')} \\`,
        `  -H ${shellSingleQuote('x-hmp-device-type: hanwang_epaper')} \\`,
        `  --data-binary ${shellSingleQuote(JSON.stringify(minimalPayload))}`,
        '',
        '# 2. 查询单条标注状态：pending_real_meeting / real_axis_bound / needs_device_captured_at / after_meeting_end。',
        `curl -sS "${localUrlFor(req, '/api/annotation-status')}?id=\${ANNOTATION_ID}"`,
        '',
        '# 3. 观察实时 SSE；页面时间轴也是通过同一类 state 广播刷新。',
        `curl -N ${shellSingleQuote(localUrlFor(req, '/api/stream'))}`,
      ].join('\n'),
      batch: [
        `curl -sS -X POST ${shellSingleQuote(batchEndpoint)} \\`,
        `  -H ${shellSingleQuote('content-type: application/json; charset=utf-8')} \\`,
        `  -H ${shellSingleQuote('x-hmp-device-id: hanwang-alpha-001')} \\`,
        `  -H ${shellSingleQuote('x-hmp-device-type: hanwang_epaper')} \\`,
        `  --data-binary ${shellSingleQuote(JSON.stringify(minimalBatchPayload))}`,
      ].join('\n'),
    },
    notes: [
      '生产接入必须上报 captured_at_ms/captured_at 或 stroke 点内绝对时间；否则服务端只能按收到 POST 的时间落点，批量/延迟上传会把会中标注画到会议结束后，且不会计入可靠真实会议轴标注。',
      '设备端优先上报 captured_at_ms 这种绝对时间；服务端会用真实会议 start_time 换算成会议相对时间。',
      '如果只能上传 stroke 点，点内 t/ts/timestamp_ms 必须是 Unix 绝对时间；0, 12, 350 这类相对笔画时间不会被当作会议时间。',
      '如果带绝对采集时间的标注先于真实飞书会议事件到达，服务端会先建立 pending 时间轴，真实会议开始事件到达后重新绑定。',
      '只有当前账号被动扫描被显式开启且 OAuth 可用时，标注到达才会优先尝试绑定正在进行的真实飞书会议，再把标注落到真实会议轴。',
      'time_ms 只建议用于离线回放或人工调试，因为它会被视为已经换算好的会议相对时间。',
    ],
  };
}

function annotationIngestRoute(state = {}) {
  const meeting = state.meeting ?? {};
  const willCreatePending = shouldCreatePendingTimeline(state, { captured_at_ms: Date.now() });
  const pending = Boolean(meeting.pending_binding);
  const visible = pending || (!willCreatePending && Boolean(meeting.meeting_id));
  const mode = pending
    ? 'append_to_pending_meeting'
    : willCreatePending
      ? 'create_pending_on_first_annotation'
      : isRealMeetingAxis(meeting)
        ? 'append_to_real_meeting'
        : meeting.source === 'local_simulation'
          ? 'append_to_local_simulation'
          : 'append_to_current_meeting';
  return {
    mode,
    current_meeting_visible: visible,
    will_create_pending_on_first_annotation: willCreatePending && !pending,
    requires_captured_at_to_create_pending: willCreatePending && !pending,
    pending_binding: pending,
    real_meeting_axis_active: isRealMeetingAxis(meeting),
    reason: !visible
      ? meeting.meeting_id === 'demo-lark-meeting-001'
        ? 'demo_sample_axis_is_not_a_real_meeting'
        : 'annotation_will_wait_for_real_lark_event'
      : null,
  };
}

function annotationInputWithRequestMetadata(input = {}, req = {}) {
  const deviceId = req.headers?.['x-hmp-device-id'] ?? req.headers?.['x-device-id'];
  const deviceType = req.headers?.['x-hmp-device-type'] ?? req.headers?.['x-device-type'];
  if (!deviceId && !deviceType) return input;
  const next = { ...input };
  if (!next.device_id && !next.device?.id && deviceId) next.device_id = String(deviceId);
  if (!next.source && deviceType) next.source = String(deviceType);
  return next;
}

function annotationBatchInputsFromBody(body = {}, req = {}) {
  const rows = Array.isArray(body)
    ? body
    : body.annotations ?? body.items ?? body.events ?? body.marks;
  if (!Array.isArray(rows)) return null;
  return rows.map((item) => annotationInputWithRequestMetadata(item, req));
}

async function handleApi(req, res, url) {
  if (req.method === 'OPTIONS') {
    return sendNoContent(res);
  }

  if (req.method === 'GET' && url.pathname === '/api/state') {
    return sendJson(res, 200, publicTimelineState(await store.load()));
  }

  if (req.method === 'GET' && url.pathname === '/api/annotation-ingest-info') {
    return sendJson(res, 200, await annotationIngestInfoPayload(req));
  }

  if (req.method === 'GET' && url.pathname === '/api/meeting-session/status') {
    const state = await store.load();
    return sendJson(res, 200, {
      generated_at: new Date().toISOString(),
      ready: true,
      start_endpoint: localUrlFor(req, '/api/meeting-session/start'),
      end_endpoint: localUrlFor(req, '/api/meeting-session/end'),
      current_meeting: state.meeting ?? null,
      real_meeting_axis_active: isRealMeetingAxis(state.meeting),
      strict_lark_event_axis: ['lark_ws_event', 'lark_http_event'].includes(state.meeting?.source),
      contract: {
        source: 'open_meeting_session',
        note: 'Call start when a real meeting session starts; use annotation endpoints for live marks; import transcript after the meeting.',
      },
    });
  }

  if (req.method === 'POST' && url.pathname === '/api/meeting-session/start') {
    const body = await readJson(req);
    try {
      return sendJson(res, 200, await startOpenMeetingSession(body));
    } catch (error) {
      return sendJson(res, error.status ?? 400, {
        error: error.message ?? String(error),
        current_meeting: error.current_meeting ?? null,
      });
    }
  }

  if (req.method === 'POST' && url.pathname === '/api/meeting-session/end') {
    const body = await readJson(req);
    try {
      return sendJson(res, 200, await endOpenMeetingSession(body));
    } catch (error) {
      return sendJson(res, error.status ?? 400, {
        error: error.message ?? String(error),
        current_meeting: error.current_meeting ?? null,
      });
    }
  }

  if (req.method === 'GET' && url.pathname === '/api/annotation-bindings') {
    const state = await store.load();
    const probeStartedAt = url.searchParams.get('probe_started_at') ?? realMeetingProbe.started_at;
    const id = url.searchParams.get('id') ?? url.searchParams.get('annotation_id');
    return sendJson(res, 200, {
      generated_at: new Date().toISOString(),
      probe_started_at: probeStartedAt ?? null,
      ...annotationBindingSummary(state, { probe_started_at: probeStartedAt, id }),
    });
  }

  if (req.method === 'GET' && (url.pathname === '/api/annotation-status' || url.pathname === '/api/annotations/status')) {
    const id = url.searchParams.get('id') ?? url.searchParams.get('annotation_id');
    if (id == null || id === '') {
      return sendJson(res, 400, {
        error: 'id or annotation_id is required',
        generated_at: new Date().toISOString(),
      });
    }
    const state = await store.load();
    const probeStartedAt = url.searchParams.get('probe_started_at') ?? realMeetingProbe.started_at;
    return sendJson(res, 200, {
      generated_at: new Date().toISOString(),
      probe_started_at: probeStartedAt ?? null,
      stream_url: localUrlFor(req, '/api/stream'),
      binding_lookup_url: `${localUrlFor(req, '/api/annotation-bindings')}?id=${encodeURIComponent(id)}`,
      ...annotationStatusForState(state, { probe_started_at: probeStartedAt, id }),
    });
  }

  if (req.method === 'GET' && url.pathname === '/api/transcript-status') {
    return sendJson(res, 200, {
      generated_at: new Date().toISOString(),
      ...transcriptStatusForState(await store.load()),
    });
  }

  if (req.method === 'GET' && url.pathname === '/api/time') {
    const receivedAtMs = Date.now();
    const rawClientSendAt = url.searchParams.get('client_send_at_ms')
      ?? url.searchParams.get('client_time_ms')
      ?? url.searchParams.get('device_time_ms');
    const clientSendAtMs = rawClientSendAt != null && rawClientSendAt !== ''
      ? Number(rawClientSendAt)
      : null;
    const sentAtMs = Date.now();
    const offsetAtReceiveMs = Number.isFinite(clientSendAtMs)
      ? receivedAtMs - clientSendAtMs
      : null;
    return sendJson(res, 200, {
      server_time_ms: sentAtMs,
      server_time: new Date(sentAtMs).toISOString(),
      server_received_at_ms: receivedAtMs,
      server_sent_at_ms: sentAtMs,
      client_send_at_ms: Number.isFinite(clientSendAtMs) ? clientSendAtMs : null,
      estimated_offset_at_receive_ms: offsetAtReceiveMs,
      device_midpoint_formula: 'clock_offset_ms = server_time_ms - ((client_send_at_ms + client_receive_at_ms) / 2)',
      captured_at_formula: 'captured_at_ms = device_mark_end_ms + clock_offset_ms',
    });
  }

  if (req.method === 'GET' && (url.pathname === '/api/stream/status' || url.pathname === '/api/stream-status')) {
    return sendJson(res, 200, publicStreamStatus());
  }

  if (req.method === 'GET' && url.pathname === '/api/stream') {
    res.writeHead(200, corsHeaders({
      'content-type': 'text/event-stream; charset=utf-8',
      'cache-control': 'no-cache, no-transform',
      connection: 'keep-alive',
    }));
    streamClients.add(res);
    streamTelemetry.opened_count += 1;
    streamTelemetry.last_opened_at = new Date().toISOString();
    sendSse(res, 'state', publicTimelineState(await store.load()));
    const ping = setInterval(() => sendSse(res, 'ping', { t: Date.now() }), 25_000);
    req.on('close', () => {
      clearInterval(ping);
      streamClients.delete(res);
      streamTelemetry.closed_count += 1;
      streamTelemetry.last_closed_at = new Date().toISOString();
    });
    return;
  }

  if (req.method === 'GET' && url.pathname === '/api/lark/config') {
    return sendJson(res, 200, larkConfigPayload(req));
  }

  if (req.method === 'GET' && url.pathname === '/api/lark/meeting-event-requirements') {
    return sendJson(res, 200, meetingEventRequirements);
  }

  if (req.method === 'GET' && url.pathname === '/api/lark/delivery-diagnostics') {
    return sendJson(res, 200, await publicDeliveryDiagnostics(req));
  }

  if (req.method === 'GET' && url.pathname === '/api/lark/ws-parser-self-test') {
    return sendJson(res, 200, wsParserSelfTestPayload());
  }

  if (req.method === 'GET' && url.pathname === '/api/lark/acceptance-report') {
    return sendJson(res, 200, await acceptanceReportPayload(req));
  }

  if (req.method === 'GET' && url.pathname === '/api/lark/real-demo/status') {
    return sendJson(res, 200, await realDemoStatusPayload(req));
  }

  if (url.pathname === '/api/lark/real-demo/wait' && (req.method === 'GET' || req.method === 'POST')) {
    const body = req.method === 'POST' ? await readJson(req) : {};
    return sendJson(res, 200, await waitForRealDemoStatus(req, {
      timeout_ms: body.timeout_ms ?? url.searchParams.get('timeout_ms'),
      interval_ms: body.interval_ms ?? url.searchParams.get('interval_ms'),
    }));
  }

  if (url.pathname === '/api/lark/real-demo/monitor' && (req.method === 'GET' || req.method === 'POST')) {
    const body = req.method === 'POST' ? await readJson(req) : {};
    const observed = await waitForRealDemoStatus(req, {
      timeout_ms: body.timeout_ms ?? url.searchParams.get('timeout_ms') ?? 180_000,
      interval_ms: body.interval_ms ?? url.searchParams.get('interval_ms') ?? 500,
    });
    const streamStatus = publicStreamStatus();
    return sendJson(res, 200, {
      ...observed,
      stream_status: streamStatus,
      completion_evidence: realDemoCompletionEvidence(observed, streamStatus),
    });
  }

  if (url.pathname === '/api/lark/real-demo/progress' && req.method === 'GET') {
    return sendJson(res, 200, await realDemoProgressPayload(req));
  }

  if (url.pathname === '/api/lark/real-demo/acceptance' && req.method === 'GET') {
    return sendJson(res, 200, await realDemoAcceptancePayload(req));
  }

  if (url.pathname === '/api/lark/real-demo/progress-stream' && req.method === 'GET') {
    const intervalMs = Math.min(Math.max(Number(url.searchParams.get('interval_ms') ?? 1000), 500), 5000);
    res.writeHead(200, corsHeaders({
      'content-type': 'text/event-stream; charset=utf-8',
      'cache-control': 'no-cache, no-transform',
      connection: 'keep-alive',
    }));
    let closed = false;
    const sendProgress = async () => {
      if (closed) return;
      try {
        sendSse(res, 'progress', await realDemoProgressPayload(req));
      } catch {
        closed = true;
      }
    };
    await sendProgress();
    const progressTimer = setInterval(sendProgress, intervalMs);
    const pingTimer = setInterval(() => {
      if (!closed) sendSse(res, 'ping', { t: Date.now() });
    }, 25_000);
    req.on('close', () => {
      closed = true;
      clearInterval(progressTimer);
      clearInterval(pingTimer);
    });
    return;
  }

  if (req.method === 'POST' && url.pathname === '/api/lark/real-demo/prepare') {
    const body = await readJson(req);
    return sendJson(res, 200, await prepareRealDemoRuntime(body, req));
  }

  if (req.method === 'GET' && url.pathname === '/api/acceptance/auto-annotation') {
    return sendJson(res, 200, publicAutoAcceptanceStatus());
  }

  if (req.method === 'GET' && url.pathname === '/api/device-simulator') {
    return sendJson(res, 200, publicDeviceSimulatorStatus());
  }

  if (req.method === 'GET' && url.pathname === '/api/device-simulator/stream') {
    return sendJson(res, 200, publicDeviceStreamSimulatorStatus());
  }

  if (req.method === 'POST' && url.pathname === '/api/device-simulator/stream') {
    const body = await readJson(req);
    const action = String(body.action ?? '').toLowerCase();
    if (action === 'stop' || body.enabled === false) {
      return sendJson(res, 200, stopDeviceStreamSimulator('stopped'));
    }
    if (action === 'start' || body.enabled !== false) {
      return sendJson(res, 200, await startDeviceStreamSimulator(body));
    }
    return sendJson(res, 200, publicDeviceStreamSimulatorStatus());
  }

  if (req.method === 'POST' && url.pathname === '/api/device-simulator') {
    const body = await readJson(req);
    const next = { ...deviceSimulator };
    if (body.enabled != null) next.enabled = Boolean(body.enabled);
    if (body.label != null && String(body.label).trim()) next.label = String(body.label).trim();
    if (body.device_id != null && String(body.device_id).trim()) next.device_id = String(body.device_id).trim();
    if (body.device_type != null && String(body.device_type).trim()) next.device_type = String(body.device_type).trim();
    const saved = saveDeviceSimulatorState(next);
    let trigger = null;
    if (body.trigger_now) {
      const current = await store.load();
      trigger = await maybeAppendDeviceSimulatorAnnotation(current, 'manual_trigger');
    }
    return sendJson(res, 200, {
      device_simulator: publicDeviceSimulatorStatus(),
      saved,
      trigger: trigger ? {
        annotation: trigger.annotation,
        skipped_reason: trigger.skipped_reason,
      } : null,
    });
  }

  if (req.method === 'POST' && url.pathname === '/api/acceptance/auto-annotation') {
    const body = await readJson(req);
    const next = {
      ...autoAcceptance,
    };
    if (body.enabled != null) next.enabled = Boolean(body.enabled);
    if (body.label != null && String(body.label).trim()) next.label = String(body.label).trim();
    const saved = saveAutoAcceptanceState(next);
    let trigger = null;
    if (body.trigger_now) {
      const current = await store.load();
      trigger = await maybeAppendAutoAcceptanceAnnotation(current, 'manual_trigger');
    }
    return sendJson(res, 200, {
      auto_acceptance: publicAutoAcceptanceStatus(),
      saved,
      trigger: trigger ? {
        annotation: trigger.annotation,
        skipped_reason: trigger.skipped_reason,
      } : null,
    });
  }

  if (req.method === 'GET' && (
    url.pathname === '/api/lark/real-meeting-probe'
      || url.pathname === '/api/lark/real-meeting-probe/status'
  )) {
    return sendJson(res, 200, publicRealMeetingProbeStatus(req));
  }

  if (req.method === 'POST' && url.pathname === '/api/lark/real-meeting-probe/start') {
    const body = await readJson(req);
    saveRealMeetingProbeState({
      active: true,
      started_at: new Date().toISOString(),
      timeout_ms: Math.min(Math.max(Number(body.timeout_ms ?? 120_000), 10_000), 10 * 60_000),
      note: body.note ? String(body.note) : null,
      auto_search: {
        enabled: body.auto_search !== false,
        interval_ms: Math.min(Math.max(Number(body.auto_search_interval_ms ?? 5000), 1000), 60_000),
        last_attempt_at: null,
        last_result: null,
      },
    });
    scheduleProbeAutoSearchLoop();
    const resetResult = await resetTemporaryAxisForProbeIfRequested(body);
    return sendJson(res, 200, {
      ...publicRealMeetingProbeStatus(req),
      temporary_axis_reset: resetResult.reset,
      temporary_axis_reset_reason: resetResult.reason,
    });
  }

  if (req.method === 'POST' && url.pathname === '/api/lark/real-meeting-probe/auto-bind') {
    const body = await readJson(req);
    return sendJson(res, 200, await autoBindProbeMeeting(body));
  }

  if (req.method === 'POST' && url.pathname === '/api/lark/restore-latest-real-meeting-axis') {
    const body = await readJson(req);
    try {
      return sendJson(res, 200, await restoreLatestRealMeetingAxisFromEventLog(body));
    } catch (error) {
      return sendJson(res, error.status ?? 500, {
        error: error.message ?? String(error),
      });
    }
  }

  if (req.method === 'POST' && url.pathname === '/api/lark/real-meeting-probe/reset') {
    saveRealMeetingProbeState({ ...defaultRealMeetingProbe });
    stopProbeAutoSearchLoop();
    return sendJson(res, 200, publicRealMeetingProbeStatus(req));
  }

  if (req.method === 'GET' && url.pathname === '/api/lark/passive-meeting-scan') {
    return sendJson(res, 200, publicPassiveMeetingScanStatus());
  }

  if (req.method === 'POST' && url.pathname === '/api/lark/passive-meeting-scan') {
    const body = await readJson(req);
    const next = { ...passiveMeetingScan };
    if (body.enabled != null) next.enabled = Boolean(body.enabled);
    if (body.tenant_fallback_enabled != null) next.tenant_fallback_enabled = Boolean(body.tenant_fallback_enabled);
    if (body.interval_ms != null) next.interval_ms = Math.min(Math.max(Number(body.interval_ms), 5000), 5 * 60_000);
    if (body.lookback_seconds != null) next.lookback_seconds = Math.min(Math.max(Number(body.lookback_seconds), 60), 2 * 60 * 60);
    if (body.lookahead_seconds != null) next.lookahead_seconds = Math.min(Math.max(Number(body.lookahead_seconds), 0), 30 * 60);
    savePassiveMeetingScanState(next);
    if (passiveMeetingScan.enabled) schedulePassiveMeetingScanLoop(100);
    else stopPassiveMeetingScanLoop();
    const trigger = body.trigger_now
      ? await autoBindPassiveMeeting({ ...body, force: true })
      : null;
    return sendJson(res, 200, {
      passive_meeting_scan: publicPassiveMeetingScanStatus(),
      trigger,
    });
  }

  if (req.method === 'GET' && (url.pathname === '/api/readiness' || url.pathname === '/api/lark/readiness')) {
    return sendJson(res, 200, await readinessPayload(req));
  }

  if (req.method === 'GET' && url.pathname === '/api/lark/events-log') {
    return sendJson(res, 200, {
      count: larkEventLog.length,
      items: larkEventLog,
    });
  }

  if (req.method === 'GET' && url.pathname === '/api/auth/lark/status') {
    return sendJson(res, 200, publicAuthState());
  }

  if (req.method === 'POST' && url.pathname === '/api/auth/lark/logout') {
    saveAuthState({ oauth_state: null, token: null, user: null });
    return sendJson(res, 200, publicAuthState());
  }

  if (req.method === 'POST' && url.pathname === '/api/auth/lark/refresh') {
    try {
      const token = await ensureUserAccessToken();
      if (!token) {
        return sendJson(res, 400, {
          error: '当前授权没有可用 refresh_token，请重新登录飞书账号。',
          auth: publicAuthState(),
        });
      }
      return sendJson(res, 200, publicAuthState());
    } catch (error) {
      return sendJson(res, 400, {
        error: error.message ?? String(error),
        auth: publicAuthState(),
      });
    }
  }

  if (req.method === 'POST' && url.pathname === '/api/lark/search-minutes') {
    const body = await readJson(req);
    const userToken = await ensureUserAccessToken();
    if (!userToken) {
      return sendJson(res, 401, { error: '请先登录或重新登录飞书账号。' });
    }
    const raw = await lark.searchMinutesWithUserToken(userToken, {
      query: body.query,
      keyword: body.keyword,
      page_size: body.page_size ?? 20,
      page_token: body.page_token,
      start_time: body.start_time,
      end_time: body.end_time,
    });
    const data = raw?.data ?? raw;
    const items = data?.items ?? data?.minutes ?? data?.list ?? [];
    return sendJson(res, 200, {
      raw,
      items,
      item_count: Array.isArray(items) ? items.length : 0,
      page_token: data?.page_token ?? data?.next_page_token ?? null,
      has_more: data?.has_more ?? false,
    });
  }

  if (req.method === 'POST' && url.pathname === '/api/lark/create-reserve') {
    const body = await readJson(req);
    let raw;
    try {
      raw = await lark.createMeetingReserve({
        title: body.title,
        topic: body.topic,
        end_time: body.end_time,
        end_time_ms: body.end_time_ms,
        meeting_connect: body.meeting_connect,
        auto_record: body.auto_record,
        password: body.password,
        owner_id: body.owner_id,
      });
    } catch (error) {
      const message = error.message ?? String(error);
      if (message.includes('vc:reserve') || message.includes('Access denied')) {
        capabilityStatus.reserve.status = 'missing_scope';
        capabilityStatus.reserve.checked_at = new Date().toISOString();
        capabilityStatus.reserve.error = message;
      } else {
        capabilityStatus.reserve.status = 'error';
        capabilityStatus.reserve.checked_at = new Date().toISOString();
        capabilityStatus.reserve.error = message;
      }
      return sendJson(res, message.includes('vc:reserve') || message.includes('Access denied') ? 403 : 500, {
        error: message,
        required_scope: message.includes('vc:reserve') ? 'vc:reserve' : null,
        permission_url: message.includes('vc:reserve')
          ? `https://open.feishu.cn/app/${process.env.LARK_APP_ID}/auth?q=vc:reserve&op_from=openapi&token_type=tenant`
          : null,
      });
    }
    if (raw.code != null && raw.code !== 0) {
      return sendJson(res, 400, {
        error: raw.msg || raw.message || `Feishu reserve error ${raw.code}`,
        raw,
      });
    }
    capabilityStatus.reserve.status = 'ok';
    capabilityStatus.reserve.checked_at = new Date().toISOString();
    capabilityStatus.reserve.error = null;
    const current = await store.load();
    const reserve = raw?.data?.reserve ?? {};
    const now = new Date();
    const meeting = {
      platform: 'lark',
      meeting_id: reserve.id ?? `reserve-${now.getTime()}`,
      external_meeting_id: reserve.meeting_no ?? null,
      meeting_url: reserve.url ?? reserve.app_link ?? null,
      minute_token: null,
      title: body.title ?? body.topic ?? reserve.topic ?? '实时标注会议',
      start_time: body.start_time ?? now.toISOString(),
      end_time: null,
      timezone: body.timezone ?? 'Asia/Shanghai',
      pending_binding: true,
      source: 'lark_reserve_pending',
      reserve_id: reserve.id ?? null,
      reserve_meeting_no: reserve.meeting_no ?? null,
      app_link: reserve.app_link ?? null,
      live_link: reserve.live_link ?? null,
    };
    const next = buildTimeline({
      meeting,
      segments: [],
      events: [],
      sequence: current.meeting?.pending_binding || current.meeting?.source === 'local_simulation'
        ? current.sequence ?? []
        : [],
    });
    return sendJson(res, 200, {
      raw,
      reserve,
      state: await saveAndBroadcast(next, 'state'),
    });
  }

  if (req.method === 'POST' && url.pathname === '/api/lark/reserve-active-meeting') {
    const body = await readJson(req);
    const current = await store.load();
    const reserveId = body.reserve_id ?? current.meeting?.reserve_id ?? current.meeting?.meeting_id;
    let raw;
    try {
      raw = await lark.fetchReserveActiveMeeting(reserveId);
    } catch (error) {
      const message = error.message ?? String(error);
      if (message.includes('vc:reserve') || message.includes('Access denied')) {
        capabilityStatus.reserve.status = 'missing_scope';
        capabilityStatus.reserve.checked_at = new Date().toISOString();
        capabilityStatus.reserve.error = message;
      }
      return sendJson(res, message.includes('vc:reserve') || message.includes('Access denied') ? 403 : 500, {
        error: message,
        required_scope: message.includes('vc:reserve') ? 'vc:reserve' : null,
        permission_url: message.includes('vc:reserve')
          ? `https://open.feishu.cn/app/${process.env.LARK_APP_ID}/auth?q=vc:reserve&op_from=openapi&token_type=tenant`
          : null,
      });
    }
    if (raw.code != null && raw.code !== 0) {
      return sendJson(res, 400, {
        error: raw.msg || raw.message || `Feishu active meeting error ${raw.code}`,
        raw,
      });
    }
    capabilityStatus.reserve.status = 'ok';
    capabilityStatus.reserve.checked_at = new Date().toISOString();
    capabilityStatus.reserve.error = null;
    const active = raw?.data?.meeting ?? {};
    const meeting = {
      ...current.meeting,
      meeting_id: active.id ?? current.meeting.meeting_id,
      external_meeting_id: active.meeting_no ?? current.meeting.external_meeting_id,
      meeting_url: active.url ?? current.meeting.meeting_url,
      title: active.topic ?? current.meeting.title,
      start_time: active.start_time ? new Date(Number(active.start_time) * 1000).toISOString() : current.meeting.start_time,
      end_time: active.end_time ? new Date(Number(active.end_time) * 1000).toISOString() : current.meeting.end_time,
      source: 'lark_active_meeting_api',
    };
    const next = mergeTimelineWithRebasedAnnotations(current, { meeting });
    return sendJson(res, 200, {
      raw,
      meeting: active,
      state: await saveAndBroadcast(next, 'state'),
    });
  }

  if (req.method === 'POST' && url.pathname === '/api/lark/search-meetings-by-no') {
    const body = await readJson(req);
    try {
      return sendJson(res, 200, await searchMeetingsByNoInput(body));
    } catch (error) {
      const fallbackNo = extractMeetingNo(body.meeting_no ?? body.meeting_url ?? body.url ?? body.text);
      const response = meetingLookupErrorPayload(error, fallbackNo);
      return sendJson(res, response.status, response.payload);
    }
  }

  if (req.method === 'POST' && url.pathname === '/api/lark/search-my-meetings') {
    const body = await readJson(req);
    try {
      return sendJson(res, 200, await searchCurrentUserMeetings(body));
    } catch (error) {
      const response = meetingSearchErrorPayload(error);
      return sendJson(res, response.status, response.payload);
    }
  }

  if (req.method === 'POST' && url.pathname === '/api/lark/search-tenant-meetings') {
    const body = await readJson(req);
    try {
      return sendJson(res, 200, await searchTenantMeetings(body));
    } catch (error) {
      const response = meetingSearchErrorPayload(error);
      return sendJson(res, response.status, response.payload);
    }
  }

  if (req.method === 'POST' && url.pathname === '/api/lark/bind-my-latest-meeting') {
    const body = await readJson(req);
    try {
      return sendJson(res, 200, await bindCurrentUserMeeting(body));
    } catch (error) {
      const response = meetingSearchErrorPayload(error);
      return sendJson(res, response.status, response.payload);
    }
  }

  if (req.method === 'POST' && url.pathname === '/api/lark/bind-tenant-latest-meeting') {
    const body = await readJson(req);
    try {
      return sendJson(res, 200, await bindTenantMeeting(body));
    } catch (error) {
      const response = meetingSearchErrorPayload(error);
      return sendJson(res, response.status, response.payload);
    }
  }

  if (req.method === 'POST' && url.pathname === '/api/lark/bind-meeting-by-no') {
    const body = await readJson(req);
    let lookup;
    try {
      lookup = await searchMeetingsByNoInput(body);
    } catch (error) {
      const fallbackNo = extractMeetingNo(body.meeting_no ?? body.meeting_url ?? body.url ?? body.text);
      const response = meetingLookupErrorPayload(error, fallbackNo);
      return sendJson(res, response.status, response.payload);
    }
    const items = lookup.items ?? [];
    if (!items.length) {
      return sendJson(res, 404, {
        error: 'No meeting found by meeting_no in the requested time range',
        meeting_no: lookup.meeting_no,
      });
    }
    const current = await store.load();
    const record = body.meeting_id
      ? items.find((item) => String(item.id ?? item.meeting_id ?? item.meeting?.id) === String(body.meeting_id)) ?? items[0]
      : items[0];
    requireReliableMeetingStart(record, body, 'lark_meeting_lookup_api');
    const manualStart = meetingStartOverride(body);
    const meeting = meetingFromLarkRecord(record, {
      meeting_id: body.meeting_id ?? undefined,
      source: 'lark_meeting_lookup_api',
      meeting_url: body.meeting_url ?? body.url ?? undefined,
      start_time: manualStart,
      start_time_source: manualStart != null ? 'manual_start_override' : undefined,
      end_time: body.end_time ?? null,
      timezone: current.meeting?.timezone ?? 'Asia/Shanghai',
    });
    const carriedSequence = carriedSequenceForNewRealAxis(current, meeting);
    const next = buildTimeline({
      meeting,
      segments: current.segments ?? [],
      events: [{
        id: `evt-lookup-${meeting.meeting_id}-start`,
        time_ms: 0,
        type: 'meeting_start',
        label: '会议号查询建轴',
        source: 'lark_meeting_lookup_api',
        metadata: { raw_type: 'vc.meeting.list_by_no', meeting_no: lookup.meeting_no },
      }],
      sequence: carriedSequence,
    });
    return sendJson(res, 200, {
      meeting_no: lookup.meeting_no,
      auth_mode: lookup.auth_mode,
      selected: record,
      state: await saveAndBroadcast(next, 'state'),
    });
  }

  if (req.method === 'GET' && url.pathname === '/api/auth/lark/start') {
    const purpose = String(url.searchParams.get('purpose') ?? url.searchParams.get('scope_mode') ?? '').toLowerCase();
    const requestedScope = [
      url.searchParams.get('scope'),
      url.searchParams.get('scopes'),
      url.searchParams.get('extra_scope'),
      url.searchParams.get('extra_scopes'),
    ].filter(Boolean).join(' ');
    const useMinutesOnly = ['minute', 'minutes', 'transcript'].includes(purpose);
    const payload = useMinutesOnly
      ? oauthStartPayload(minuteOAuthScopes, { ignoreDefaultScopes: true, scopeMode: 'minutes' })
      : oauthStartPayload(requestedScope);
    if (['1', 'true', 'yes'].includes(String(url.searchParams.get('redirect') ?? '').toLowerCase())) {
      return sendRedirect(res, payload.auth_url);
    }
    return sendJson(res, 200, payload);
  }

  if (req.method === 'GET' && url.pathname === '/api/auth/lark/callback') {
    const error = url.searchParams.get('error');
    if (error) {
      return sendHtml(res, 400, `<h1>飞书授权失败</h1><pre>${error}</pre>`);
    }
    const code = url.searchParams.get('code');
    const state = url.searchParams.get('state');
    if (!code) return sendHtml(res, 400, '<h1>飞书授权失败</h1><p>缺少 code。</p>');
    if (!oauthStateMatches(state)) {
      return sendHtml(res, 400, '<h1>飞书授权失败</h1><p>state 不匹配。</p>');
    }
    const token = await lark.exchangeOAuthCode(code);
    const tokenWithTime = { ...token, obtained_at_ms: Date.now() };
    let user = null;
    try {
      user = await lark.fetchUserInfo(token.access_token);
    } catch (userError) {
      user = { error: userError.message };
    }
    saveAuthState({ oauth_state: null, oauth_state_history: [], token: tokenWithTime, user });
    if (hasOAuthScope('vc:meeting.search:read')) {
      savePassiveMeetingScanState({
        ...passiveMeetingScan,
        enabled: true,
        tenant_fallback_enabled: false,
        updated_at: new Date().toISOString(),
      });
    }
    const passiveScanResult = await runImmediatePassiveMeetingScan('oauth_callback_passive_scan');
    schedulePassiveMeetingScanLoop(1000);
    const authCompletePayload = {
      type: 'lark-auth-complete',
      passive_scan: publicPassiveScanResult(passiveScanResult),
    };
    return sendHtml(res, 200, `
      <!doctype html>
      <meta charset="utf-8">
      <title>飞书授权完成</title>
      <body style="font-family: system-ui, sans-serif; padding: 32px;">
        <h1>飞书授权完成</h1>
        <p>可以回到 demo 页面；如果当前正在开飞书会议，服务端会立即尝试扫描并绑定真实会议轴。</p>
        <script>
          if (window.opener) {
            window.opener.postMessage(${scriptJson(authCompletePayload)}, '*');
            setTimeout(() => window.close(), 1200);
          } else {
            setTimeout(() => window.location.replace('/'), 1200);
          }
        </script>
      </body>
    `);
  }

  if (req.method === 'POST' && url.pathname === '/api/demo/reset') {
    saveRealDemoSessionState({ ...defaultRealDemoSession });
    const state = await store.resetDemo();
    broadcastState(state, 'state');
    return sendJson(res, 200, state);
  }

  if (req.method === 'POST' && url.pathname === '/api/live/start-meeting') {
    const body = await readJson(req);
    const current = await store.load();
    if (realDemoSession.active && !body.force) {
      return sendJson(res, 409, {
        error: '真实演示模式正在等待飞书真实会议开始事件；本地模拟按钮不会创建时间轴。需要调试时传 force=true。',
        meeting_source: current.meeting?.source ?? null,
        real_demo_session: publicRealDemoSessionStatus(),
      });
    }
    if (hasActiveMeeting(current) && !isLocalSimulationMeeting(current.meeting) && !body.force) {
      return sendJson(res, 409, {
        error: '当前已经是飞书真实会议轴；本地模拟开始不会覆盖真实轴。需要调试时传 force=true。',
        meeting_source: current.meeting?.source ?? null,
      });
    }
    const now = new Date();
    const meeting = {
      platform: 'lark',
      meeting_id: body.meeting_id ?? `live-${now.getTime()}`,
      external_meeting_id: body.external_meeting_id ?? body.lark_meeting_id ?? null,
      meeting_url: body.meeting_url ?? body.url ?? null,
      minute_token: body.minute_token ?? null,
      title: body.title ?? '实时标注会议',
      start_time: body.start_time ?? now.toISOString(),
      end_time: null,
      timezone: body.timezone ?? 'Asia/Shanghai',
      source: 'local_simulation',
    };
    const carriedSequence = current.meeting?.pending_binding
      ? rebasePendingSequence(current.sequence ?? [], meeting)
      : [];
    const next = buildTimeline({
      meeting,
      segments: [],
      events: [{
        id: 'evt-live-meeting-start',
        time_ms: 0,
        type: 'meeting_start',
        label: '本地模拟开始',
        source: 'local_simulation',
        metadata: { raw_type: 'local.live_meeting.started' },
      }],
      sequence: carriedSequence,
    });
    return sendJson(res, 200, await saveAndBroadcast(next, 'state'));
  }

  if (req.method === 'POST' && url.pathname === '/api/live/end-meeting') {
    const body = await readJson(req);
    const current = await store.load();
    if (realDemoSession.active && !body.force) {
      return sendJson(res, 409, {
        error: '真实演示模式正在等待飞书真实会议事件；本地模拟结束不会写入时间轴。需要调试时传 force=true。',
        meeting_source: current.meeting?.source ?? null,
        real_demo_session: publicRealDemoSessionStatus(),
      });
    }
    if (!isLocalSimulationMeeting(current.meeting) && !body.force) {
      return sendJson(res, 409, {
        error: '当前不是本地模拟会议轴；本地模拟结束不会写入真实飞书轴。需要调试时传 force=true。',
        meeting_source: current.meeting?.source ?? null,
      });
    }
    const meetingStart = Date.parse(current.meeting?.start_time ?? '');
    const now = new Date();
    const explicitTimeMs = body.time_ms == null ? null : Number(body.time_ms);
    const clockTimeMs = Number.isFinite(meetingStart) ? Math.max(0, now.getTime() - meetingStart) : 0;
    const timeMs = Number.isFinite(explicitTimeMs)
      ? Math.max(0, Math.round(explicitTimeMs))
      : clockTimeMs;
    const eventMap = new Map((current.events ?? []).map((x) => [x.id, x]));
    eventMap.set('evt-live-meeting-end', {
      id: 'evt-live-meeting-end',
      time_ms: timeMs,
      type: 'meeting_end',
      label: '本地模拟结束',
      source: 'local_simulation',
      metadata: { raw_type: 'local.live_meeting.ended' },
    });
    const next = mergeTimelineWithRebasedAnnotations(current, {
      meeting: { ...current.meeting, end_time: body.end_time ?? isoAtMeetingOffset(current.meeting, timeMs, now) },
      events: [...eventMap.values()],
    });
    return sendJson(res, 200, await saveAndBroadcast(next, 'state'));
  }

  if (req.method === 'POST' && url.pathname === '/api/import/sequence') {
    const body = await readJson(req);
    const current = await store.load();
    const next = mergeTimeline(current, { sequence: body.sequence ?? body.items ?? body });
    return sendJson(res, 200, await saveAndBroadcast(next, 'state'));
  }

  if (req.method === 'POST' && url.pathname === '/api/annotations') {
    const body = annotationInputWithRequestMetadata(await readJson(req), req);
    return sendJson(res, 200, await appendAnnotation(body));
  }

  if (req.method === 'POST' && url.pathname === '/api/annotations/batch') {
    const body = await readJson(req);
    const rows = annotationBatchInputsFromBody(body, req);
    try {
      return sendJson(res, 200, await appendAnnotationBatch(rows));
    } catch (error) {
      return sendJson(res, error.status ?? 400, { error: error.message ?? String(error) });
    }
  }

  if (req.method === 'POST' && url.pathname === '/api/live/sequence-event') {
    const body = await readJson(req);
    return sendJson(res, 200, await appendAnnotation(body));
  }

  if (req.method === 'POST' && url.pathname === '/api/import/lark-transcript') {
    const body = await readJson(req);
    const current = await store.load();
    const meeting = { ...current.meeting, ...(body.meeting ?? {}) };
    const segments = normalizeTranscript(body.transcript ?? body, meeting);
    const next = mergeTimelineWithRebasedAnnotations(current, { meeting, segments });
    return sendJson(res, 200, await saveAndBroadcast(next, 'state'));
  }

  if (req.method === 'POST' && url.pathname === '/api/lark/sync-minute') {
    const body = await readJson(req);
    const minuteToken = extractMinuteToken(body.minute_token ?? body.minuteToken ?? body.minute_url ?? body.minuteUrl);
    if (!minuteToken) return sendJson(res, 400, { error: 'minute_token is required' });
    const mode = body.mode ?? 'auto';
    const userToken = await ensureUserAccessToken();
    const useUserToken = mode !== 'tenant' && userToken;
    const transcriptRaw = useUserToken
      ? await lark.fetchMinuteTranscriptWithUserToken(minuteToken, userToken)
      : await lark.fetchMinuteTranscript(minuteToken);
    const current = await store.load();
    const meeting = {
      ...current.meeting,
      platform: 'lark',
      meeting_id: body.meeting_id ?? current.meeting.meeting_id,
      minute_token: minuteToken,
      title: body.title ?? current.meeting.title,
      start_time: body.start_time ?? current.meeting.start_time,
      end_time: body.end_time ?? current.meeting.end_time,
    };
    const segments = normalizeTranscript(transcriptRaw, meeting);
    const next = mergeTimelineWithRebasedAnnotations(current, { meeting, segments });
    return sendJson(res, 200, {
      ...(await saveAndBroadcast(next, 'state')),
      sync: {
        minute_token: minuteToken,
        auth_mode: useUserToken ? 'user_oauth' : 'tenant_access_token',
        raw_shape: Object.keys(transcriptRaw ?? {}),
        segment_count: segments.length,
      },
    });
  }

  if (req.method === 'POST' && url.pathname === '/api/lark/events') {
    const payload = await readJson(req);
    if (payload?.challenge) return sendJson(res, 200, { challenge: payload.challenge });
    if (payload?.type === 'url_verification' && payload?.challenge) {
      return sendJson(res, 200, { challenge: payload.challenge });
    }
    if (payload?.encrypt) {
      return sendJson(res, 400, {
        error: 'Encrypted callback is not enabled in this PoC. Disable callback encryption for first validation.',
      });
    }
    verifyLarkToken(payload);
    const eventType = eventTypeFromParsed(payload);
    const parsed = eventPayloadFromParsed(payload);
    const candidate = timelineEventCandidate(eventType, payload);
    const logEntry = pushLarkEventLog({
      transport: 'http_webhook',
      event_type: eventType,
      timeline_candidate: candidate,
      parsed,
      preview: payload,
    });
    const eventCallbackUrl = process.env.LARK_EVENT_CALLBACK_URL || localUrlFor(req, '/api/lark/events');
    const result = await processLarkEventPayload(payload, {
      source: publicWebhookStatus(eventCallbackUrl) ? 'lark_http_event' : 'lark_http_local_event',
    });
    logEntry.timeline_processed = result.ok !== false;
    logEntry.timeline_started = result.timeline_started;
    if (result.ignored_reason) logEntry.ignored_reason = result.ignored_reason;
    persistLarkEventLog();
    return sendJson(res, 200, result);
  }

  return sendJson(res, 404, { error: 'API route not found' });
}

const server = createServer(async (req, res) => {
  try {
    const url = new URL(req.url ?? '/', `http://${req.headers.host ?? 'localhost'}`);
    if (url.pathname.startsWith('/api/')) return await handleApi(req, res, url);

    const filePath = safePublicPath(req.url ?? '/');
    if (!filePath || !existsSync(filePath)) return sendText(res, 404, 'Not found');
    const type = contentTypes[extname(filePath)] ?? 'application/octet-stream';
    res.writeHead(200, corsHeaders({ 'content-type': type, 'cache-control': 'no-cache' }));
    createReadStream(filePath).pipe(res);
  } catch (error) {
    console.error(error);
    sendJson(res, 500, { error: error.message ?? String(error) });
  }
});

server.listen(port, () => {
  console.log(`Lark timeline demo: http://localhost:${port}`);
  console.log(`Lark configured: ${lark.isConfigured ? 'yes' : 'no'}`);
  startLarkWsReceiver();
  scheduleProbeAutoSearchLoop();
  schedulePassiveMeetingScanLoop(1000);
  maybeAutoArmRealDemoOnStartup()
    .then((result) => {
      if (result.armed) {
        console.log(`Real demo auto-armed: ${result.reason}`);
        return null;
      }
      return resumeRealDemoDeviceStreamOnStartup()
        .then((resumeResult) => {
          if (resumeResult.resumed) console.log(`Real demo device stream resumed: ${resumeResult.reason}`);
          return resumeResult;
        });
    })
    .catch((error) => {
      console.error('Real demo startup arm/resume failed:', error);
    });
});

function shutdown() {
  stopProbeAutoSearchLoop();
  stopPassiveMeetingScanLoop();
  stopDeviceStreamSimulator('server_shutdown');
  try {
    larkWsClient?.close?.({ force: true });
  } catch {
    // best effort shutdown
  }
}

process.once('SIGINT', () => {
  shutdown();
  process.exit(0);
});
process.once('SIGTERM', () => {
  shutdown();
  process.exit(0);
});
