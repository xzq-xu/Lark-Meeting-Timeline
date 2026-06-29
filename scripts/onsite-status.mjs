#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname } from 'node:path';

const args = new Map();
for (const raw of process.argv.slice(2)) {
  const [key, ...rest] = raw.replace(/^--/, '').split('=');
  args.set(key, rest.length ? rest.join('=') : 'true');
}

const baseUrl = String(args.get('url') || process.env.REAL_DEMO_BASE_URL || 'http://localhost:8787').replace(/\/$/, '');
const json = args.get('json') === 'true';
const reportFile = String(args.get('report-file') || args.get('write-report') || process.env.ONSITE_STATUS_REPORT_FILE || '');
const requirePracticalReady = args.get('require-practical-ready') === 'true';

async function fetchJson(path) {
  const response = await fetch(`${baseUrl}${path}`);
  const text = await response.text();
  let body = null;
  try {
    body = text ? JSON.parse(text) : null;
  } catch {
    body = text;
  }
  if (!response.ok) {
    throw new Error(`GET ${path} HTTP ${response.status}: ${typeof body === 'string' ? body : JSON.stringify(body)}`);
  }
  return body;
}

function bool(value) {
  return value ? 'yes' : 'no';
}

function compactText(value, max = 180) {
  const text = String(value ?? '').replace(/\s+/g, ' ').trim();
  return text.length > max ? `${text.slice(0, max - 1)}...` : text;
}

function summarize({ demoStatus, diagnostics, passiveScan, authStatus, streamStatus, ingestInfo }) {
  const gates = demoStatus.gates ?? {};
  const evidence = diagnostics.evidence ?? {};
  const audit = diagnostics.real_meeting_event_audit ?? {};
  const auth = authStatus.meeting_search ?? {};
  const eventPathReady = Boolean(gates.event_receiver_ready && audit.local_handlers_ready !== false && audit.parser_self_test_passed !== false);
  const eventDeliverySeen = Number(evidence.real_event_count ?? 0) > 0
    || Number(evidence.ws_event_count ?? 0) > 0
    || Number(evidence.public_webhook_event_count ?? 0) > 0;
  const scanFallbackEnabled = Boolean(passiveScan.enabled);
  const scanFallbackReady = Boolean(scanFallbackEnabled && auth.usable);
  const openAnnotationReady = Boolean(gates.open_annotation_ready || ingestInfo.endpoint);
  const realtimeReady = Boolean(gates.realtime_stream_ready || typeof streamStatus.broadcast_count === 'number');
  const deviceStreamReady = Boolean(gates.device_stream_enabled || demoStatus.device_stream_simulator?.enabled || demoStatus.device_stream_simulator?.timer_active);
  const realAxisActive = Boolean(gates.real_meeting_axis_active);
  const realAxisAnnotationCount = Number(gates.real_axis_annotation_count ?? 0);
  const productComplete = Boolean(demoStatus.complete || (realAxisActive && realAxisAnnotationCount > 0));
  const minimumStartReady = Boolean(openAnnotationReady && realtimeReady && deviceStreamReady && (eventPathReady || scanFallbackReady));
  const practicalStartReady = Boolean(openAnnotationReady && realtimeReady && deviceStreamReady && (scanFallbackReady || eventDeliverySeen));
  const status = productComplete
    ? 'complete'
    : practicalStartReady
      ? 'ready_practical_demo'
      : minimumStartReady
        ? 'ready_but_event_delivery_unproven'
        : 'needs_setup';
  const recommendedAction = productComplete
    ? '真实会议轴和实时标注已完成；转写可会后处理。'
    : scanFallbackReady
      ? '直接开启飞书会议；如事件未投递，当前用户会议扫描兜底会自动绑定真实会议轴。'
      : auth.needs_reauth
        ? '先完成 vc:meeting.search:read 重新授权；授权回调后服务端会自动开启当前用户会议扫描并立即尝试绑定。'
        : eventPathReady
          ? '可以直接开会验证飞书事件投递；若开会后事件仍为 0，请检查开放平台事件订阅/发布/可见范围。'
          : demoStatus.next_action ?? '先修复真实演示必要条件。';
  const ok = requirePracticalReady ? practicalStartReady || productComplete : minimumStartReady || productComplete;
  return {
    type: 'onsite_status',
    ok,
    status,
    base_url: baseUrl,
    generated_at: new Date().toISOString(),
    recommended_action: recommendedAction,
    product_complete: productComplete,
    minimum_start_ready: minimumStartReady,
    practical_start_ready: practicalStartReady,
    gates: {
      event_path_ready: eventPathReady,
      event_delivery_seen: eventDeliverySeen,
      scan_fallback_enabled: scanFallbackEnabled,
      scan_fallback_ready: scanFallbackReady,
      open_annotation_ready: openAnnotationReady,
      realtime_stream_ready: realtimeReady,
      device_stream_ready: deviceStreamReady,
      real_axis_active: realAxisActive,
      real_axis_annotation_count: realAxisAnnotationCount,
    },
    auth: {
      meeting_search_usable: Boolean(auth.usable),
      reason: auth.reason ?? null,
      needs_reauth: Boolean(auth.needs_reauth),
      required_scope: auth.required_scope ?? 'vc:meeting.search:read',
      scope_present: Boolean(auth.scope_present),
    },
    passive_scan: {
      enabled: Boolean(passiveScan.enabled),
      tenant_fallback_enabled: Boolean(passiveScan.tenant_fallback_enabled),
      server_loop: passiveScan.server_loop ?? null,
      last_result: passiveScan.last_result ?? null,
    },
    event_delivery: {
      status: diagnostics.status ?? null,
      audit_status: audit.status ?? null,
      ws_state: diagnostics.receiver?.ws_state ?? null,
      ws_event_count: Number(evidence.ws_event_count ?? 0),
      real_event_count: Number(evidence.real_event_count ?? 0),
      ws_timeline_started_count: Number(evidence.ws_timeline_started_count ?? 0),
      failed_open_platform_checks: (audit.required_open_platform_checks ?? [])
        .filter((check) => !check.ok)
        .map((check) => check.id),
    },
    endpoints: {
      annotation: ingestInfo.endpoint ?? null,
      batch: ingestInfo.batch_endpoint ?? null,
      stream: ingestInfo.stream_url ?? null,
      acceptance: ingestInfo.real_demo_acceptance_url ?? null,
    },
  };
}

function reportPayload(summary, { exitCode, error = null } = {}) {
  return {
    ...(summary ?? {
      type: 'onsite_status',
      ok: false,
      status: 'error',
      base_url: baseUrl,
      generated_at: new Date().toISOString(),
    }),
    report: {
      written_at: new Date().toISOString(),
      file: reportFile || null,
      command: 'scripts/onsite-status.mjs',
      exit_code: exitCode,
      error: error ? String(error.message ?? error) : null,
    },
  };
}

async function writeReport(summary, options) {
  if (!reportFile) return;
  await mkdir(dirname(reportFile), { recursive: true });
  await writeFile(reportFile, `${JSON.stringify(reportPayload(summary, options), null, 2)}\n`, 'utf8');
}

function printHuman(summary) {
  console.log([
    'onsite_status',
    `ok=${bool(summary.ok)}`,
    `status=${summary.status}`,
    `minimum=${bool(summary.minimum_start_ready)}`,
    `practical=${bool(summary.practical_start_ready)}`,
    `event_path=${bool(summary.gates.event_path_ready)}`,
    `events=${summary.event_delivery.ws_event_count}`,
    `scan=${bool(summary.gates.scan_fallback_ready)}`,
    `auth=${summary.auth.reason ?? (summary.auth.meeting_search_usable ? 'ready' : 'unknown')}`,
    `real_axis=${bool(summary.gates.real_axis_active)}`,
    `annotations=${summary.gates.real_axis_annotation_count}`,
  ].join(' | '));
  console.log(`next=${compactText(summary.recommended_action, 240)}`);
  if (summary.event_delivery.failed_open_platform_checks.length) {
    console.log(`failed_open_platform_checks=${summary.event_delivery.failed_open_platform_checks.join(',')}`);
  }
  if (reportFile) console.log(`report_file=${reportFile}`);
}

try {
  const [demoStatus, diagnostics, passiveScan, authStatus, streamStatus, ingestInfo] = await Promise.all([
    fetchJson('/api/lark/real-demo/status'),
    fetchJson('/api/lark/delivery-diagnostics'),
    fetchJson('/api/lark/passive-meeting-scan'),
    fetchJson('/api/auth/lark/status'),
    fetchJson('/api/stream/status'),
    fetchJson('/api/annotation-ingest-info'),
  ]);
  const summary = summarize({ demoStatus, diagnostics, passiveScan, authStatus, streamStatus, ingestInfo });
  await writeReport(summary, { exitCode: summary.ok ? 0 : 2 });
  if (json) console.log(JSON.stringify(summary));
  else printHuman(summary);
  if (!summary.ok) process.exitCode = 2;
} catch (error) {
  await writeReport(null, { exitCode: 2, error }).catch((reportError) => {
    console.error(`failed to write report: ${reportError.message ?? String(reportError)}`);
  });
  console.error(error.message ?? String(error));
  process.exitCode = 2;
}
