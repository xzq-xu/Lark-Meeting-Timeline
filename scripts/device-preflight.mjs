#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname } from 'node:path';

const args = new Map();
for (const raw of process.argv.slice(2)) {
  const [key, ...rest] = raw.replace(/^--/, '').split('=');
  args.set(key, rest.length ? rest.join('=') : 'true');
}

const baseUrl = String(args.get('url') || process.env.DEVICE_MARK_BASE_URL || process.env.REAL_DEMO_BASE_URL || 'http://localhost:8787').replace(/\/$/, '');
const json = args.get('json') === 'true';
const deviceId = String(args.get('device-id') || 'hanwang-preflight-001');
const deviceType = String(args.get('device-type') || 'hanwang_epaper');
const reportFile = String(args.get('report-file') || args.get('write-report') || process.env.DEVICE_PREFLIGHT_REPORT_FILE || '');

function bool(value) {
  return value ? 'yes' : 'no';
}

function compactText(value, max = 140) {
  const text = String(value ?? '').replace(/\s+/g, ' ').trim();
  return text.length > max ? `${text.slice(0, max - 1)}...` : text;
}

async function fetchJson(url, options = {}) {
  const response = await fetch(url, options);
  const text = await response.text();
  let body = null;
  try {
    body = text ? JSON.parse(text) : null;
  } catch {
    body = text;
  }
  if (!response.ok) {
    throw new Error(`${options.method || 'GET'} ${url} HTTP ${response.status}: ${typeof body === 'string' ? body : JSON.stringify(body)}`);
  }
  return { response, body };
}

async function loadIngestInfo() {
  return (await fetchJson(`${baseUrl}/api/annotation-ingest-info`)).body;
}

async function checkCors(info = {}) {
  const endpoint = info.endpoint || `${baseUrl}/api/annotations`;
  const response = await fetch(endpoint, {
    method: 'OPTIONS',
    headers: {
      origin: 'http://device.local',
      'access-control-request-method': 'POST',
      'access-control-request-headers': 'content-type,x-hmp-device-id,x-hmp-device-type',
    },
  });
  const allowHeaders = response.headers.get('access-control-allow-headers') ?? '';
  return {
    ok: response.status === 204
      && /x-hmp-device-id/i.test(allowHeaders)
      && /x-hmp-device-type/i.test(allowHeaders),
    status: response.status,
    allow_headers: allowHeaders,
  };
}

async function syncClock(info = {}) {
  const endpoint = info.clock_sync?.endpoint || info.time_sync_url || `${baseUrl}/api/time`;
  const clientSendAtMs = Date.now();
  const separator = endpoint.includes('?') ? '&' : '?';
  const { body } = await fetchJson(`${endpoint}${separator}client_send_at_ms=${clientSendAtMs}`);
  const clientReceiveAtMs = Date.now();
  const serverTimeMs = Number(body.server_time_ms);
  const midpointMs = (clientSendAtMs + clientReceiveAtMs) / 2;
  const offsetMs = Number.isFinite(serverTimeMs) ? Math.round(serverTimeMs - midpointMs) : null;
  return {
    ok: Number.isFinite(offsetMs),
    offset_ms: offsetMs,
    rtt_ms: clientReceiveAtMs - clientSendAtMs,
    endpoint,
    server_time_ms: body.server_time_ms ?? null,
  };
}

async function loadStreamStatus(info = {}) {
  const url = info.stream_status_url || `${baseUrl}/api/stream/status`;
  const { body } = await fetchJson(url);
  return {
    ok: typeof body.current_clients === 'number' && typeof body.broadcast_count === 'number',
    url,
    current_clients: body.current_clients,
    broadcast_count: body.broadcast_count,
    last_broadcast_at: body.last_broadcast_at ?? null,
  };
}

async function loadAcceptance(info = {}) {
  const url = info.real_demo_acceptance_url || `${baseUrl}/api/lark/real-demo/acceptance`;
  const { body } = await fetchJson(url);
  return {
    ok: typeof body.product_acceptance_complete === 'boolean' && typeof body.verdict === 'string',
    url,
    verdict: body.verdict,
    product_acceptance_complete: body.product_acceptance_complete,
    strict_event_acceptance_complete: body.strict_event_acceptance_complete,
    recommended_command: body.recommended_command,
    missing: Array.isArray(body.missing_product_requirements)
      ? body.missing_product_requirements.map((item) => item.id)
      : [],
  };
}

function buildSummary(info, checks) {
  const route = info.annotation_route ?? {};
  const contract = info.device_client_contract ?? {};
  const ok = Boolean(
    info.endpoint
      && info.schema_url
      && contract.preferred_time_field === 'captured_at_ms'
      && checks.cors.ok
      && checks.clock.ok
      && checks.stream.ok
      && checks.acceptance.ok,
  );
  return {
    type: 'device_preflight',
    ok,
    endpoint: info.endpoint,
    batch_endpoint: info.batch_endpoint,
    schema_url: info.schema_url,
    stream_url: info.stream_url,
    status_url: info.annotation_status_url,
    binding_lookup_url: info.binding_lookup_url,
    real_demo_acceptance_url: info.real_demo_acceptance_url,
    route: {
      mode: route.mode,
      real_meeting_axis_active: route.real_meeting_axis_active,
      pending_binding: route.pending_binding,
      requires_captured_at_to_create_pending: route.requires_captured_at_to_create_pending,
    },
    device: {
      id: deviceId,
      type: deviceType,
      preferred_time_field: contract.preferred_time_field,
      clock_sync_required: contract.clock_sync_required,
      product_acceptance_condition: contract.product_acceptance_condition,
    },
    checks,
  };
}

function reportPayload(summary, { exitCode = summary?.ok ? 0 : 2, error = null } = {}) {
  return {
    ...(summary ?? { type: 'device_preflight', ok: false }),
    report: {
      written_at: new Date().toISOString(),
      file: reportFile || null,
      command: 'scripts/device-preflight.mjs',
      base_url: baseUrl,
      exit_code: exitCode,
      error: error ? String(error.message ?? error) : null,
    },
  };
}

async function writeReport(summary, options = {}) {
  if (!reportFile) return;
  await mkdir(dirname(reportFile), { recursive: true });
  await writeFile(reportFile, `${JSON.stringify(reportPayload(summary, options), null, 2)}\n`, 'utf8');
}

function printHuman(summary) {
  const checks = summary.checks;
  console.log([
    'device_preflight',
    `ok=${bool(summary.ok)}`,
    `endpoint=${summary.endpoint}`,
    `route=${summary.route.mode}`,
    `real_axis=${bool(summary.route.real_meeting_axis_active)}`,
    `cors=${bool(checks.cors.ok)}`,
    `clock=${bool(checks.clock.ok)} offset_ms=${checks.clock.offset_ms ?? 'unknown'} rtt_ms=${checks.clock.rtt_ms ?? 'unknown'}`,
    `stream=${bool(checks.stream.ok)} clients=${checks.stream.current_clients ?? 0}`,
    `acceptance=${checks.acceptance.verdict}`,
    checks.acceptance.recommended_command ? `next="${compactText(checks.acceptance.recommended_command)}"` : '',
  ].filter(Boolean).join(' | '));
  if (checks.acceptance.missing?.length) {
    console.log(`missing_product_requirements=${checks.acceptance.missing.join(',')}`);
  }
  if (reportFile) console.log(`report_file=${reportFile}`);
}

try {
  const info = await loadIngestInfo();
  const checks = {
    cors: await checkCors(info),
    clock: await syncClock(info),
    stream: await loadStreamStatus(info),
    acceptance: await loadAcceptance(info),
  };
  const summary = buildSummary(info, checks);
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
