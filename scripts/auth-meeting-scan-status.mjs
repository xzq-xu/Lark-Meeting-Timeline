#!/usr/bin/env node

import { spawn } from 'node:child_process';
import { mkdir, writeFile } from 'node:fs/promises';
import { dirname } from 'node:path';

const args = new Map();
for (const raw of process.argv.slice(2)) {
  const [key, ...rest] = raw.replace(/^--/, '').split('=');
  args.set(key, rest.length ? rest.join('=') : 'true');
}

const baseUrl = String(args.get('url') || process.env.REAL_DEMO_BASE_URL || 'http://localhost:8787').replace(/\/$/, '');
const json = args.get('json') === 'true';
const openAuth = args.get('open') === 'true';
const waitAuth = args.get('wait') === 'true';
const scanAfterAuth = args.get('scan-after-auth') !== 'false';
const requireReady = args.get('require-ready') === 'true';
const timeoutMs = Math.min(Math.max(Number(args.get('timeout-ms') ?? 300_000), 1_000), 30 * 60_000);
const intervalMs = Math.min(Math.max(Number(args.get('interval-ms') ?? 2_000), 500), 10_000);
const reportFile = String(args.get('report-file') || args.get('write-report') || process.env.AUTH_STATUS_REPORT_FILE || '');

async function fetchJson(path, options = {}) {
  const response = await fetch(`${baseUrl}${path}`, options);
  const text = await response.text();
  let body = null;
  try {
    body = text ? JSON.parse(text) : null;
  } catch {
    body = text;
  }
  if (!response.ok) {
    throw new Error(`${options.method || 'GET'} ${path} HTTP ${response.status}: ${typeof body === 'string' ? body : JSON.stringify(body)}`);
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

function redactUrl(value) {
  if (!value) return value;
  try {
    const url = new URL(value);
    if (url.searchParams.has('state')) url.searchParams.set('state', '<redacted>');
    return url.toString();
  } catch {
    return String(value).replace(/state=[^&]+/g, 'state=<redacted>');
  }
}

function callbackPreservedInUrl(value, callbackUrl) {
  if (!value || !callbackUrl) return false;
  const decodedOnce = decodeURIComponent(String(value));
  const decodedTwice = decodeURIComponent(decodedOnce);
  return decodedOnce.includes(callbackUrl) || decodedTwice.includes(callbackUrl);
}

async function checkAuthUrl(authStart = {}) {
  if (!authStart.auth_url) return null;
  const redirects = [];
  let currentUrl = authStart.auth_url;
  let finalStatus = null;
  try {
    for (let i = 0; i < 8; i += 1) {
      const response = await fetch(currentUrl, {
        redirect: 'manual',
        headers: { 'user-agent': 'Mozilla/5.0' },
      });
      finalStatus = response.status;
      const location = response.headers.get('location');
      redirects.push({
        status: response.status,
        url: redactUrl(currentUrl),
        location: redactUrl(location),
      });
      if (!location || response.status < 300 || response.status >= 400) break;
      currentUrl = new URL(location, currentUrl).toString();
    }
    const callbackUrl = authStart.redirect_uri ?? authStart.callback_url ?? null;
    const callbackPreserved = redirects.some((entry) => (
      callbackPreservedInUrl(entry.url, callbackUrl)
        || callbackPreservedInUrl(entry.location, callbackUrl)
    ));
    return {
      ok: Boolean(finalStatus && finalStatus < 500 && callbackPreserved),
      status: callbackPreserved ? 'callback_preserved' : 'callback_not_observed',
      final_http_status: finalStatus,
      final_url: redactUrl(currentUrl),
      callback_preserved: callbackPreserved,
      redirect_count: Math.max(0, redirects.length - 1),
      redirects,
    };
  } catch (error) {
    return {
      ok: false,
      status: 'auth_url_check_failed',
      error: error.message ?? String(error),
      final_url: redactUrl(currentUrl),
      callback_preserved: false,
      redirects,
    };
  }
}

async function openUrl(url) {
  const command = process.platform === 'darwin'
    ? 'open'
    : process.platform === 'win32'
      ? 'cmd'
      : 'xdg-open';
  const commandArgs = process.platform === 'win32'
    ? ['/c', 'start', '', url]
    : [url];
  const child = spawn(command, commandArgs, {
    detached: true,
    stdio: 'ignore',
  });
  child.unref();
}

async function loadAuthStart() {
  return fetchJson('/api/auth/lark/start?scope=vc%3Ameeting%3Areadonly');
}

async function loadSnapshot(authStart = null, authUrlCheck = null) {
  const [auth, passiveScan, demoStatus] = await Promise.all([
    fetchJson('/api/auth/lark/status'),
    fetchJson('/api/lark/passive-meeting-scan'),
    fetchJson('/api/lark/real-demo/status'),
  ]);
  const meetingSearch = auth.meeting_search ?? {};
  const ready = Boolean(meetingSearch.usable);
  const status = ready
    ? 'auth_ready'
    : meetingSearch.reason === 'missing_scope'
      ? 'scope_missing'
      : meetingSearch.reason === 'oauth_token_expired'
        ? 'oauth_expired'
        : auth.token_present
          ? 'auth_not_usable'
          : 'login_required';
  return {
    type: 'auth_meeting_scan_status',
    ok: ready,
    status,
    base_url: baseUrl,
    generated_at: new Date().toISOString(),
    recommended_action: ready
      ? '当前用户会议扫描授权可用；可以直接开启飞书会议，或立即触发一次扫描。'
      : '打开授权 URL 并完成 vc:meeting.search:read 授权；callback 回到本机后服务端会自动开启当前用户会议扫描。',
    auth: {
      authenticated: Boolean(auth.authenticated),
      token_present: Boolean(auth.token_present),
      expired: Boolean(auth.expired),
      refresh_token_present: Boolean(auth.refresh_token_present),
      meeting_search_usable: ready,
      scope_present: Boolean(meetingSearch.scope_present),
      reason: meetingSearch.reason ?? null,
      required_scope: meetingSearch.required_scope ?? 'vc:meeting.search:read',
      next_action: meetingSearch.next_action ?? null,
    },
    passive_scan: {
      enabled: Boolean(passiveScan.enabled),
      tenant_fallback_enabled: Boolean(passiveScan.tenant_fallback_enabled),
      server_loop: passiveScan.server_loop ?? null,
      last_result: passiveScan.last_result ?? null,
    },
    real_demo: {
      status: demoStatus.status ?? null,
      ready_to_open_meeting: Boolean(demoStatus.ready_to_open_meeting),
      real_axis_active: Boolean(demoStatus.gates?.real_meeting_axis_active),
      real_axis_annotation_count: Number(demoStatus.gates?.real_axis_annotation_count ?? 0),
    },
    auth_start: authStart ? {
      auth_url: redactUrl(authStart.auth_url),
      redirect_uri: authStart.redirect_uri ?? authStart.callback_url ?? null,
      requested_scope: authStart.requested_scope ?? null,
      scopes: authStart.scopes ?? [],
      scope_present: authStart.scope_present ?? {},
      state_created_at: authStart.state_created_at ?? null,
      state_expires_at: authStart.state_expires_at ?? null,
    } : null,
    auth_url_check: authUrlCheck,
  };
}

async function triggerScan() {
  return fetchJson('/api/lark/passive-meeting-scan', {
    method: 'POST',
    headers: { 'content-type': 'application/json; charset=utf-8' },
    body: JSON.stringify({
      enabled: true,
      tenant_fallback_enabled: false,
      trigger_now: true,
      source: 'auth_status_ready_scan',
    }),
  });
}

function reportPayload(summary, { exitCode, error = null } = {}) {
  return {
    ...(summary ?? {
      type: 'auth_meeting_scan_status',
      ok: false,
      status: 'error',
      base_url: baseUrl,
      generated_at: new Date().toISOString(),
    }),
    report: {
      written_at: new Date().toISOString(),
      file: reportFile || null,
      command: 'scripts/auth-meeting-scan-status.mjs',
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

function printHuman(summary, { opened = false, scan = null } = {}) {
  console.log([
    'auth_meeting_scan_status',
    `ok=${bool(summary.ok)}`,
    `status=${summary.status}`,
    `auth=${summary.auth.reason ?? (summary.auth.meeting_search_usable ? 'ready' : 'unknown')}`,
    `scope=${bool(summary.auth.scope_present)}`,
    `scan_enabled=${bool(summary.passive_scan.enabled)}`,
    `tenant_fallback=${bool(summary.passive_scan.tenant_fallback_enabled)}`,
    `callback_preserved=${bool(summary.auth_url_check?.callback_preserved)}`,
    `opened=${bool(opened)}`,
    scan ? `scan=${scan.trigger?.status ?? scan.passive_meeting_scan?.last_result?.status ?? 'unknown'}` : '',
  ].filter(Boolean).join(' | '));
  console.log(`next=${compactText(summary.recommended_action, 240)}`);
  if (summary.auth_start?.auth_url) console.log(`auth_url=${summary.auth_start.auth_url}`);
  if (reportFile) console.log(`report_file=${reportFile}`);
}

try {
  const authStart = await loadAuthStart();
  const authUrlCheck = await checkAuthUrl(authStart);
  if (openAuth) await openUrl(authStart.auth_url);
  let summary = await loadSnapshot(authStart, authUrlCheck);
  const startedAt = Date.now();
  while (waitAuth && !summary.ok && Date.now() - startedAt < timeoutMs) {
    await new Promise((resolve) => setTimeout(resolve, intervalMs));
    summary = await loadSnapshot(authStart, authUrlCheck);
  }
  let scan = null;
  if (summary.ok && scanAfterAuth) {
    scan = await triggerScan();
    summary = await loadSnapshot(authStart, authUrlCheck);
  }
  const exitCode = summary.ok || !requireReady ? 0 : 2;
  await writeReport({ ...summary, scan_after_auth: scan }, { exitCode });
  if (json) console.log(JSON.stringify({ ...summary, scan_after_auth: scan }));
  else printHuman(summary, { opened: openAuth, scan });
  process.exitCode = exitCode;
} catch (error) {
  await writeReport(null, { exitCode: 2, error }).catch((reportError) => {
    console.error(`failed to write report: ${reportError.message ?? String(reportError)}`);
  });
  console.error(error.message ?? String(error));
  process.exitCode = 2;
}
