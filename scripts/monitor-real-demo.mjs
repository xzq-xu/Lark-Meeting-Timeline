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
const once = args.get('once') === 'true';
const json = args.get('json') === 'true';
const prepare = args.get('prepare') === 'true';
const openAuth = args.get('open-auth') === 'true';
const waitAuth = args.get('wait-auth') === 'true';
const scanAfterAuth = args.get('scan-after-auth') !== 'false';
const autoMark = args.get('auto-mark') === 'true';
const requireEventAxis = args.get('require-event-axis') === 'true' || args.get('strict-event-axis') === 'true';
const markId = String(args.get('mark-id') || `monitor-real-demo-${Date.now()}`);
const markLabel = String(args.get('mark-label') || 'monitor 自动标注');
const markSource = String(args.get('mark-source') || 'monitor_open_annotation');
const intervalMs = Math.min(Math.max(Number(args.get('interval-ms') ?? 1000), 500), 5000);
const timeoutMs = Number(args.get('timeout-ms') ?? 0);
const authTimeoutMs = Number(args.get('auth-timeout-ms') ?? timeoutMs ?? 0);
const reportFile = args.get('report-file') || args.get('write-report') || process.env.REAL_DEMO_REPORT_FILE || '';
let autoMarkWritten = false;
let lastProgress = null;

function bool(value) {
  return value ? 'yes' : 'no';
}

function compactText(value, max = 110) {
  const text = String(value ?? '').replace(/\s+/g, ' ').trim();
  return text.length > max ? `${text.slice(0, max - 1)}...` : text;
}

function authLabel(progress = {}) {
  const auth = progress.result?.auth?.meeting_search ?? {};
  if (auth.usable) return 'ok';
  return auth.reason ?? 'unavailable';
}

function scanLabel(progress = {}) {
  const scan = progress.result?.passive_meeting_scan?.last_result ?? {};
  if (!scan.reason && !scan.status) return 'none';
  const label = scan.reason ?? scan.status;
  return scan.next_allowed_at ? `${label}->${scan.next_allowed_at}` : label;
}

function axisSourceLabel(evidence = {}) {
  if (!evidence.real_meeting_axis_active) return 'none';
  return evidence.meeting_source || 'unknown';
}

function axisModeLabel(evidence = {}) {
  if (!evidence.real_meeting_axis_active) return 'not_built';
  return evidence.axis_creation_mode || (evidence.event_axis_built ? 'meeting_start_event' : 'unknown');
}

function eventAxisReady(evidence = {}) {
  return Boolean(evidence.real_meeting_axis_active && evidence.event_axis_built);
}

function requirementsLabel(evidence = {}) {
  const requirements = Array.isArray(evidence.requirements) ? evidence.requirements : [];
  if (!requirements.length) return 'none';
  const failed = requirements.filter((item) => !item.ok).map((item) => item.id);
  return failed.length ? `todo:${failed.join(',')}` : 'all_ok';
}

function authReady(progress = {}) {
  return Boolean(progress.result?.auth?.meeting_search?.usable);
}

function authSummary(progress = {}) {
  const auth = progress.result?.auth?.meeting_search ?? {};
  return [
    `auth=${authLabel(progress)}`,
    auth.required_scope ? `required_scope=${auth.required_scope}` : '',
    auth.next_action ? `next="${compactText(auth.next_action)}"` : '',
  ].filter(Boolean).join(' | ');
}

function summaryLine(progress = {}) {
  const evidence = progress.completion_evidence ?? {};
  return [
    new Date(progress.generated_at ?? Date.now()).toISOString(),
    `status=${progress.status ?? 'unknown'}`,
    `complete=${bool(evidence.real_demo_complete ?? progress.observed)}`,
    `real_axis=${bool(evidence.real_meeting_axis_active)}`,
    `event_axis=${bool(eventAxisReady(evidence))}`,
    `axis_mode=${axisModeLabel(evidence)}`,
    `axis_source=${axisSourceLabel(evidence)}`,
    `annotations=${evidence.real_axis_annotation_count ?? 0}`,
    `events=${evidence.ws_event_count ?? 0}`,
    `audit=${evidence.event_audit_status ?? 'unknown'}`,
    `requirements=${requirementsLabel(evidence)}`,
    `auth=${authLabel(progress)}`,
    `scan=${compactText(scanLabel(progress), 72)}`,
    `device_stream=${evidence.device_stream_status ?? 'unknown'}`,
    requireEventAxis ? 'strict_event_axis=yes' : '',
    evidence.next_action ? `next="${compactText(evidence.next_action)}"` : '',
  ].filter(Boolean).join(' | ');
}

async function getProgress() {
  const response = await fetch(`${baseUrl}/api/lark/real-demo/progress`);
  const body = await response.text();
  if (!response.ok) throw new Error(`progress HTTP ${response.status}: ${body}`);
  lastProgress = JSON.parse(body);
  return lastProgress;
}

async function writeReport(progress = lastProgress, extra = {}) {
  if (!reportFile) return;
  const payload = {
    type: 'real_demo_acceptance_report',
    written_at: new Date().toISOString(),
    base_url: baseUrl,
    command: 'scripts/monitor-real-demo.mjs',
    options: {
      once,
      prepare,
      open_auth: openAuth,
      wait_auth: waitAuth,
      scan_after_auth: scanAfterAuth,
      auto_mark: autoMark,
      require_event_axis: requireEventAxis,
      timeout_ms: timeoutMs,
      auth_timeout_ms: authTimeoutMs,
      mark_id: markId,
      mark_source: markSource,
    },
    outcome: {
      exit_code: extra.exit_code ?? null,
      error: extra.error ?? null,
      completed: Boolean(progress?.completion_evidence?.real_demo_complete ?? progress?.observed),
      strict_event_axis_complete: Boolean(progress?.completion_evidence?.strict_event_axis_complete),
      product_acceptance_complete: Boolean(progress?.completion_evidence?.real_demo_complete ?? progress?.observed),
      axis_creation_mode: progress?.completion_evidence?.axis_creation_mode ?? null,
      meeting_source: progress?.completion_evidence?.meeting_source ?? null,
      real_axis_annotation_count: progress?.completion_evidence?.real_axis_annotation_count ?? 0,
      last_broadcast_real_axis_annotation_count: progress?.completion_evidence?.last_broadcast_real_axis_annotation_count ?? 0,
      requirements: Array.isArray(progress?.completion_evidence?.requirements)
        ? progress.completion_evidence.requirements.map((item) => ({ id: item.id, ok: Boolean(item.ok), observed: item.observed ?? null }))
        : [],
    },
    progress: progress ?? null,
  };
  await mkdir(dirname(reportFile), { recursive: true });
  await writeFile(reportFile, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
  if (!json) console.log(`report_file=${reportFile}`);
}

async function prepareRealDemo() {
  const response = await fetch(`${baseUrl}/api/lark/real-demo/prepare`, {
    method: 'POST',
    headers: { 'content-type': 'application/json; charset=utf-8' },
    body: JSON.stringify({
      auto_annotation: true,
      device_simulator: true,
      device_stream: true,
      passive_scan: true,
      note: 'monitor_real_demo_cli',
    }),
  });
  const body = await response.text();
  if (!response.ok) throw new Error(`prepare HTTP ${response.status}: ${body}`);
  return JSON.parse(body);
}

async function writeOpenAnnotation(progress = {}) {
  if (!autoMark || autoMarkWritten) return { written: false, reason: 'disabled_or_already_written' };
  const evidence = progress.completion_evidence ?? {};
  if (!evidence.real_meeting_axis_active) return { written: false, reason: 'waiting_for_real_axis' };
  if (Number(evidence.real_axis_annotation_count ?? 0) > 0) return { written: false, reason: 'annotation_already_present' };

  const response = await fetch(`${baseUrl}/api/annotations`, {
    method: 'POST',
    headers: {
      'content-type': 'application/json; charset=utf-8',
      'x-hmp-device-id': 'monitor-real-demo',
      'x-hmp-device-type': markSource,
    },
    body: JSON.stringify({
      id: markId,
      source: markSource,
      kind: 'handwriting_trigger',
      label: markLabel,
      text_candidates: [markLabel, 'real-time annotation'],
      intent: 'attention',
      captured_at_ms: Date.now(),
      payload: {
        demo: 'monitor-real-demo',
        note: 'Written through the open annotation API after a real meeting axis appeared.',
      },
    }),
  });
  const body = await response.text();
  if (!response.ok) throw new Error(`auto-mark HTTP ${response.status}: ${body}`);
  autoMarkWritten = true;
  return { written: true, result: JSON.parse(body) };
}

function printAutoMarkResult(mark = {}) {
  if (!mark.written) return;
  const ack = mark.result?.ack ?? {};
  if (json) {
    console.log(JSON.stringify({
      type: 'auto_mark_written',
      annotation_id: ack.annotation_id,
      on_real_axis: ack.on_real_axis,
      normalized_time_ms: ack.normalized_time_ms,
      binding_state: ack.binding_state,
    }));
    return;
  }
  console.log([
    'auto_mark_written=true',
    `id=${ack.annotation_id ?? markId}`,
    `on_real_axis=${bool(ack.on_real_axis)}`,
    `time_ms=${ack.normalized_time_ms ?? 'unknown'}`,
    `binding=${ack.binding_state ?? 'unknown'}`,
  ].join(' | '));
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

async function openAuthIfNeeded() {
  const progress = await getProgress();
  const authUrl = progress.result?.auth_start?.redirect_url;
  const needsAuth = progress.result?.auth?.meeting_search?.needs_reauth;
  if (!authUrl || !needsAuth) {
    return {
      opened: false,
      reason: needsAuth ? 'auth_url_missing' : 'auth_not_required',
      progress,
    };
  }
  await openUrl(authUrl);
  return { opened: true, url: authUrl, progress };
}

async function waitForAuthIfRequested() {
  if (!waitAuth) return { waited: false, ready: false, progress: null };
  const startedAt = Date.now();
  let lastSignature = '';
  for (;;) {
    const progress = await getProgress();
    const nextSignature = JSON.stringify({
      ready: authReady(progress),
      auth: authLabel(progress),
      next: progress.result?.auth?.meeting_search?.next_action,
    });
    if (nextSignature !== lastSignature) {
      lastSignature = nextSignature;
      if (json) console.log(JSON.stringify({ type: 'auth_wait', ready: authReady(progress), progress }));
      else console.log(`auth_wait | ready=${bool(authReady(progress))} | ${authSummary(progress)}`);
    }
    if (authReady(progress)) return { waited: true, ready: true, progress };
    if (authTimeoutMs > 0 && Date.now() - startedAt >= authTimeoutMs) {
      return { waited: true, ready: false, progress };
    }
    await new Promise((resolve) => setTimeout(resolve, Math.min(Math.max(intervalMs, 1000), 5000)));
  }
}

async function triggerPassiveScanAfterAuthIfNeeded(authWaitResult = {}) {
  if (!scanAfterAuth || !authWaitResult.ready) {
    return { triggered: false, reason: scanAfterAuth ? 'auth_not_ready' : 'disabled' };
  }
  const response = await fetch(`${baseUrl}/api/lark/passive-meeting-scan`, {
    method: 'POST',
    headers: { 'content-type': 'application/json; charset=utf-8' },
    body: JSON.stringify({
      enabled: true,
      trigger_now: true,
      source: 'monitor_auth_ready_scan',
    }),
  });
  const body = await response.text();
  if (!response.ok) throw new Error(`auth passive scan HTTP ${response.status}: ${body}`);
  return { triggered: true, result: JSON.parse(body) };
}

function printAuthScanResult(scan = {}) {
  if (!scan.triggered) return;
  const trigger = scan.result?.trigger ?? {};
  if (json) {
    console.log(JSON.stringify({
      type: 'auth_scan_triggered',
      status: trigger.status ?? null,
      reason: trigger.reason ?? null,
      selected_meeting_id: trigger.selected_meeting_id ?? trigger.state?.meeting?.meeting_id ?? null,
      meeting_source: trigger.state?.meeting?.source ?? null,
    }));
    return;
  }
  console.log([
    'auth_scan_triggered=true',
    `status=${trigger.status ?? 'unknown'}`,
    trigger.reason ? `reason=${trigger.reason}` : '',
    trigger.selected_meeting_id || trigger.state?.meeting?.meeting_id
      ? `meeting=${trigger.selected_meeting_id ?? trigger.state?.meeting?.meeting_id}`
      : '',
    trigger.state?.meeting?.source ? `source=${trigger.state.meeting.source}` : '',
  ].filter(Boolean).join(' | '));
}

function signature(progress = {}) {
  const evidence = progress.completion_evidence ?? {};
  return JSON.stringify({
    status: progress.status,
    complete: evidence.real_demo_complete ?? progress.observed,
    real_axis: evidence.real_meeting_axis_active,
    event_axis: eventAxisReady(evidence),
    axis_mode: axisModeLabel(evidence),
    axis_source: axisSourceLabel(evidence),
    annotations: evidence.real_axis_annotation_count,
    events: evidence.ws_event_count,
    audit: evidence.event_audit_status,
    requirements: requirementsLabel(evidence),
    auth: authLabel(progress),
    scan: scanLabel(progress),
    device_stream: evidence.device_stream_status,
    strict_event_axis: requireEventAxis,
    next: evidence.next_action,
  });
}

async function monitorStream() {
  const controller = new AbortController();
  let timeout = null;
  if (timeoutMs > 0) {
    timeout = setTimeout(() => controller.abort(new Error(`timeout after ${timeoutMs}ms`)), timeoutMs);
  }
  process.once('SIGINT', () => controller.abort(new Error('interrupted')));

  const response = await fetch(`${baseUrl}/api/lark/real-demo/progress-stream?interval_ms=${intervalMs}`, {
    signal: controller.signal,
  });
  if (!response.ok) {
    const body = await response.text();
    throw new Error(`progress-stream HTTP ${response.status}: ${body}`);
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';
  let lastSignature = '';

  const handleBlock = async (block) => {
    const event = block.split('\n').find((line) => line.startsWith('event: '))?.slice(7);
    if (event !== 'progress') return false;
    const data = block
      .split('\n')
      .filter((line) => line.startsWith('data: '))
      .map((line) => line.slice(6))
      .join('\n');
    if (!data) return false;
    const progress = JSON.parse(data);
    lastProgress = progress;
    const nextSignature = signature(progress);
    if (nextSignature !== lastSignature) {
      lastSignature = nextSignature;
      if (json) console.log(JSON.stringify(progress));
      else console.log(summaryLine(progress));
    }
    const mark = await writeOpenAnnotation(progress);
    printAutoMarkResult(mark);
    const evidence = progress.completion_evidence ?? {};
    const complete = Boolean(evidence.real_demo_complete ?? progress.observed);
    return complete && (!requireEventAxis || eventAxisReady(evidence));
  };

  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) return 1;
      buffer += decoder.decode(value, { stream: true });
      let splitAt = buffer.indexOf('\n\n');
      while (splitAt >= 0) {
        const complete = await handleBlock(buffer.slice(0, splitAt));
        buffer = buffer.slice(splitAt + 2);
        if (complete) return 0;
        splitAt = buffer.indexOf('\n\n');
      }
    }
  } finally {
    if (timeout) clearTimeout(timeout);
    await reader.cancel().catch(() => {});
  }
}

try {
  if (prepare) {
    const prepared = await prepareRealDemo();
    if (!json) {
      const session = prepared.real_demo_session ?? {};
      console.log(`prepared=true | prepared_at=${session.prepared_at ?? 'unknown'} | device_stream=${prepared.device_stream_simulator?.status ?? 'unknown'}`);
    }
  }
  if (openAuth) {
    const auth = await openAuthIfNeeded();
    if (!json) {
      console.log(auth.opened
        ? `auth_opened=true | url=${auth.url}`
        : `auth_opened=false | reason=${auth.reason}`);
    }
  }
  const authWaitResult = await waitForAuthIfRequested();
  if (authWaitResult.waited && !authWaitResult.ready) {
    throw new Error(`auth wait timed out: ${authSummary(authWaitResult.progress)}`);
  }
  const authScanResult = await triggerPassiveScanAfterAuthIfNeeded(authWaitResult);
  printAuthScanResult(authScanResult);
  if (once) {
    let progress = await getProgress();
    const mark = await writeOpenAnnotation(progress);
    printAutoMarkResult(mark);
    if (mark.written) progress = await getProgress();
    if (json) console.log(JSON.stringify(progress));
    else console.log(summaryLine(progress));
    if (requireEventAxis && progress.completion_evidence?.real_demo_complete && !eventAxisReady(progress.completion_evidence)) {
      throw new Error('strict event-axis required, but current completed axis was not created from a meeting_start event');
    }
    await writeReport(progress, { exit_code: 0 });
  } else {
    process.exitCode = await monitorStream();
    await writeReport(lastProgress, { exit_code: process.exitCode });
  }
} catch (error) {
  console.error(error.message ?? String(error));
  await writeReport(lastProgress, { exit_code: 2, error: error.message ?? String(error) }).catch(() => {});
  process.exitCode = 2;
}
