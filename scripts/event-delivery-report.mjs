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
const reportFile = String(args.get('report-file') || args.get('write-report') || process.env.EVENT_DELIVERY_REPORT_FILE || '');
const failOnNoDelivery = args.get('fail-on-no-delivery') === 'true';

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

function boolLabel(value) {
  return value ? 'yes' : 'no';
}

function compactText(value, max = 140) {
  const text = String(value ?? '').replace(/\s+/g, ' ').trim();
  return text.length > max ? `${text.slice(0, max - 1)}...` : text;
}

function summarize({ diagnostics, eventsLog, progress, acceptance }) {
  const receiver = diagnostics.receiver ?? {};
  const evidence = diagnostics.evidence ?? {};
  const audit = diagnostics.real_meeting_event_audit ?? {};
  const completion = progress.completion_evidence ?? {};
  const wsConnected = receiver.ws_state === 'connected';
  const publicWebhookReady = Boolean(receiver.http_callback_public_https);
  const localReceiverReady = Boolean(wsConnected || publicWebhookReady);
  const localHandlersReady = Array.isArray(receiver.missing_registered_events)
    ? receiver.missing_registered_events.length === 0
    : Boolean(audit.local_handlers_ready);
  const parserReady = Boolean(audit.parser_self_test_passed ?? audit.evidence?.parser_self_test?.passed);
  const eventDeliverySeen = Number(evidence.real_event_count ?? 0) > 0
    || Number(evidence.ws_event_count ?? 0) > 0
    || Number(evidence.public_webhook_event_count ?? 0) > 0;
  const timelineStartedSeen = Number(evidence.ws_timeline_started_count ?? 0) > 0
    || Boolean(completion.event_axis_built);
  const realAxisActive = Boolean(completion.real_meeting_axis_active);
  const productComplete = Boolean(acceptance.product_acceptance_complete);
  const strictComplete = Boolean(acceptance.strict_event_acceptance_complete);
  const status = productComplete
    ? strictComplete ? 'complete_strict_event_axis' : 'complete_scan_or_binding_axis'
    : realAxisActive
      ? 'real_axis_waiting_annotation_or_broadcast'
      : eventDeliverySeen
        ? timelineStartedSeen ? 'event_seen_waiting_axis_or_annotation' : 'event_seen_no_axis'
        : localReceiverReady && localHandlersReady && parserReady
          ? 'waiting_for_lark_delivery'
          : 'receiver_not_ready';
  const ok = localReceiverReady && localHandlersReady && parserReady && (!failOnNoDelivery || eventDeliverySeen);
  return {
    type: 'event_delivery_report',
    ok,
    status,
    base_url: baseUrl,
    generated_at: new Date().toISOString(),
    local_receiver_ready: localReceiverReady,
    local_handlers_ready: localHandlersReady,
    parser_self_test_passed: parserReady,
    event_delivery_seen: eventDeliverySeen,
    timeline_started_seen: timelineStartedSeen,
    real_axis_active: realAxisActive,
    product_acceptance_complete: productComplete,
    strict_event_acceptance_complete: strictComplete,
    recommended_next_action: audit.next_action ?? progress.result?.operator_runbook?.primary_next_action ?? null,
    receiver: {
      ws_state: receiver.ws_state ?? null,
      ws_enabled: receiver.ws_enabled ?? null,
      registered_event_types: receiver.registered_event_types ?? [],
      required_direct_events: receiver.required_direct_events ?? [],
      fallback_context_events: receiver.fallback_context_events ?? [],
      missing_registered_events: receiver.missing_registered_events ?? [],
      missing_fallback_events: receiver.missing_fallback_events ?? [],
      http_callback_public_https: publicWebhookReady,
    },
    evidence: {
      ws_event_count: Number(evidence.ws_event_count ?? 0),
      real_event_count: Number(evidence.real_event_count ?? 0),
      public_webhook_event_count: Number(evidence.public_webhook_event_count ?? 0),
      ws_timeline_candidate_count: Number(evidence.ws_timeline_candidate_count ?? 0),
      ws_timeline_processed_count: Number(evidence.ws_timeline_processed_count ?? 0),
      ws_timeline_started_count: Number(evidence.ws_timeline_started_count ?? 0),
      last_ws_event_type: evidence.last_ws_event_type ?? null,
      event_audit_status: audit.status ?? null,
      recent_event_log_count: Array.isArray(eventsLog.items) ? eventsLog.items.length : 0,
      latest_event_log_item: Array.isArray(eventsLog.items) ? eventsLog.items[0] ?? null : null,
      real_axis_annotation_count: Number(completion.real_axis_annotation_count ?? 0),
      last_broadcast_real_axis_annotation_count: Number(completion.last_broadcast_real_axis_annotation_count ?? 0),
      axis_creation_mode: completion.axis_creation_mode ?? null,
      meeting_source: completion.meeting_source ?? null,
    },
    open_platform_checks: audit.required_open_platform_checks ?? [],
  };
}

function reportPayload(summary, { exitCode, error = null } = {}) {
  return {
    ...(summary ?? {
      type: 'event_delivery_report',
      ok: false,
      status: 'error',
      base_url: baseUrl,
      generated_at: new Date().toISOString(),
    }),
    report: {
      written_at: new Date().toISOString(),
      file: reportFile || null,
      command: 'scripts/event-delivery-report.mjs',
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
    'event_delivery_report',
    `ok=${boolLabel(summary.ok)}`,
    `status=${summary.status}`,
    `receiver=${summary.receiver.ws_state ?? 'unknown'}`,
    `events=${summary.evidence.ws_event_count}`,
    `candidates=${summary.evidence.ws_timeline_candidate_count}`,
    `started=${summary.evidence.ws_timeline_started_count}`,
    `real_axis=${boolLabel(summary.real_axis_active)}`,
    `annotations=${summary.evidence.real_axis_annotation_count}`,
    `audit=${summary.evidence.event_audit_status ?? 'unknown'}`,
  ].join(' | '));
  if (summary.recommended_next_action) {
    console.log(`next=${compactText(summary.recommended_next_action, 220)}`);
  }
  const failedChecks = summary.open_platform_checks
    .filter((check) => !check.ok)
    .map((check) => check.id);
  if (failedChecks.length) console.log(`failed_open_platform_checks=${failedChecks.join(',')}`);
  if (reportFile) console.log(`report_file=${reportFile}`);
}

try {
  const [diagnostics, eventsLog, progress, acceptance] = await Promise.all([
    fetchJson('/api/lark/delivery-diagnostics'),
    fetchJson('/api/lark/events-log'),
    fetchJson('/api/lark/real-demo/progress'),
    fetchJson('/api/lark/real-demo/acceptance'),
  ]);
  const summary = summarize({ diagnostics, eventsLog, progress, acceptance });
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
