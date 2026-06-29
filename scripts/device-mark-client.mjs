#!/usr/bin/env node

const args = new Map();
for (const raw of process.argv.slice(2)) {
  const [key, ...rest] = raw.replace(/^--/, '').split('=');
  args.set(key, rest.length ? rest.join('=') : 'true');
}

const baseUrl = String(args.get('url') || process.env.DEVICE_MARK_BASE_URL || process.env.REAL_DEMO_BASE_URL || 'http://localhost:8787').replace(/\/$/, '');
const json = args.get('json') === 'true';
const waitRealAxis = args.get('wait-real-axis') === 'true' || args.get('wait') === 'true';
const timeoutMs = Math.min(Math.max(Number(args.get('timeout-ms') ?? (waitRealAxis ? 60_000 : 0)), 0), 10 * 60_000);
const intervalMs = Math.min(Math.max(Number(args.get('interval-ms') ?? 500), 100), 5000);
const now = Date.now();
const annotationId = String(args.get('id') || args.get('annotation-id') || `device-mark-${now}`);
const label = String(args.get('label') || 'why?');
const source = String(args.get('source') || 'hanwang_epaper_cli');
const deviceId = String(args.get('device-id') || 'hanwang-cli-001');
const deviceType = String(args.get('device-type') || 'hanwang_epaper');
const clockSyncEnabled = args.get('sync-clock') !== 'false' && args.get('clock-sync') !== 'false';
const inlineMeetingSessionEnabled = args.get('meeting-session') === 'true'
  || args.get('start-meeting-session') === 'true'
  || args.has('meeting-id')
  || args.has('meeting_id')
  || args.has('meeting-title')
  || args.has('meeting_title')
  || args.has('meeting-url')
  || args.has('meeting_url')
  || args.has('meeting-start-ms')
  || args.has('meeting_start_ms')
  || args.has('meeting-start')
  || args.has('meeting_start');

function bool(value) {
  return value ? 'yes' : 'no';
}

function compactText(value, max = 120) {
  const text = String(value ?? '').replace(/\s+/g, ' ').trim();
  return text.length > max ? `${text.slice(0, max - 1)}...` : text;
}

function hasExplicitCapturedAt() {
  return (args.get('captured-at-ms') ?? args.get('captured_at_ms') ?? args.get('captured-at') ?? args.get('captured_at')) != null;
}

function parseCapturedAtMs(clockOffsetMs = 0) {
  const explicit = args.get('captured-at-ms') ?? args.get('captured_at_ms');
  if (explicit != null) {
    const value = Number(explicit);
    if (!Number.isFinite(value)) throw new Error(`invalid --captured-at-ms: ${explicit}`);
    return Math.round(value);
  }
  const capturedAt = args.get('captured-at') ?? args.get('captured_at');
  if (capturedAt != null) {
    const value = Date.parse(String(capturedAt));
    if (!Number.isFinite(value)) throw new Error(`invalid --captured-at: ${capturedAt}`);
    return value;
  }
  return Math.round(Date.now() + clockOffsetMs);
}

function firstArg(...names) {
  for (const name of names) {
    const value = args.get(name);
    if (value != null && value !== '') return value;
  }
  return null;
}

function cleanObject(input = {}) {
  return Object.fromEntries(Object.entries(input).filter(([, value]) => value != null && value !== ''));
}

function buildMeetingSession() {
  if (!inlineMeetingSessionEnabled) return null;
  const startMs = firstArg('meeting-start-ms', 'meeting_start_ms', 'meeting-start-time-ms', 'meeting_start_time_ms');
  const session = cleanObject({
    platform: firstArg('meeting-platform', 'meeting_platform') || 'lark',
    meeting_id: firstArg('meeting-id', 'meeting_id', 'session-id', 'session_id'),
    title: firstArg('meeting-title', 'meeting_title', 'title'),
    meeting_url: firstArg('meeting-url', 'meeting_url', 'join-url', 'join_url'),
    start_time_ms: startMs == null ? null : Number(startMs),
    start_time: firstArg('meeting-start', 'meeting_start', 'meeting-start-time', 'meeting_start_time'),
    detector_source: firstArg('detector-source', 'detector_source') || 'device_mark_client',
  });
  if (session.start_time_ms != null && !Number.isFinite(session.start_time_ms)) {
    throw new Error(`invalid --meeting-start-ms: ${startMs}`);
  }
  return session;
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
  return body;
}

async function loadIngestInfo() {
  return fetchJson(`${baseUrl}/api/annotation-ingest-info`);
}

function statusUrlFromInfo(info = {}) {
  const template = info.annotation_status_url || `${baseUrl}/api/annotation-status?id={annotation_id}`;
  return template.replace('{annotation_id}', encodeURIComponent(annotationId));
}

async function syncDeviceClock(info = {}) {
  if (!clockSyncEnabled) return { applied: false, offset_ms: 0, reason: 'disabled' };
  if (hasExplicitCapturedAt()) return { applied: false, offset_ms: 0, reason: 'explicit_captured_at' };
  const endpoint = info.clock_sync?.endpoint || info.time_sync_url || `${baseUrl}/api/time`;
  const clientSendAtMs = Date.now();
  const separator = endpoint.includes('?') ? '&' : '?';
  const sync = await fetchJson(`${endpoint}${separator}client_send_at_ms=${clientSendAtMs}`);
  const clientReceiveAtMs = Date.now();
  const serverTimeMs = Number(sync.server_time_ms);
  if (!Number.isFinite(serverTimeMs)) return { applied: false, offset_ms: 0, reason: 'server_time_missing', response: sync };
  const midpointMs = (clientSendAtMs + clientReceiveAtMs) / 2;
  return {
    applied: true,
    offset_ms: Math.round(serverTimeMs - midpointMs),
    client_send_at_ms: clientSendAtMs,
    client_receive_at_ms: clientReceiveAtMs,
    server_time_ms: serverTimeMs,
    rtt_ms: clientReceiveAtMs - clientSendAtMs,
  };
}

function buildPayload(info = {}, clockSync = {}) {
  const capturedAtMs = parseCapturedAtMs(clockSync.offset_ms ?? 0);
  const base = info.minimal_payload ?? {};
  const payload = {
    ...base,
    id: annotationId,
    source,
    captured_at_ms: capturedAtMs,
    kind: String(args.get('kind') || base.kind || 'handwriting_trigger'),
    label,
    text_candidates: String(args.get('text-candidates') || '')
      ? String(args.get('text-candidates')).split('|').map((item) => item.trim()).filter(Boolean)
      : [label],
    intent: String(args.get('intent') || base.intent || 'attention'),
    strokes: Array.isArray(base.strokes) ? base.strokes : [],
    payload: {
      ...(base.payload ?? {}),
      origin: 'device_mark_client',
      client: 'scripts/device-mark-client.mjs',
    },
  };
  const meetingSession = buildMeetingSession();
  if (meetingSession) payload.meeting_session = meetingSession;
  if (args.get('force-meeting-session') === 'true' || args.get('force_meeting_session') === 'true') {
    payload.force_meeting_session = true;
  }
  return payload;
}

async function postAnnotation(info = {}) {
  const endpoint = info.endpoint || `${baseUrl}/api/annotations`;
  const clockSync = await syncDeviceClock(info);
  const payload = buildPayload(info, clockSync);
  const result = await fetchJson(endpoint, {
    method: 'POST',
    headers: {
      'content-type': 'application/json; charset=utf-8',
      'x-hmp-device-id': deviceId,
      'x-hmp-device-type': deviceType,
    },
    body: JSON.stringify(payload),
  });
  return { result, clockSync };
}

async function getStatus(info = {}) {
  return fetchJson(statusUrlFromInfo(info));
}

function statusLine(prefix, status = {}, ack = null) {
  return [
    prefix,
    `id=${status.annotation_id ?? ack?.annotation_id ?? annotationId}`,
    `status=${status.status ?? 'unknown'}`,
    `accepted=${bool(status.accepted ?? ack?.accepted)}`,
    `on_real_axis=${bool(status.on_real_axis ?? ack?.on_real_axis)}`,
    `pending=${bool(status.pending_real_meeting ?? ack?.pending_real_meeting)}`,
    `time_source=${status.time_source ?? ack?.time_source ?? 'unknown'}`,
    `time_ms=${status.normalized_time_ms ?? ack?.normalized_time_ms ?? 'unknown'}`,
    status.binding_state || ack?.binding_state ? `binding=${status.binding_state ?? ack?.binding_state}` : '',
    status.next_action ? `next="${compactText(status.next_action)}"` : '',
  ].filter(Boolean).join(' | ');
}

async function waitForRealAxis(info = {}, initialStatus = null) {
  const startedAt = Date.now();
  let status = initialStatus ?? await getStatus(info);
  while (true) {
    if (status.status === 'real_axis_bound' || status.on_real_axis) return status;
    if (['needs_device_captured_at', 'after_meeting_end'].includes(status.status)) return status;
    if (timeoutMs > 0 && Date.now() - startedAt >= timeoutMs) return status;
    await new Promise((resolve) => setTimeout(resolve, intervalMs));
    status = await getStatus(info);
  }
}

try {
  const info = await loadIngestInfo();
  const posted = await postAnnotation(info);
  const result = posted.result;
  const clockSync = posted.clockSync;
  const ack = result.ack ?? {};
  const firstStatus = await getStatus(info);
  const finalStatus = waitRealAxis ? await waitForRealAxis(info, firstStatus) : firstStatus;
  if (json) {
    console.log(JSON.stringify({
      type: 'device_mark_result',
      annotation_id: annotationId,
      endpoint: info.endpoint,
      status_url: statusUrlFromInfo(info),
      stream_url: info.stream_url,
      clock_sync: clockSync,
      ack,
      meeting_session_binding: result.meeting_session_binding ?? null,
      passive_binding: result.passive_binding ?? null,
      status: finalStatus,
    }));
  } else {
    if (clockSync?.applied) console.log(`clock_sync=applied | offset_ms=${clockSync.offset_ms} | rtt_ms=${clockSync.rtt_ms}`);
    console.log(statusLine('device_mark_ack', firstStatus, ack));
    if (result.meeting_session_binding?.attempted) {
      console.log(`meeting_session=${result.meeting_session_binding.status || 'attempted'} | meeting_id=${result.meeting_session_binding.meeting_id || 'unknown'} | source=${result.meeting_session_binding.source || 'unknown'}`);
    }
    if (waitRealAxis) console.log(statusLine('device_mark_final', finalStatus, ack));
    if (info.stream_url) console.log(`stream_url=${info.stream_url}`);
  }
  if (waitRealAxis && !(finalStatus.status === 'real_axis_bound' || finalStatus.on_real_axis)) {
    throw new Error(`device mark did not bind to real axis: status=${finalStatus.status}`);
  }
} catch (error) {
  console.error(error.message ?? String(error));
  process.exitCode = 2;
}
