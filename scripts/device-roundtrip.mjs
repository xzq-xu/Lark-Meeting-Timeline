#!/usr/bin/env node

const args = new Map();
for (const raw of process.argv.slice(2)) {
  const [key, ...rest] = raw.replace(/^--/, '').split('=');
  args.set(key, rest.length ? rest.join('=') : 'true');
}

const baseUrl = String(args.get('url') || process.env.DEVICE_MARK_BASE_URL || process.env.REAL_DEMO_BASE_URL || 'http://localhost:8787').replace(/\/$/, '');
const json = args.get('json') === 'true';
const timeoutMs = Math.min(Math.max(Number(args.get('timeout-ms') ?? 10_000), 500), 120_000);
const now = Date.now();
const annotationId = String(args.get('id') || args.get('annotation-id') || `device-roundtrip-${now}`);
const label = String(args.get('label') || 'device roundtrip');
const deviceId = String(args.get('device-id') || 'hanwang-roundtrip-001');
const deviceType = String(args.get('device-type') || 'hanwang_epaper');

function bool(value) {
  return value ? 'yes' : 'no';
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

async function syncClock(info = {}) {
  const endpoint = info.clock_sync?.endpoint || info.time_sync_url || `${baseUrl}/api/time`;
  const clientSendAtMs = Date.now();
  const separator = endpoint.includes('?') ? '&' : '?';
  const body = await fetchJson(`${endpoint}${separator}client_send_at_ms=${clientSendAtMs}`);
  const clientReceiveAtMs = Date.now();
  const serverTimeMs = Number(body.server_time_ms);
  if (!Number.isFinite(serverTimeMs)) return { applied: false, offset_ms: 0, reason: 'server_time_missing' };
  return {
    applied: true,
    offset_ms: Math.round(serverTimeMs - ((clientSendAtMs + clientReceiveAtMs) / 2)),
    rtt_ms: clientReceiveAtMs - clientSendAtMs,
  };
}

function statusUrlFromInfo(info = {}) {
  const template = info.annotation_status_url || `${baseUrl}/api/annotation-status?id={annotation_id}`;
  return template.replace('{annotation_id}', encodeURIComponent(annotationId));
}

function buildPayload(clockSync = {}) {
  return {
    id: annotationId,
    source: 'hanwang_epaper_roundtrip',
    captured_at_ms: Math.round(Date.now() + (clockSync.offset_ms ?? 0)),
    kind: 'handwriting_trigger',
    label,
    text_candidates: [label],
    intent: 'roundtrip_check',
    strokes: [],
    payload: {
      origin: 'device_roundtrip',
      client: 'scripts/device-roundtrip.mjs',
    },
  };
}

async function postAnnotation(info = {}, clockSync = {}) {
  return fetchJson(info.endpoint || `${baseUrl}/api/annotations`, {
    method: 'POST',
    headers: {
      'content-type': 'application/json; charset=utf-8',
      'x-hmp-device-id': deviceId,
      'x-hmp-device-type': deviceType,
    },
    body: JSON.stringify(buildPayload(clockSync)),
  });
}

function waitForSseAnnotation(info = {}) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(new Error(`timeout after ${timeoutMs}ms`)), timeoutMs);
  let readyResolve;
  let readyReject;
  const ready = new Promise((resolve, reject) => {
    readyResolve = resolve;
    readyReject = reject;
  });
  const done = (async () => {
    let reader = null;
    try {
      const response = await fetch(info.stream_url || `${baseUrl}/api/stream`, { signal: controller.signal });
      if (!response.ok) throw new Error(`SSE stream HTTP ${response.status}`);
      readyResolve();
      reader = response.body.getReader();
      const decoder = new TextDecoder();
      let buffer = '';
      for (;;) {
        const { done: streamDone, value } = await reader.read();
        if (streamDone) throw new Error('SSE stream ended before annotation was observed');
        buffer += decoder.decode(value, { stream: true });
        let splitAt = buffer.indexOf('\n\n');
        while (splitAt >= 0) {
          const block = buffer.slice(0, splitAt);
          buffer = buffer.slice(splitAt + 2);
          const event = block.split('\n').find((line) => line.startsWith('event: '))?.slice(7);
          if (event === 'state') {
            const data = block
              .split('\n')
              .filter((line) => line.startsWith('data: '))
              .map((line) => line.slice(6))
              .join('\n');
            if (data) {
              const state = JSON.parse(data);
              const item = state.sequence?.find((row) => row.id === annotationId);
              if (item) {
                return { observed: true, state, item };
              }
            }
          }
          splitAt = buffer.indexOf('\n\n');
        }
      }
    } catch (error) {
      readyReject(error);
      throw error;
    } finally {
      clearTimeout(timer);
      if (reader) await reader.cancel().catch(() => {});
    }
  })();
  return { ready, done };
}

try {
  const info = await loadIngestInfo();
  const clockSync = await syncClock(info);
  const sse = waitForSseAnnotation(info);
  await sse.ready;
  const result = await postAnnotation(info, clockSync);
  const sseResult = await sse.done;
  const status = await fetchJson(statusUrlFromInfo(info));
  const summary = {
    type: 'device_roundtrip_result',
    ok: Boolean(result.ack?.accepted && sseResult.observed),
    annotation_id: annotationId,
    endpoint: info.endpoint,
    stream_url: info.stream_url,
    status_url: statusUrlFromInfo(info),
    clock_sync: clockSync,
    ack: result.ack,
    status,
    sse: {
      observed: sseResult.observed,
      item_id: sseResult.item.id,
      item_label: sseResult.item.label,
      item_time_ms: sseResult.item.time_ms,
      meeting_id: sseResult.state.meeting?.meeting_id ?? null,
    },
  };
  if (json) console.log(JSON.stringify(summary));
  else {
    console.log([
      'device_roundtrip',
      `ok=${bool(summary.ok)}`,
      `id=${summary.annotation_id}`,
      `binding=${summary.ack?.binding_state ?? 'unknown'}`,
      `sse=${bool(summary.sse.observed)}`,
      `time_ms=${summary.sse.item_time_ms ?? 'unknown'}`,
      `status=${summary.status.status}`,
    ].join(' | '));
  }
  if (!summary.ok) process.exitCode = 2;
} catch (error) {
  console.error(error.message ?? String(error));
  process.exitCode = 2;
}
