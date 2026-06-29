import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { once } from 'node:events';
import net from 'node:net';

async function freePort() {
  const server = net.createServer();
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  const address = server.address();
  server.close();
  await once(server, 'close');
  return address.port;
}

async function waitForServer(baseUrl, timeoutMs = 10_000) {
  const deadline = Date.now() + timeoutMs;
  let lastError = null;
  while (Date.now() < deadline) {
    try {
      const response = await fetch(`${baseUrl}/api/state`);
      if (response.ok) return;
      lastError = new Error(`HTTP ${response.status}`);
    } catch (error) {
      lastError = error;
    }
    await new Promise((resolve) => setTimeout(resolve, 80));
  }
  throw lastError ?? new Error('server did not start');
}

async function postJson(baseUrl, path, body) {
  const response = await fetch(`${baseUrl}${path}`, {
    method: 'POST',
    headers: { 'content-type': 'application/json; charset=utf-8' },
    body: JSON.stringify(body),
  });
  const json = await response.json();
  assert.equal(response.ok, true, `${path} failed: ${JSON.stringify(json)}`);
  return json;
}

async function getJson(baseUrl, path) {
  const response = await fetch(`${baseUrl}${path}`);
  const json = await response.json();
  assert.equal(response.ok, true, `${path} failed: ${JSON.stringify(json)}`);
  return json;
}

async function waitForSseState(baseUrl, predicate, timeoutMs = 5_000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(`${baseUrl}/api/stream`, { signal: controller.signal });
    assert.equal(response.ok, true, 'SSE stream should open');
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const chunks = buffer.split('\n\n');
      buffer = chunks.pop() ?? '';
      for (const chunk of chunks) {
        const event = chunk.split('\n').find((line) => line.startsWith('event: '))?.slice(7);
        const data = chunk.split('\n').find((line) => line.startsWith('data: '))?.slice(6);
        if (event === 'state' && data) {
          const state = JSON.parse(data);
          if (predicate(state)) {
            await reader.cancel();
            return state;
          }
        }
      }
    }
  } finally {
    clearTimeout(timer);
  }
  throw new Error('SSE state predicate was not observed');
}

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-open-meeting-session-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    TIMELINE_DATA_DIR: tempDir,
  },
  stdio: ['ignore', 'pipe', 'pipe'],
});

let output = '';
child.stdout.on('data', (chunk) => {
  output += chunk.toString();
});
child.stderr.on('data', (chunk) => {
  output += chunk.toString();
});

try {
  await waitForServer(baseUrl);

  const startMs = 1_782_442_800_000;
  const session = await postJson(baseUrl, '/api/meeting-session/start', {
    platform: 'lark',
    meeting_id: 'open-session-real-lark-001',
    title: 'Open meeting session acceptance',
    meeting_url: 'https://vc.feishu.cn/j/open-session-real-lark-001',
    start_time_ms: startMs,
    detector_source: 'desktop_meeting_observer',
  });
  assert.equal(session.ok, true);
  assert.equal(session.meeting.source, 'open_meeting_session');
  assert.equal(session.meeting.start_time_source, 'operator_supplied_start_time');
  assert.equal(session.contract.strict_lark_event_axis, false);
  assert.equal(session.contract.product_axis, true);

  const status = await getJson(baseUrl, '/api/meeting-session/status');
  assert.equal(status.real_meeting_axis_active, true);
  assert.equal(status.strict_lark_event_axis, false);

  const ssePromise = waitForSseState(baseUrl, (state) => (
    state.meeting?.source === 'open_meeting_session'
      && state.sequence?.some((item) => item.id === 'open-session-ann-1' && item.time_ms === 12_000)
  ));

  const annotation = await postJson(baseUrl, '/api/annotations', {
    id: 'open-session-ann-1',
    source: 'hanwang_epaper',
    device_id: 'hanwang-real-device-001',
    captured_at_ms: startMs + 12_000,
    kind: 'handwriting_trigger',
    label: 'why?',
    text_candidates: ['why?', 'why'],
    intent: 'question',
    strokes: [
      [{ x: 0.1, y: 0.2, t: startMs + 11_800 }],
      [{ x: 0.2, y: 0.2, t: startMs + 12_000 }],
    ],
  });
  assert.equal(annotation.ack.accepted, true);
  assert.equal(annotation.ack.binding_state, 'real_meeting_bound');
  assert.equal(annotation.ack.on_real_axis, true);
  assert.equal(annotation.ack.meeting_id, 'open-session-real-lark-001');
  assert.equal(annotation.ack.normalized_time_ms, 12_000);

  const sseState = await ssePromise;
  assert.equal(sseState.meeting.source, 'open_meeting_session');

  const acceptance = await getJson(baseUrl, '/api/lark/real-demo/acceptance');
  assert.equal(acceptance.product_acceptance_complete, true);
  assert.equal(acceptance.strict_event_acceptance_complete, false);
  assert.equal(acceptance.meeting_axis.source, 'open_meeting_session');
  assert.equal(acceptance.realtime_annotation.real_axis_annotation_count, 1);

  const ended = await postJson(baseUrl, '/api/meeting-session/end', {
    end_time_ms: startMs + 60_000,
  });
  assert.equal(ended.ok, true);
  assert.equal(ended.state.meeting.end_time, new Date(startMs + 60_000).toISOString());
  assert.equal(ended.state.events.some((event) => event.type === 'meeting_end' && event.source === 'open_meeting_session'), true);

  console.log('ok open meeting session protocol binds realtime annotations');
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
  if (process.exitCode) {
    console.error(output);
  }
}
