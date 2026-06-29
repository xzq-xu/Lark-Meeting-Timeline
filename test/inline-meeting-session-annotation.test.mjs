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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-inline-meeting-session-'));
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

  const info = await getJson(baseUrl, '/api/annotation-ingest-info');
  assert.equal(info.meeting_session_inline_annotation.supported, true);
  assert.match(info.device_client_contract.meeting_session_start_condition, /include meeting_session/);

  const startMs = 1_782_542_800_000;
  const ssePromise = waitForSseState(baseUrl, (state) => (
    state.meeting?.source === 'open_meeting_session'
      && state.meeting?.meeting_id === 'inline-session-meeting-001'
      && state.sequence?.some((item) => item.id === 'inline-session-ann-1' && item.time_ms === 8000)
  ));

  const result = await postJson(baseUrl, '/api/annotations', {
    id: 'inline-session-ann-1',
    source: 'hanwang_epaper',
    device_id: 'hanwang-real-device-001',
    captured_at_ms: startMs + 8000,
    kind: 'handwriting_trigger',
    label: 'why?',
    text_candidates: ['why?', 'why'],
    intent: 'question',
    meeting_session: {
      platform: 'lark',
      meeting_id: 'inline-session-meeting-001',
      title: 'Inline meeting session',
      meeting_url: 'https://vc.feishu.cn/j/inline-session-meeting-001',
      start_time_ms: startMs,
      detector_source: 'hanwang_host_app',
    },
  });

  assert.equal(result.meeting_session_binding.attempted, true);
  assert.equal(result.meeting_session_binding.status, 'started');
  assert.equal(result.ack.binding_state, 'real_meeting_bound');
  assert.equal(result.ack.on_real_axis, true);
  assert.equal(result.ack.meeting_id, 'inline-session-meeting-001');
  assert.equal(result.ack.normalized_time_ms, 8000);
  assert.equal(result.state.meeting.source, 'open_meeting_session');
  assert.equal(result.state.meeting.start_time_source, 'operator_supplied_start_time');

  const sseState = await ssePromise;
  assert.equal(sseState.meeting.source, 'open_meeting_session');

  const second = await postJson(baseUrl, '/api/annotations', {
    id: 'inline-session-ann-2',
    source: 'hanwang_epaper',
    captured_at_ms: startMs + 12_000,
    kind: 'attention',
    label: '重点',
    meeting_session: {
      platform: 'lark',
      meeting_id: 'inline-session-meeting-001',
      start_time_ms: startMs,
    },
  });
  assert.equal(second.meeting_session_binding.status, 'already_bound');
  assert.equal(second.ack.binding_state, 'real_meeting_bound');
  assert.equal(second.ack.normalized_time_ms, 12_000);

  const acceptance = await getJson(baseUrl, '/api/lark/real-demo/acceptance');
  assert.equal(acceptance.product_acceptance_complete, true);
  assert.equal(acceptance.strict_event_acceptance_complete, false);

  console.log('ok inline meeting_session annotation starts axis and lands mark');
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
  if (process.exitCode) {
    console.error(output);
  }
}
