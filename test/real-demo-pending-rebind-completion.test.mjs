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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-real-demo-pending-rebind-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'real-demo-pending-rebind-app',
    LARK_APP_SECRET: 'real-demo-pending-rebind-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://real-demo-pending-rebind.example.com/api/lark/events',
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

  await postJson(baseUrl, '/api/lark/real-demo/prepare', {
    auto_annotation: false,
    device_simulator: false,
    device_stream: false,
    passive_scan: false,
    auto_open_session_on_annotation: false,
  });

  const startSeconds = 1_782_442_800;
  const pending = await postJson(baseUrl, '/api/annotations', {
    id: 'pending-before-real-start',
    source: 'future_epaper_device',
    device_id: 'hanwang-devkit-pending',
    captured_at_ms: startSeconds * 1000 + 31_000,
    kind: 'handwriting_trigger',
    label: '会中先写下的 why?',
    text_candidates: ['why?', 'why'],
    intent: 'question',
    strokes: [
      [
        { x: 0.1, y: 0.2, t: startSeconds * 1000 + 30_800 },
        { x: 0.2, y: 0.24, t: startSeconds * 1000 + 31_000 },
      ],
    ],
  });

  assert.equal(pending.ack.binding_state, 'pending_real_meeting');
  assert.equal(pending.ack.pending_real_meeting, true);
  assert.equal(pending.ack.created_pending_timeline, true);
  assert.equal(pending.state.meeting.source, 'annotation_fallback');
  assert.equal(pending.state.meeting.pending_binding, true);
  assert.equal(pending.state.sequence.find((item) => item.id === 'pending-before-real-start')?.time_ms, 0);

  const pendingStatus = await getJson(baseUrl, '/api/lark/real-demo/status');
  assert.equal(pendingStatus.complete, false);
  assert.equal(pendingStatus.gates.real_meeting_axis_active, false);
  assert.equal(pendingStatus.evidence.annotation_binding.pending_count, 1);

  const ssePromise = waitForSseState(baseUrl, (state) => (
    state.meeting?.meeting_id === 'pending-rebind-real-meeting'
      && state.sequence?.some((item) => (
        item.id === 'pending-before-real-start'
          && item.source === 'future_epaper_device'
          && item.time_ms === 31_000
          && item.time_source === 'captured_at'
      ))
  ));

  const started = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'pending-rebind-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'pending-rebind-real-meeting',
        topic: 'Real demo pending rebind meeting',
        url: 'https://vc.feishu.cn/j/pending-rebind',
        start_time: String(startSeconds),
      },
    },
  });

  assert.equal(started.ok, true);
  assert.equal(started.timeline_started, true);
  assert.equal(started.state.meeting.source, 'lark_http_event');
  assert.equal(started.state.meeting.pending_binding, false);
  assert.equal(started.state.sequence.find((item) => item.id === 'pending-before-real-start')?.time_ms, 31_000);

  const sseState = await ssePromise;
  assert.equal(sseState.meeting.meeting_id, 'pending-rebind-real-meeting');

  const completed = await getJson(baseUrl, '/api/lark/real-demo/status');
  assert.equal(completed.status, 'complete');
  assert.equal(completed.complete, true);
  assert.equal(completed.gates.real_meeting_axis_active, true);
  assert.equal(completed.gates.real_axis_annotation_count, 1);
  assert.equal(completed.evidence.annotation_binding.items[0].id, 'pending-before-real-start');
  assert.equal(completed.evidence.annotation_binding.items[0].on_real_axis, true);
  assert.equal(completed.evidence.annotation_binding.items[0].meeting_id, 'pending-rebind-real-meeting');

  const monitor = await postJson(baseUrl, '/api/lark/real-demo/monitor', {
    timeout_ms: 1000,
    interval_ms: 100,
  });
  assert.equal(monitor.observed, true);
  assert.equal(monitor.completion_evidence.real_demo_complete, true);
  assert.equal(monitor.completion_evidence.real_meeting_axis_active, true);
  assert.equal(monitor.completion_evidence.real_axis_annotation_count, 1);

  console.log('ok real demo pending annotation rebind completion');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
