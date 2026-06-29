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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-real-demo-open-annotation-sse-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'real-demo-open-annotation-app',
    LARK_APP_SECRET: 'real-demo-open-annotation-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://real-demo-open-annotation.example.com/api/lark/events',
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
    passive_scan: false,
  });

  const startSeconds = 1_782_442_800;
  const started = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'real-demo-open-annotation-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'real-demo-open-annotation-meeting',
        topic: 'Real demo open annotation meeting',
        url: 'https://vc.feishu.cn/j/real-demo-open-annotation',
        start_time: String(startSeconds),
      },
    },
  });
  assert.equal(started.ok, true);
  assert.equal(started.timeline_started, true);
  assert.equal(started.auto_acceptance_annotation, null);
  assert.equal(started.device_simulator_annotation, null);

  const waitingStatus = await getJson(baseUrl, '/api/lark/real-demo/status');
  assert.equal(waitingStatus.status, 'real_axis_waiting_annotation');
  assert.equal(waitingStatus.complete, false);
  assert.equal(waitingStatus.gates.real_meeting_axis_active, true);
  assert.equal(waitingStatus.gates.real_axis_annotation_count, 0);

  const ssePromise = waitForSseState(baseUrl, (state) => (
    state.meeting?.meeting_id === 'real-demo-open-annotation-meeting'
      && state.sequence?.some((item) => (
        item.id === 'external-open-ann-1'
          && item.source === 'future_epaper_device'
          && item.time_ms === 42_000
      ))
  ));

  const annotation = await postJson(baseUrl, '/api/annotations', {
    id: 'external-open-ann-1',
    source: 'future_epaper_device',
    device_id: 'hanwang-devkit-001',
    captured_at_ms: startSeconds * 1000 + 42_000,
    kind: 'handwriting_trigger',
    label: '手写 why?',
    text_candidates: ['why?', 'why'],
    intent: 'question',
    strokes: [
      [
        { x: 0.11, y: 0.21, t: startSeconds * 1000 + 41_900 },
        { x: 0.15, y: 0.22, t: startSeconds * 1000 + 42_000 },
      ],
    ],
  });
  assert.equal(annotation.ack.accepted, true);
  assert.equal(annotation.ack.binding_state, 'real_meeting_bound');
  assert.equal(annotation.ack.on_real_axis, true);
  assert.equal(annotation.ack.meeting_id, 'real-demo-open-annotation-meeting');
  assert.equal(annotation.ack.normalized_time_ms, 42_000);
  assert.equal(annotation.item.time_source, 'captured_at');

  const sseState = await ssePromise;
  const sseItem = sseState.sequence.find((item) => item.id === 'external-open-ann-1');
  assert.equal(sseItem.label, '手写 why?');
  assert.equal(sseItem.payload.device.id, 'hanwang-devkit-001');

  const completed = await getJson(baseUrl, '/api/lark/real-demo/status');
  assert.equal(completed.status, 'complete');
  assert.equal(completed.complete, true);
  assert.equal(completed.gates.real_axis_annotation_count, 1);
  assert.equal(completed.evidence.annotation_binding.items[0].source, 'future_epaper_device');

  const streamStatus = await getJson(baseUrl, '/api/stream/status');
  assert.ok(streamStatus.last_broadcast_at);
  assert.equal(streamStatus.last_event, 'state');
  assert.equal(streamStatus.last_state_summary.real_meeting_axis_active, true);
  assert.equal(streamStatus.last_state_summary.real_axis_annotation_count, 1);

  const monitor = await postJson(baseUrl, '/api/lark/real-demo/monitor', {
    timeout_ms: 1000,
    interval_ms: 100,
  });
  assert.equal(monitor.observed, true);
  assert.equal(monitor.completion_evidence.real_demo_complete, true);
  assert.equal(monitor.completion_evidence.real_meeting_axis_active, true);
	  assert.equal(monitor.completion_evidence.real_axis_annotation_count, 1);
	  assert.equal(monitor.completion_evidence.last_broadcast_real_axis_annotation_count, 1);

	  await postJson(baseUrl, '/api/lark/real-demo/prepare', {
	    auto_annotation: false,
	    device_simulator: true,
	    device_stream: true,
	    passive_scan: false,
	  });
	  await new Promise((resolve) => setTimeout(resolve, 1800));
	  const freshStatus = await getJson(baseUrl, '/api/lark/real-demo/status');
	  assert.equal(freshStatus.complete, false);
	  assert.equal(freshStatus.gates.real_meeting_axis_active, false);
	  assert.equal(freshStatus.gates.real_axis_annotation_count, 0);
	  assert.equal(freshStatus.gates.real_axis_annotation_count_total, 1);
	  assert.equal(freshStatus.device_stream_simulator.status, 'waiting_for_real_axis');

	  console.log('ok real demo open annotation SSE');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
