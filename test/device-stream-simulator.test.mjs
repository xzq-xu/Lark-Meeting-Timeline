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

async function waitForSseState(baseUrl, predicate, timeoutMs = 7_000) {
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

async function waitForJson(baseUrl, path, predicate, timeoutMs = 5_000) {
  const deadline = Date.now() + timeoutMs;
  let last = null;
  while (Date.now() < deadline) {
    last = await getJson(baseUrl, path);
    if (predicate(last)) return last;
    await new Promise((resolve) => setTimeout(resolve, 80));
  }
  throw new Error(`condition not met for ${path}: ${JSON.stringify(last)}`);
}

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-device-stream-simulator-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'device-stream-app',
    LARK_APP_SECRET: 'device-stream-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://device-stream.example.com/api/lark/events',
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

  const startedStream = await postJson(baseUrl, '/api/device-simulator/stream', {
    action: 'start',
    interval_ms: 250,
    max_count: 3,
    label_prefix: '测试流式标注',
  });
  assert.equal(startedStream.enabled, true);
  assert.equal(startedStream.max_count, 3);

  const waiting = await waitForJson(baseUrl, '/api/device-simulator/stream', (json) => json.status === 'waiting_for_real_axis');
  assert.equal(waiting.count, 0);

  const startSeconds = 1_782_442_800;
  const ssePromise = waitForSseState(baseUrl, (state) => (
    state.meeting?.meeting_id === 'device-stream-real-meeting'
      && state.sequence?.filter((item) => item.id?.startsWith('device-stream-device-stream-real-meeting-')).length === 3
  ));

  await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'device-stream-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'device-stream-real-meeting',
        topic: 'Device stream real meeting',
        url: 'https://vc.feishu.cn/j/device-stream',
        start_time: String(startSeconds),
      },
    },
  });

  const sseState = await ssePromise;
  const streamed = sseState.sequence.filter((item) => item.id.startsWith('device-stream-device-stream-real-meeting-'));
  assert.equal(streamed.length, 3);
  assert.deepEqual(streamed.map((item) => item.source), [
    'hanwang_epaper_simulator',
    'hanwang_epaper_simulator',
    'hanwang_epaper_simulator',
  ]);
  assert.equal(streamed.every((item) => item.time_source === 'captured_at'), true);

  const finalStream = await waitForJson(baseUrl, '/api/device-simulator/stream', (json) => json.status === 'complete' && json.count === 3);
  assert.equal(finalStream.enabled, false);
  assert.equal(finalStream.last_meeting_id, 'device-stream-real-meeting');

  const bindings = await getJson(baseUrl, '/api/annotation-bindings');
  assert.equal(bindings.binding_state, 'real_meeting_bound');
  assert.equal(bindings.real_axis_count, 3);

  console.log('ok device stream simulator');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
