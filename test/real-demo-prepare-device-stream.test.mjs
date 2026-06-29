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

async function waitForJson(baseUrl, path, predicate, timeoutMs = 6_000) {
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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-real-demo-prepare-device-stream-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'real-demo-prepare-stream-app',
    LARK_APP_SECRET: 'real-demo-prepare-stream-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://real-demo-prepare-stream.example.com/api/lark/events',
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

  const prepared = await postJson(baseUrl, '/api/lark/real-demo/prepare', {
    auto_annotation: false,
    passive_scan: false,
    device_stream_interval_ms: 250,
    device_stream_max_count: 3,
  });
  assert.equal(prepared.prepared, true);
  assert.equal(prepared.device_stream_simulator.enabled, true);
  assert.equal(prepared.device_stream_simulator.max_count, 3);

  const waiting = await waitForJson(baseUrl, '/api/device-simulator/stream', (json) => json.status === 'waiting_for_real_axis');
  assert.equal(waiting.count, 0);

  const statusBefore = await getJson(baseUrl, '/api/lark/real-demo/status');
  assert.equal(statusBefore.gates.device_stream_enabled, true);
  assert.equal(statusBefore.gates.real_meeting_axis_active, false);

  const startSeconds = 1_782_442_800;
  await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'real-demo-prepare-stream-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'real-demo-prepare-stream-meeting',
        topic: 'Real demo prepare stream meeting',
        url: 'https://vc.feishu.cn/j/real-demo-prepare-stream',
        start_time: String(startSeconds),
      },
    },
  });

  const completedStream = await waitForJson(baseUrl, '/api/device-simulator/stream', (json) => (
    json.status === 'complete' && json.count === 3 && json.last_meeting_id === 'real-demo-prepare-stream-meeting'
  ));
  assert.equal(completedStream.enabled, false);

  const bindings = await getJson(baseUrl, '/api/annotation-bindings');
  const streamed = bindings.items.filter((item) => item.id.startsWith('device-stream-real-demo-prepare-stream-meeting-'));
  assert.equal(streamed.length, 3);
  assert.equal(streamed.every((item) => item.on_real_axis), true);

  const finalStatus = await getJson(baseUrl, '/api/lark/real-demo/status');
  assert.equal(finalStatus.complete, true);
  assert.equal(finalStatus.gates.real_axis_annotation_count, 4);

  console.log('ok real demo prepare starts device stream');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
