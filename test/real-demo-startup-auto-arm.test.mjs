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

async function waitForJson(baseUrl, path, predicate, timeoutMs = 8_000) {
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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-real-demo-startup-auto-arm-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'real-demo-startup-auto-arm-app',
    LARK_APP_SECRET: 'real-demo-startup-auto-arm-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://real-demo-startup-auto-arm.example.com/api/lark/events',
    TIMELINE_DATA_DIR: tempDir,
    REAL_DEMO_AUTO_ARM: '1',
    REAL_DEMO_AUTO_ANNOTATION: '0',
    REAL_DEMO_DEVICE_SIMULATOR: '1',
    REAL_DEMO_DEVICE_STREAM: '1',
    REAL_DEMO_DEVICE_STREAM_INTERVAL_MS: '250',
    REAL_DEMO_DEVICE_STREAM_COUNT: '2',
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

  const armed = await waitForJson(baseUrl, '/api/lark/real-demo/status', (json) => (
    json.real_demo_session?.active === true
      && json.device_simulator?.stream?.status === 'waiting_for_real_axis'
  ));
  assert.equal(armed.gates.real_meeting_axis_active, false);
  assert.equal(armed.gates.scan_fallback_ready, false);
  assert.equal(armed.passive_meeting_scan.enabled, true);
  assert.equal(armed.passive_meeting_scan.tenant_fallback_enabled, false);
  assert.equal(armed.device_simulator.stream.max_count, 2);

  const startSeconds = Math.floor(Date.now() / 1000) - 5;
  const started = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'startup-auto-arm-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'startup-auto-arm-meeting',
        topic: 'Startup auto-arm direct meeting',
        url: 'https://vc.feishu.cn/j/startup-auto-arm',
        start_time: String(startSeconds),
      },
    },
  });
  assert.equal(started.ok, true);
  assert.equal(started.timeline_started, true);
  assert.equal(started.state.meeting.source, 'lark_http_event');

  const stream = await waitForJson(baseUrl, '/api/device-simulator/stream', (json) => (
    json.status === 'complete' && json.count === 2 && json.last_meeting_id === 'startup-auto-arm-meeting'
  ));
  assert.equal(stream.enabled, false);

  const bindings = await getJson(baseUrl, '/api/annotation-bindings');
  const streamed = bindings.items.filter((item) => item.id.startsWith('device-stream-startup-auto-arm-meeting-'));
  assert.equal(streamed.length, 2);
  assert.equal(streamed.every((item) => item.on_real_axis), true);

  const progress = await getJson(baseUrl, '/api/lark/real-demo/progress');
  assert.equal(progress.completion_evidence.real_meeting_axis_active, true);
  assert.equal(progress.completion_evidence.real_axis_annotation_count >= 2, true);
  assert.equal(progress.completion_evidence.device_stream_status, 'complete');

  console.log('ok real demo startup auto-arm direct meeting annotations');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
