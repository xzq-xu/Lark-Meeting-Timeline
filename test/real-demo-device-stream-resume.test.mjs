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

function startServer(port, tempDir, output) {
  const child = spawn(process.execPath, ['src/server.mjs'], {
    cwd: new URL('..', import.meta.url),
    env: {
      ...process.env,
      PORT: String(port),
      LARK_WS_EVENTS: '0',
      LARK_APP_ID: 'real-demo-device-stream-resume-app',
      LARK_APP_SECRET: 'real-demo-device-stream-resume-secret',
      LARK_VERIFICATION_TOKEN: '',
      LARK_EVENT_CALLBACK_URL: 'https://real-demo-device-stream-resume.example.com/api/lark/events',
      TIMELINE_DATA_DIR: tempDir,
    },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  child.stdout.on('data', (chunk) => {
    output.text += chunk.toString();
  });
  child.stderr.on('data', (chunk) => {
    output.text += chunk.toString();
  });
  return child;
}

async function stopServer(child) {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
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

async function waitForJson(baseUrl, path, predicate, timeoutMs = 6_000) {
  const deadline = Date.now() + timeoutMs;
  let last = null;
  while (Date.now() < deadline) {
    const response = await fetch(`${baseUrl}${path}`);
    last = await response.json();
    assert.equal(response.ok, true, `${path} failed: ${JSON.stringify(last)}`);
    if (predicate(last)) return last;
    await new Promise((resolve) => setTimeout(resolve, 120));
  }
  throw new Error(`condition not observed for ${path}: ${JSON.stringify(last)}`);
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

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-real-demo-device-stream-resume-'));
const baseUrl = `http://127.0.0.1:${port}`;
const output = { text: '' };
let child = startServer(port, tempDir, output);

try {
  await waitForServer(baseUrl);

  await postJson(baseUrl, '/api/lark/real-demo/prepare', {
    auto_annotation: false,
    device_simulator: true,
    device_stream: true,
    passive_scan: false,
    device_stream_interval_ms: 300,
    device_stream_max_count: 2,
  });

  await waitForJson(baseUrl, '/api/device-simulator/stream', (stream) => (
    stream.enabled === true && stream.status === 'waiting_for_real_axis'
  ));

  await stopServer(child);
  child = startServer(port, tempDir, output);
  await waitForServer(baseUrl);

  const restored = await waitForJson(baseUrl, '/api/device-simulator/stream', (stream) => (
    stream.enabled === true && stream.status === 'waiting_for_real_axis'
  ));
  assert.equal(restored.timer_active, true);

  const startSeconds = 1_782_442_800;
  await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'device-stream-resume-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'device-stream-resume-real-meeting',
        topic: 'Device stream resume meeting',
        start_time: String(startSeconds),
      },
    },
  });

  const stateWithStreamMark = await waitForJson(baseUrl, '/api/state', (state) => (
    state.sequence?.some((item) => String(item.id).startsWith('device-stream-device-stream-resume-real-meeting-'))
  ), 8_000);
  const streamMark = stateWithStreamMark.sequence.find((item) => String(item.id).startsWith('device-stream-device-stream-resume-real-meeting-'));
  assert.equal(streamMark.source, 'hanwang_epaper_simulator');
  assert.equal(streamMark.time_source, 'captured_at');

  const status = await waitForJson(baseUrl, '/api/lark/real-demo/status', (body) => (
    body.gates.real_axis_annotation_count > 0
  ));
  assert.equal(status.gates.real_meeting_axis_active, true);
  assert.equal(status.gates.real_axis_annotation_count > 0, true);

  console.log('ok real demo device stream resumes after restart');
} catch (error) {
  console.error(output.text);
  throw error;
} finally {
  await stopServer(child).catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
