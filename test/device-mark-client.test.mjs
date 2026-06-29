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

async function runDeviceMark(baseUrl, args = []) {
  const child = spawn(process.execPath, [
    'scripts/device-mark-client.mjs',
    `--url=${baseUrl}`,
    '--json',
    ...args,
  ], {
    cwd: new URL('..', import.meta.url),
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  let stdout = '';
  let stderr = '';
  child.stdout.on('data', (chunk) => {
    stdout += chunk.toString();
  });
  child.stderr.on('data', (chunk) => {
    stderr += chunk.toString();
  });
  const [code] = await once(child, 'exit');
  return {
    code,
    stdout: stdout.trim(),
    stderr: stderr.trim(),
    json: stdout.trim() ? JSON.parse(stdout) : null,
  };
}

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-device-mark-client-'));
const baseUrl = `http://127.0.0.1:${port}`;
const server = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'device-mark-client-app',
    LARK_APP_SECRET: 'device-mark-client-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://device-mark-client.example.com/api/lark/events',
    TIMELINE_DATA_DIR: tempDir,
  },
  stdio: ['ignore', 'pipe', 'pipe'],
});

let serverOutput = '';
server.stdout.on('data', (chunk) => {
  serverOutput += chunk.toString();
});
server.stderr.on('data', (chunk) => {
  serverOutput += chunk.toString();
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
  const pending = await runDeviceMark(baseUrl, [
    '--id=device-client-pending',
    '--label=from device cli',
    `--captured-at-ms=${startSeconds * 1000 + 12_000}`,
  ]);
  assert.equal(pending.code, 0, pending.stderr);
  assert.equal(pending.json.type, 'device_mark_result');
  assert.equal(pending.json.annotation_id, 'device-client-pending');
  assert.equal(pending.json.ack.binding_state, 'pending_real_meeting');
  assert.equal(pending.json.status.status, 'pending_real_meeting');
  assert.equal(pending.json.status.on_real_axis, false);
  assert.match(pending.json.status_url, /\/api\/annotation-status\?id=device-client-pending$/);
  assert.match(pending.json.stream_url, /\/api\/stream$/);

  const pendingStatus = await getJson(baseUrl, '/api/annotation-status?id=device-client-pending');
  assert.equal(pendingStatus.status, 'pending_real_meeting');
  assert.equal(pendingStatus.pending_real_meeting, true);

  await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'device-client-real-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'device-client-real-meeting',
        topic: 'Device client real meeting',
        url: 'https://vc.feishu.cn/j/device-client-real',
        start_time: String(startSeconds),
      },
    },
  });

  const reboundStatus = await getJson(baseUrl, '/api/annotation-status?id=device-client-pending');
  assert.equal(reboundStatus.status, 'real_axis_bound');
  assert.equal(reboundStatus.on_real_axis, true);
  assert.equal(reboundStatus.normalized_time_ms, 12_000);
  assert.equal(reboundStatus.time_source, 'captured_at');

  const realAxis = await runDeviceMark(baseUrl, [
    '--id=device-client-real-axis',
    '--label=after axis',
    `--captured-at-ms=${startSeconds * 1000 + 18_000}`,
    '--wait-real-axis',
    '--timeout-ms=2000',
  ]);
  assert.equal(realAxis.code, 0, realAxis.stderr);
  assert.equal(realAxis.json.ack.binding_state, 'real_meeting_bound');
  assert.equal(realAxis.json.status.status, 'real_axis_bound');
  assert.equal(realAxis.json.status.on_real_axis, true);
  assert.equal(realAxis.json.status.normalized_time_ms, 18_000);

  const syncedClock = await runDeviceMark(baseUrl, [
    '--id=device-client-clock-sync',
    '--label=clock sync default',
  ]);
  assert.equal(syncedClock.code, 0, syncedClock.stderr);
  assert.equal(syncedClock.json.clock_sync.applied, true);
  assert.equal(typeof syncedClock.json.clock_sync.offset_ms, 'number');
  assert.equal(typeof syncedClock.json.clock_sync.rtt_ms, 'number');
  assert.equal(syncedClock.json.ack.time_source, 'captured_at');
  assert.equal(syncedClock.json.status.status, 'real_axis_bound');

  const inlineStartMs = startSeconds * 1000 + 120_000;
  const inlineSession = await runDeviceMark(baseUrl, [
    '--id=device-client-inline-session',
    '--label=inline session mark',
    `--captured-at-ms=${inlineStartMs + 7000}`,
    '--meeting-session=true',
    '--meeting-id=device-client-inline-meeting',
    '--meeting-title=Device client inline meeting',
    '--meeting-url=https://vc.feishu.cn/j/device-client-inline',
    `--meeting-start-ms=${inlineStartMs}`,
    '--force-meeting-session=true',
    '--wait-real-axis',
    '--timeout-ms=2000',
  ]);
  assert.equal(inlineSession.code, 0, inlineSession.stderr);
  assert.equal(inlineSession.json.meeting_session_binding.attempted, true);
  assert.equal(inlineSession.json.meeting_session_binding.status, 'started');
  assert.equal(inlineSession.json.meeting_session_binding.meeting_id, 'device-client-inline-meeting');
  assert.equal(inlineSession.json.ack.binding_state, 'real_meeting_bound');
  assert.equal(inlineSession.json.ack.meeting_id, 'device-client-inline-meeting');
  assert.equal(inlineSession.json.status.status, 'real_axis_bound');
  assert.equal(inlineSession.json.status.normalized_time_ms, 7000);

  const inlineState = await getJson(baseUrl, '/api/state');
  assert.equal(inlineState.meeting.source, 'open_meeting_session');
  assert.equal(inlineState.meeting.meeting_id, 'device-client-inline-meeting');

  console.log('ok device mark client');
} catch (error) {
  console.error(serverOutput);
  throw error;
} finally {
  server.kill('SIGTERM');
  await once(server, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
