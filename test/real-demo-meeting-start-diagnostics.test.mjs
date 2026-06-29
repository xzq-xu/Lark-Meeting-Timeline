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

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-real-demo-meeting-start-diagnostics-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'real-demo-meeting-start-diagnostics-app',
    LARK_APP_SECRET: 'real-demo-meeting-start-diagnostics-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://real-demo-meeting-start-diagnostics.example.com/api/lark/events',
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
  });

  const startSeconds = 1_782_442_800;
  const started = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'meeting-start-diagnostics-start',
      event_type: 'vc.meeting.meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting_id: 'meeting-start-diagnostics-real-meeting',
      topic: 'Meeting started diagnostics',
      meeting_url: 'https://vc.feishu.cn/j/meeting-start-diagnostics',
      start_time: String(startSeconds),
    },
  });

  assert.equal(started.ok, true);
  assert.equal(started.timeline_started, true);
  assert.equal(started.state.meeting.meeting_id, 'meeting-start-diagnostics-real-meeting');
  assert.equal(started.state.meeting.source, 'lark_http_event');

  const diagnostics = await getJson(baseUrl, '/api/lark/delivery-diagnostics');
  assert.equal(diagnostics.status, 'meeting_start_event_received');
  assert.equal(diagnostics.evidence.real_meeting_axis_active, true);
  assert.equal(diagnostics.evidence.reserve_start_event.event_type, 'vc.meeting.meeting_started_v1');
  assert.equal(diagnostics.evidence.start_event.event_type, 'vc.meeting.meeting_started_v1');

  const readiness = await getJson(baseUrl, '/api/readiness');
  const entryCheck = readiness.checks.find((check) => check.id === 'direct_meeting_event_seen');
  assert.equal(entryCheck.ok, true);
  assert.match(entryCheck.detail, /vc\.meeting\.meeting_started_v1/);
  const realEntryCheck = readiness.checks.find((check) => check.id === 'real_meeting_entry');
  assert.equal(realEntryCheck.ok, true);

  const realDemoStatus = await getJson(baseUrl, '/api/lark/real-demo/status');
  assert.equal(realDemoStatus.status, 'real_axis_waiting_annotation');
  assert.equal(realDemoStatus.gates.real_event_seen, true);
  assert.equal(realDemoStatus.gates.real_meeting_axis_active, true);
  assert.equal(realDemoStatus.complete, false);

  console.log('ok real demo meeting_started diagnostics');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
