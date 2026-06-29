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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-restore-real-axis-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'restore-real-axis-app',
    LARK_APP_SECRET: 'restore-real-axis-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://restore-real-axis.example.com/api/lark/events',
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

  const startSeconds = 1_782_711_012;
  const started = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'restore-axis-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'restore-axis-real-meeting',
        meeting_no: '152898256',
        topic: 'Restore latest real axis',
        start_time: String(startSeconds),
      },
    },
  });
  assert.equal(started.ok, true);
  assert.equal(started.timeline_started, true);
  assert.equal(started.state.meeting.source, 'lark_http_event');

  await postJson(baseUrl, '/api/demo/reset', {});

  const ended = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'restore-axis-end',
      event_type: 'vc.meeting.all_meeting_ended_v1',
      create_time: String(startSeconds + 40),
    },
    event: {
      meeting: {
        id: 'restore-axis-real-meeting',
        meeting_no: '152898256',
        topic: 'Restore latest real axis',
        start_time: String(startSeconds),
        end_time: String(startSeconds + 40),
      },
    },
  });
  assert.equal(ended.ok, false);
  assert.equal(ended.ignored_reason, 'meeting_end_without_known_start_axis');

  const probe = await postJson(baseUrl, '/api/lark/real-meeting-probe/start', {
    timeout_ms: 120_000,
    reset_temporary_axis: true,
    auto_search: false,
  });
  assert.equal(probe.status, 'waiting');
  assert.equal(probe.missed_event.reason, 'meeting_start_before_probe');
  assert.equal(probe.missed_event.start_event.event_type, 'vc.meeting.all_meeting_started_v1');
  assert.equal(probe.missed_event.end_event.event_type, 'vc.meeting.all_meeting_ended_v1');

  const currentBeforeRestore = await getJson(baseUrl, '/api/state');
  assert.equal(currentBeforeRestore.meeting.meeting_id, 'demo-lark-meeting-001');

  await postJson(baseUrl, '/api/acceptance/auto-annotation', {
    enabled: true,
    label: 'should not append during restore',
  });
  await postJson(baseUrl, '/api/device-simulator', {
    enabled: true,
    label: 'device should not append during restore',
  });

  const restored = await postJson(baseUrl, '/api/lark/restore-latest-real-meeting-axis', {});
  assert.equal(restored.restored, true);
  assert.equal(restored.start_result.ok, true);
  assert.equal(restored.end_result.ok, true);
  assert.equal(restored.state.meeting.meeting_id, 'restore-axis-real-meeting');
  assert.equal(restored.state.meeting.external_meeting_id, '152898256');
  assert.equal(restored.state.meeting.source, 'lark_http_event');
  assert.equal(restored.state.meeting.start_time, '2026-06-29T05:30:12.000Z');
  assert.equal(restored.state.meeting.end_time, '2026-06-29T05:30:52.000Z');
  assert.equal(restored.state.events.some((event) => event.type === 'meeting_start'), true);
  assert.equal(restored.state.events.some((event) => event.type === 'meeting_end'), true);
  assert.equal(restored.state.sequence.length, 0, 'restore should not append synthetic acceptance/device annotations');

  console.log('ok restore latest real axis from event log');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
