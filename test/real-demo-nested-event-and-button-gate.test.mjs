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

async function postJson(baseUrl, path, body, expectedOk = true) {
  const response = await fetch(`${baseUrl}${path}`, {
    method: 'POST',
    headers: { 'content-type': 'application/json; charset=utf-8' },
    body: JSON.stringify(body),
  });
  const json = await response.json();
  assert.equal(response.ok, expectedOk, `${path} expected ok=${expectedOk}: ${JSON.stringify(json)}`);
  return json;
}

async function getJson(baseUrl, path) {
  const response = await fetch(`${baseUrl}${path}`);
  const json = await response.json();
  assert.equal(response.ok, true, `${path} failed: ${JSON.stringify(json)}`);
  return json;
}

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-real-demo-nested-event-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'real-demo-nested-event-app',
    LARK_APP_SECRET: 'real-demo-nested-event-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://real-demo-nested-event.example.com/api/lark/events',
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
    auto_open_session_on_annotation: false,
  });

  const blockedStart = await postJson(baseUrl, '/api/live/start-meeting', {
    title: 'should not create axis',
  }, false);
  assert.match(blockedStart.error, /真实演示模式/);

  const startSeconds = 1_782_442_800;
  const pending = await postJson(baseUrl, '/api/annotations', {
    id: 'nested-shape-pending-mark',
    source: 'hanwang_epaper',
    captured_at_ms: startSeconds * 1000 + 25_000,
    kind: 'handwriting_trigger',
    label: '真实会议前后的手写标注',
    text_candidates: ['真实会议前后的手写标注'],
  });
  assert.equal(pending.state.meeting.pending_binding, true);

  const started = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    data: {
      header: {
        event_id: 'nested-data-start',
        event_type: 'vc.meeting.all_meeting_started_v1',
        create_time: String(startSeconds),
      },
      event: {
        meeting: {
          id: 'nested-real-meeting',
          topic: 'Nested real event shape',
          start_time: String(startSeconds),
          url: 'https://vc.feishu.cn/j/nested-real-event',
        },
      },
    },
  });

  assert.equal(started.ok, true);
  assert.equal(started.timeline_started, true);
  assert.equal(started.state.meeting.source, 'lark_http_event');
  assert.equal(started.state.meeting.meeting_id, 'nested-real-meeting');
  assert.equal(started.state.meeting.start_time, '2026-06-26T03:00:00.000Z');
  assert.equal(started.state.events.find((event) => event.id === 'nested-data-start')?.time_ms, 0);
  assert.equal(started.state.sequence.find((item) => item.id === 'nested-shape-pending-mark')?.time_ms, 25_000);

  const log = await getJson(baseUrl, '/api/lark/events-log');
  assert.equal(log.items[0].event_type, 'vc.meeting.all_meeting_started_v1');
  assert.equal(log.items[0].timeline_processed, true);
  assert.equal(log.items[0].timeline_started, true);

  const status = await getJson(baseUrl, '/api/lark/real-demo/status');
  assert.equal(status.gates.real_meeting_axis_active, true);

  console.log('ok real demo nested event shape and button gate');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
