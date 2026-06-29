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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-real-demo-auto-open-session-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'real-demo-auto-open-app',
    LARK_APP_SECRET: 'real-demo-auto-open-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://real-demo-auto-open.example.com/api/lark/events',
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
  const capturedAtMs = startSeconds * 1000 + 12_000;
  const firstMark = await postJson(baseUrl, '/api/annotations', {
    id: 'auto-open-first-device-mark',
    source: 'hanwang_epaper',
    device_id: 'hanwang-real-device',
    captured_at_ms: capturedAtMs,
    kind: 'handwriting_trigger',
    label: '手写：为什么?',
    text_candidates: ['为什么?', 'why?'],
    intent: 'question',
    strokes: [
      [
        { x: 0.1, y: 0.2, t: capturedAtMs - 120 },
        { x: 0.2, y: 0.24, t: capturedAtMs },
      ],
    ],
  });

  assert.equal(firstMark.meeting_session_binding.attempted, false);
  assert.equal(firstMark.passive_binding.attempted, false);
  assert.equal(firstMark.auto_open_session_binding.attempted, true);
  assert.equal(firstMark.auto_open_session_binding.status, 'started');
  assert.equal(firstMark.auto_open_session_binding.source, 'open_meeting_session');
  assert.equal(firstMark.ack.on_real_axis, true);
  assert.equal(firstMark.ack.pending_real_meeting, false);
  assert.equal(firstMark.ack.created_pending_timeline, false);
  assert.equal(firstMark.state.meeting.source, 'open_meeting_session');
  assert.equal(firstMark.state.meeting.start_time, new Date(capturedAtMs).toISOString());
  assert.equal(firstMark.state.sequence.find((item) => item.id === 'auto-open-first-device-mark')?.time_ms, 0);

  const openSessionStatus = await getJson(baseUrl, '/api/lark/real-demo/status');
  assert.equal(openSessionStatus.gates.real_meeting_axis_active, true);
  assert.equal(openSessionStatus.gates.real_axis_annotation_count, 1);
  assert.equal(openSessionStatus.evidence.meeting.source, 'open_meeting_session');

  const started = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'auto-open-real-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'auto-open-real-meeting',
        topic: 'Real demo auto open takeover meeting',
        url: 'https://vc.feishu.cn/j/auto-open-real',
        start_time: String(startSeconds),
      },
    },
  });

  assert.equal(started.ok, true);
  assert.equal(started.timeline_started, true);
  assert.equal(started.state.meeting.source, 'lark_http_event');
  assert.equal(started.state.meeting.meeting_id, 'auto-open-real-meeting');
  assert.equal(started.state.sequence.find((item) => item.id === 'auto-open-first-device-mark')?.time_ms, 12_000);

  const finalStatus = await getJson(baseUrl, '/api/lark/real-demo/status');
  assert.equal(finalStatus.status, 'complete');
  assert.equal(finalStatus.gates.real_meeting_axis_active, true);
  assert.equal(finalStatus.gates.real_axis_annotation_count, 1);
  assert.equal(finalStatus.evidence.meeting.source, 'lark_http_event');
  assert.equal(finalStatus.evidence.annotation_binding.items[0].meeting_id, 'auto-open-real-meeting');

  console.log('ok real demo auto-open session annotation takeover');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
