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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-real-demo-end-no-bootstrap-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'real-demo-end-no-bootstrap-app',
    LARK_APP_SECRET: 'real-demo-end-no-bootstrap-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://real-demo-end-no-bootstrap.example.com/api/lark/events',
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

  const startSeconds = 1_782_442_800;
  const endSeconds = startSeconds + 60;

  const pending = await postJson(baseUrl, '/api/annotations', {
    id: 'real-demo-pending-before-end',
    source: 'hanwang_epaper',
    captured_at_ms: startSeconds * 1000 + 20_000,
    kind: 'handwriting_trigger',
    label: '会中写下的标注',
    text_candidates: ['会中写下的标注'],
  });
  assert.equal(pending.state.meeting.source, 'annotation_fallback');
  assert.equal(pending.state.meeting.pending_binding, true);

  const endOnly = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'real-demo-end-before-start',
      event_type: 'vc.meeting.all_meeting_ended_v1',
      create_time: String(endSeconds),
    },
    event: {
      meeting: {
        id: 'real-demo-meeting',
        topic: 'Real demo should wait for start',
        start_time: String(startSeconds),
        end_time: String(endSeconds),
      },
    },
  });
  assert.equal(endOnly.ok, false);
  assert.equal(endOnly.ignored_reason, 'real_demo_requires_meeting_start_event');
  assert.equal(endOnly.timeline_started, false);
  assert.equal(endOnly.state.meeting.source, 'annotation_fallback');
  assert.equal(endOnly.state.meeting.pending_binding, true);

  const statusAfterEndOnly = await getJson(baseUrl, '/api/lark/real-demo/status');
  assert.equal(statusAfterEndOnly.complete, false);
  assert.equal(statusAfterEndOnly.gates.real_meeting_axis_active, false);
  assert.equal(statusAfterEndOnly.gates.real_axis_annotation_count, 0);

  const startEvent = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'real-demo-start-after-ignored-end',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'real-demo-meeting',
        topic: 'Real demo should wait for start',
        start_time: String(startSeconds),
      },
    },
  });
  assert.equal(startEvent.ok, true);
  assert.equal(startEvent.timeline_started, true);
  assert.equal(startEvent.state.meeting.source, 'lark_http_event');
  assert.equal(startEvent.state.meeting.pending_binding, false);
  const rebasedMark = startEvent.state.sequence.find((item) => item.id === 'real-demo-pending-before-end');
  assert.equal(rebasedMark.time_ms, 20_000);
  assert.equal(rebasedMark.time_source, 'captured_at');

  const endAfterStart = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'real-demo-end-after-start',
      event_type: 'vc.meeting.all_meeting_ended_v1',
      create_time: String(endSeconds),
    },
    event: {
      meeting: {
        id: 'real-demo-meeting',
        topic: 'Real demo should wait for start',
        start_time: String(startSeconds),
        end_time: String(endSeconds),
      },
    },
  });
  assert.equal(endAfterStart.ok, true);
  assert.equal(endAfterStart.state.meeting.end_time, '2026-06-26T03:01:00.000Z');
  const endEvent = endAfterStart.state.events.find((event) => event.id === 'real-demo-end-after-start');
  const finalMark = endAfterStart.state.sequence.find((item) => item.id === 'real-demo-pending-before-end');
  assert.equal(endEvent.time_ms, 60_000);
  assert.ok(finalMark.time_ms < endEvent.time_ms, 'mark should stay before meeting end');

  const bindings = await getJson(baseUrl, '/api/annotation-bindings');
  assert.equal(bindings.binding_state, 'real_meeting_bound');
  assert.equal(bindings.real_axis_count, 1);

  console.log('ok real demo end event does not bootstrap axis');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
