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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-timeline-probe-routing-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'probe-routing-app',
    LARK_APP_SECRET: 'probe-routing-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://probe-routing.example.com/api/lark/events',
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

  const stalePending = await postJson(baseUrl, '/api/annotations', {
    id: 'stale-before-probe',
    source: 'previous_device_debug',
    captured_at_ms: 1_782_442_700_000,
    kind: 'handwriting_trigger',
    label: '上一次调试残留标注',
  });
  assert.equal(stalePending.state.meeting.source, 'annotation_fallback');
  assert.equal(stalePending.state.meeting.pending_binding, true);

  const cleanProbe = await postJson(baseUrl, '/api/lark/real-meeting-probe/start', {
    timeout_ms: 180_000,
    note: 'probe routing cleanup test',
    reset_temporary_axis: true,
  });
  assert.equal(cleanProbe.temporary_axis_reset, true);
  const cleaned = await getJson(baseUrl, '/api/state');
  assert.equal(cleaned.meeting.meeting_id, 'demo-lark-meeting-001');
  assert.equal(cleaned.sequence.some((item) => item.id === 'stale-before-probe'), false);

  await postJson(baseUrl, '/api/live/start-meeting', {
    title: 'Old local simulation',
    meeting_id: 'old-local-simulation',
    meeting_url: 'https://vc.feishu.cn/j/local-simulation',
    minute_token: 'local-simulation-minute-token',
  });
  await postJson(baseUrl, '/api/live/end-meeting', {});

  await postJson(baseUrl, '/api/lark/real-meeting-probe/start', {
    timeout_ms: 180_000,
    note: 'probe routing test',
    reset_temporary_axis: true,
  });

  const startSeconds = 1_782_442_800;
  const pending = await postJson(baseUrl, '/api/annotations', {
    id: 'probe-routing-ann',
    source: 'probe_routing_epaper',
    captured_at_ms: startSeconds * 1000 + 30_000,
    kind: 'handwriting_trigger',
    label: 'probe 期间先到的标注',
    text_candidates: ['probe 期间先到的标注'],
  });
  assert.equal(pending.state.meeting.source, 'annotation_fallback');
  assert.equal(pending.state.meeting.pending_binding, true);
  assert.equal(pending.state.meeting.meeting_url, null);
  assert.equal(pending.state.meeting.minute_token, null);
  assert.equal(pending.state.sequence.find((item) => item.id === 'probe-routing-ann')?.time_ms, 0);
  const pendingBindings = await getJson(baseUrl, '/api/annotation-bindings');
  assert.equal(pendingBindings.binding_state, 'pending_real_meeting');
  assert.equal(pendingBindings.pending_count, 1);
  assert.equal(pendingBindings.real_axis_count, 0);
  const pendingOne = await getJson(baseUrl, '/api/annotation-bindings?id=probe-routing-ann');
  assert.equal(pendingOne.found, true);
  assert.equal(pendingOne.item.binding_state, 'pending_real_meeting');
  assert.equal(pendingOne.item.pending_real_meeting, true);

  const rebound = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'probe-routing-direct-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'probe-routing-real-meeting',
        topic: 'Probe routing real meeting',
        url: 'https://vc.feishu.cn/j/probe-routing',
        start_time: String(startSeconds),
      },
    },
  });
  assert.equal(rebound.state.meeting.meeting_id, 'probe-routing-real-meeting');
  assert.equal(rebound.state.meeting.source, 'lark_http_event');
  assert.equal(rebound.state.meeting.pending_binding, false);
  const item = rebound.state.sequence.find((entry) => entry.id === 'probe-routing-ann');
  assert.equal(item.time_ms, 30_000);
  assert.equal(item.time_source, 'captured_at');
  const reboundBindings = await getJson(baseUrl, '/api/annotation-bindings');
  assert.equal(reboundBindings.binding_state, 'real_meeting_bound');
  assert.equal(reboundBindings.real_axis_count, 1);
  assert.equal(reboundBindings.items.find((entry) => entry.id === 'probe-routing-ann')?.on_real_axis, true);
  const reboundOne = await getJson(baseUrl, '/api/annotation-bindings?id=probe-routing-ann');
  assert.equal(reboundOne.found, true);
  assert.equal(reboundOne.item.binding_state, 'real_meeting_bound');
  assert.equal(reboundOne.item.on_real_axis, true);
  assert.equal(reboundOne.item.meeting_id, 'probe-routing-real-meeting');
  const report = await getJson(baseUrl, '/api/lark/acceptance-report');
  assert.equal(report.current_validation.annotation_binding_state, 'real_meeting_bound');
  assert.equal(report.current_validation.real_axis_annotation_count_after_probe, 1);

  console.log('ok probe annotation routing');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
