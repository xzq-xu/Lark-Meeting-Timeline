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

async function postJson(baseUrl, path, body, headers = {}) {
  const response = await fetch(`${baseUrl}${path}`, {
    method: 'POST',
    headers: { 'content-type': 'application/json; charset=utf-8', ...headers },
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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-annotation-status-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'annotation-status-contract-app',
    LARK_APP_SECRET: 'annotation-status-contract-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://annotation-status-contract.example.com/api/lark/events',
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

  const info = await getJson(baseUrl, '/api/annotation-ingest-info');
  assert.match(info.annotation_status_url, /\/api\/annotation-status\?id=\{annotation_id\}$/);
  assert.match(info.device_client_contract.status_polling_endpoint, /\/api\/annotation-status\?id=\{annotation_id\}$/);

  const missingId = await fetch(`${baseUrl}/api/annotation-status`);
  const missingIdJson = await missingId.json();
  assert.equal(missingId.status, 400);
  assert.match(missingIdJson.error, /id or annotation_id is required/);

  const notFound = await getJson(baseUrl, '/api/annotation-status?id=missing-status-ann');
  assert.equal(notFound.found, false);
  assert.equal(notFound.status, 'not_found');
  assert.equal(notFound.accepted, false);

  const startSeconds = 1_782_442_800;
  const capturedAt = startSeconds * 1000 + 12_000;
  const pending = await postJson(baseUrl, '/api/annotations', {
    id: 'status-pending-ann',
    captured_at_ms: capturedAt,
    kind: 'handwriting_trigger',
    label: 'why?',
    text_candidates: ['why?', 'why'],
  }, {
    'x-hmp-device-id': 'hanwang-status-001',
  });
  assert.equal(pending.ack.binding_state, 'pending_real_meeting');

  const pendingStatus = await getJson(baseUrl, '/api/annotation-status?id=status-pending-ann');
  assert.equal(pendingStatus.found, true);
  assert.equal(pendingStatus.status, 'pending_real_meeting');
  assert.equal(pendingStatus.pending_real_meeting, true);
  assert.equal(pendingStatus.on_real_axis, false);
  assert.equal(pendingStatus.poll_after_ms, 1000);

  await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'status-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'status-real-meeting',
        topic: 'Annotation status contract',
        start_time: String(startSeconds),
      },
    },
  });

  const boundStatus = await getJson(baseUrl, '/api/annotation-status?id=status-pending-ann');
  assert.equal(boundStatus.status, 'real_axis_bound');
  assert.equal(boundStatus.binding_state, 'real_meeting_bound');
  assert.equal(boundStatus.on_real_axis, true);
  assert.equal(boundStatus.normalized_time_ms, 12_000);
  assert.equal(boundStatus.next_action, 'done');

  await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'status-end',
      event_type: 'vc.meeting.all_meeting_ended_v1',
      create_time: String(startSeconds + 60),
    },
    event: {
      meeting: {
        id: 'status-real-meeting',
        topic: 'Annotation status contract',
        end_time: String(startSeconds + 60),
      },
    },
  });

  await postJson(baseUrl, '/api/annotations', {
    id: 'status-after-end-ann',
    captured_at_ms: startSeconds * 1000 + 90_000,
    kind: 'handwriting_trigger',
    label: 'too late',
  }, {
    'x-hmp-device-id': 'hanwang-status-001',
  });
  const afterEndStatus = await getJson(baseUrl, '/api/annotation-status?id=status-after-end-ann');
  assert.equal(afterEndStatus.status, 'after_meeting_end');
  assert.equal(afterEndStatus.on_real_axis, false);
  assert.equal(afterEndStatus.after_meeting_end_ms, 30_000);
  assert.ok(afterEndStatus.warnings.includes('normalized_after_meeting_end'));

  const noCaptured = await postJson(baseUrl, '/api/annotations', {
    id: 'status-no-captured-ann',
    kind: 'handwriting_trigger',
    label: 'missing captured time',
  }, {
    'x-hmp-device-id': 'hanwang-status-001',
  });
  assert.equal(noCaptured.ack.requires_device_captured_at, true);
  const noCapturedStatus = await getJson(baseUrl, '/api/annotations/status?id=status-no-captured-ann');
  assert.equal(noCapturedStatus.status, 'needs_device_captured_at');
  assert.equal(noCapturedStatus.requires_device_captured_at, true);
  assert.ok(noCapturedStatus.warnings.includes('missing_captured_at_ms'));

  console.log('ok annotation status contract');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
