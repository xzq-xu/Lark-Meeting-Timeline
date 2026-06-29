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

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-device-batch-ingest-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'device-batch-contract-app',
    LARK_APP_SECRET: 'device-batch-contract-secret',
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

  const info = await (await fetch(`${baseUrl}/api/annotation-ingest-info`)).json();
  assert.match(info.batch_endpoint, /\/api\/annotations\/batch$/);
  assert.equal(info.response_batch.max_items, 200);
  assert.equal(info.response_batch.object_path, 'acks');
  assert.equal(Array.isArray(info.minimal_batch_payload.annotations), true);

  const capturedAt = Date.parse('2026-06-26T03:00:10.000Z');
  const result = await postJson(baseUrl, '/api/annotations/batch', {
    annotations: [
      {
        id: 'batch-ann-1',
        captured_at_ms: capturedAt,
        kind: 'handwriting_trigger',
        label: 'batch first',
        text_candidates: ['batch first'],
      },
      {
        id: 'batch-ann-2',
        captured_at_ms: capturedAt + 20_000,
        kind: 'attention',
        label: 'batch second',
        text_candidates: ['batch second'],
      },
    ],
  }, {
    'x-hmp-device-id': 'hanwang-batch-001',
  });

  assert.equal(result.accepted, true);
  assert.equal(result.count, 2);
  assert.equal(result.acks.length, 2);
  assert.equal(result.items.length, 2);
  assert.equal(result.state.meeting.source, 'annotation_fallback');
  assert.equal(result.state.meeting.pending_binding, true);
  assert.equal(result.acks[0].created_pending_timeline, true);
  assert.equal(result.acks[1].created_pending_timeline, false);
  assert.equal(result.acks[0].operation, 'created');
  assert.equal(result.acks[1].operation, 'created');
  assert.equal(result.items[0].time_ms, 0);
  assert.equal(result.items[1].time_ms, 20_000);
  assert.equal(result.items[0].source, 'hanwang-batch-001');
  assert.equal(result.items[1].source, 'hanwang-batch-001');
  assert.equal(result.state.sequence.filter((item) => item.id.startsWith('batch-ann-')).length, 2);

  const retry = await postJson(baseUrl, '/api/annotations/batch', [
    {
      id: 'batch-ann-2',
      captured_at_ms: capturedAt + 25_000,
      kind: 'attention',
      label: 'batch second updated',
      text_candidates: ['batch second updated'],
    },
  ], {
    'x-hmp-device-id': 'hanwang-batch-001',
  });
  assert.equal(retry.accepted, true);
  assert.equal(retry.count, 1);
  assert.equal(retry.acks[0].operation, 'updated');
  assert.equal(retry.acks[0].replaced_existing, true);
  assert.equal(retry.state.sequence.filter((item) => item.id === 'batch-ann-2').length, 1);
  assert.equal(retry.state.sequence.find((item) => item.id === 'batch-ann-2')?.label, 'batch second updated');

  const invalid = await fetch(`${baseUrl}/api/annotations/batch`, {
    method: 'POST',
    headers: { 'content-type': 'application/json; charset=utf-8' },
    body: JSON.stringify({ annotations: [] }),
  });
  const invalidJson = await invalid.json();
  assert.equal(invalid.status, 400);
  assert.match(invalidJson.error, /annotations array is required/);

  console.log('ok device batch ingest contract');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
