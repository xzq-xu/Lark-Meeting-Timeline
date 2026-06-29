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

async function stop(child) {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
}

function startServer(port, tempDir) {
  return spawn(process.execPath, ['src/server.mjs'], {
    cwd: new URL('..', import.meta.url),
    env: {
      ...process.env,
      PORT: String(port),
      LARK_WS_EVENTS: '0',
      LARK_APP_ID: 'probe-test-app',
      LARK_APP_SECRET: 'probe-test-secret',
      TIMELINE_DATA_DIR: tempDir,
    },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
}

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-timeline-probe-'));
const baseUrl = `http://127.0.0.1:${port}`;
let child = null;

try {
  child = startServer(port, tempDir);
  await waitForServer(baseUrl);
  const started = await (await fetch(`${baseUrl}/api/lark/real-meeting-probe/start`, {
    method: 'POST',
    headers: { 'content-type': 'application/json; charset=utf-8' },
    body: JSON.stringify({ timeout_ms: 180_000, note: 'probe persistence test' }),
  })).json();
  assert.equal(started.status, 'waiting');
  assert.equal(started.active, true);
  assert.equal(started.auto_search.enabled, true);
  assert.equal(started.auto_search.server_loop.scheduled, true);
  await stop(child);

  child = startServer(port, tempDir);
  await waitForServer(baseUrl);
  const restored = await (await fetch(`${baseUrl}/api/lark/real-meeting-probe`)).json();
  assert.equal(restored.status, 'waiting');
  assert.equal(restored.active, true);
  assert.equal(restored.started_at, started.started_at);
  assert.equal(restored.note, 'probe persistence test');
  assert.equal(restored.auto_search.server_loop.scheduled, true);

  const reset = await (await fetch(`${baseUrl}/api/lark/real-meeting-probe/reset`, {
    method: 'POST',
    headers: { 'content-type': 'application/json; charset=utf-8' },
    body: '{}',
  })).json();
  assert.equal(reset.status, 'idle');
  assert.equal(reset.active, false);
  assert.equal(reset.started_at, null);
  assert.equal(reset.auto_search.server_loop.scheduled, false);
  await stop(child);

  child = startServer(port, tempDir);
  await waitForServer(baseUrl);
  const resetRestored = await (await fetch(`${baseUrl}/api/lark/real-meeting-probe`)).json();
  assert.equal(resetRestored.status, 'idle');
  assert.equal(resetRestored.active, false);
  assert.equal(resetRestored.started_at, null);

  console.log('ok probe persistence');
} finally {
  if (child) await stop(child).catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
