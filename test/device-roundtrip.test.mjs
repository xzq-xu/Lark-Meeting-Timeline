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

async function runDeviceRoundtrip(baseUrl, args = []) {
  const child = spawn(process.execPath, [
    'scripts/device-roundtrip.mjs',
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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-device-roundtrip-'));
const baseUrl = `http://127.0.0.1:${port}`;
const server = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'device-roundtrip-app',
    LARK_APP_SECRET: 'device-roundtrip-secret',
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

  const result = await runDeviceRoundtrip(baseUrl, [
    '--id=device-roundtrip-001',
    '--label=roundtrip mark',
    '--timeout-ms=5000',
  ]);
  assert.equal(result.code, 0, result.stderr);
  assert.equal(result.json.type, 'device_roundtrip_result');
  assert.equal(result.json.ok, true);
  assert.equal(result.json.annotation_id, 'device-roundtrip-001');
  assert.equal(result.json.ack.accepted, true);
  assert.equal(result.json.ack.pending_real_meeting, true);
  assert.equal(result.json.status.status, 'pending_real_meeting');
  assert.equal(result.json.sse.observed, true);
  assert.equal(result.json.sse.item_id, 'device-roundtrip-001');
  assert.equal(result.json.sse.item_label, 'roundtrip mark');
  assert.equal(typeof result.json.clock_sync.offset_ms, 'number');

  console.log('ok device roundtrip');
} catch (error) {
  console.error(serverOutput);
  throw error;
} finally {
  server.kill('SIGTERM');
  await once(server, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
