import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
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

async function runDevicePreflight(baseUrl, args = []) {
  const child = spawn(process.execPath, [
    'scripts/device-preflight.mjs',
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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-device-preflight-'));
const baseUrl = `http://127.0.0.1:${port}`;
const server = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'device-preflight-app',
    LARK_APP_SECRET: 'device-preflight-secret',
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

  const reportFile = join(tempDir, 'device-preflight-report.json');
  const result = await runDevicePreflight(baseUrl, [`--report-file=${reportFile}`]);
  assert.equal(result.code, 0, result.stderr);
  assert.equal(result.json.type, 'device_preflight');
  assert.equal(result.json.ok, true);
  assert.match(result.json.endpoint, /\/api\/annotations$/);
  assert.match(result.json.real_demo_acceptance_url, /\/api\/lark\/real-demo\/acceptance$/);
  assert.equal(result.json.route.mode, 'create_pending_on_first_annotation');
  assert.equal(result.json.device.preferred_time_field, 'captured_at_ms');
  assert.equal(result.json.device.clock_sync_required, true);
  assert.match(result.json.device.product_acceptance_condition, /product_acceptance_complete/);
  assert.equal(result.json.checks.cors.ok, true);
  assert.equal(result.json.checks.clock.ok, true);
  assert.equal(typeof result.json.checks.clock.offset_ms, 'number');
  assert.equal(typeof result.json.checks.clock.rtt_ms, 'number');
  assert.equal(result.json.checks.stream.ok, true);
  assert.equal(result.json.checks.acceptance.ok, true);
  assert.equal(result.json.checks.acceptance.product_acceptance_complete, false);
  assert.match(result.json.checks.acceptance.recommended_command, /accept:real-meeting/);
  assert.ok(result.json.checks.acceptance.missing.includes('real_meeting_axis_active'));

  const report = JSON.parse(await readFile(reportFile, 'utf8'));
  assert.equal(report.type, 'device_preflight');
  assert.equal(report.ok, true);
  assert.equal(report.endpoint, result.json.endpoint);
  assert.equal(report.checks.cors.ok, true);
  assert.equal(report.checks.acceptance.verdict, result.json.checks.acceptance.verdict);
  assert.equal(report.report.file, reportFile);
  assert.equal(report.report.command, 'scripts/device-preflight.mjs');
  assert.equal(report.report.base_url, baseUrl);
  assert.equal(report.report.exit_code, 0);
  assert.equal(report.report.error, null);

  console.log('ok device preflight');
} catch (error) {
  console.error(serverOutput);
  throw error;
} finally {
  server.kill('SIGTERM');
  await once(server, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
