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

async function runMonitor(baseUrl, extraArgs = []) {
  const child = spawn(process.execPath, ['scripts/monitor-real-demo.mjs', `--url=${baseUrl}`, '--once', ...extraArgs], {
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
  return { code, stdout, stderr };
}

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-real-demo-monitor-cli-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'real-demo-monitor-cli-app',
    LARK_APP_SECRET: 'real-demo-monitor-cli-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://real-demo-monitor-cli.example.com/api/lark/events',
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

  const result = await runMonitor(baseUrl);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /status=ready_to_open_meeting/);
  assert.match(result.stdout, /real_axis=no/);
  assert.match(result.stdout, /event_axis=no/);
  assert.match(result.stdout, /axis_mode=not_built/);
  assert.match(result.stdout, /axis_source=none/);
  assert.match(result.stdout, /annotations=0/);
  assert.match(result.stdout, /audit=no_event_delivery_observed/);
  assert.match(result.stdout, /requirements=todo:real_meeting_axis_active,event_meeting_axis_built,open_annotation_on_real_axis,realtime_state_broadcast/);
  assert.match(result.stdout, /auth=/);
  assert.match(result.stdout, /scan=/);

  const preparedResult = await runMonitor(baseUrl, ['--prepare']);
  assert.equal(preparedResult.code, 0, preparedResult.stderr);
  assert.match(preparedResult.stdout, /prepared=true/);
  assert.match(preparedResult.stdout, /device_stream=/);
  assert.match(preparedResult.stdout, /status=ready_to_open_meeting/);

  const openAuthResult = await runMonitor(baseUrl, ['--open-auth']);
  assert.equal(openAuthResult.code, 0, openAuthResult.stderr);
  assert.match(openAuthResult.stdout, /auth_opened=true|auth_opened=false/);
  assert.match(openAuthResult.stdout, /status=ready_to_open_meeting/);

  const waitAuthTimeout = await runMonitor(baseUrl, ['--wait-auth', '--auth-timeout-ms=20', '--interval-ms=500']);
  assert.equal(waitAuthTimeout.code, 2);
  assert.match(waitAuthTimeout.stdout, /auth_wait/);
  assert.match(waitAuthTimeout.stderr, /auth wait timed out/);

  await postJson(baseUrl, '/api/acceptance/auto-annotation', { enabled: false });
  await postJson(baseUrl, '/api/device-simulator', { enabled: false });
  await postJson(baseUrl, '/api/device-simulator/stream', { enabled: false });
  const startSeconds = Math.floor(Date.now() / 1000) - 30;
  await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'monitor-cli-auto-mark-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'monitor-cli-auto-mark-meeting',
        topic: 'Monitor CLI auto mark meeting',
        start_time: String(startSeconds),
      },
    },
  });

  const reportFile = join(tempDir, 'monitor-report.json');
  const autoMarkResult = await runMonitor(baseUrl, [
    '--auto-mark',
    '--require-event-axis',
    '--mark-id=monitor-cli-auto-mark',
    '--mark-label=monitor CLI 自动标注',
    `--report-file=${reportFile}`,
  ]);
  assert.equal(autoMarkResult.code, 0, autoMarkResult.stderr);
  assert.match(autoMarkResult.stdout, /auto_mark_written=true/);
  assert.match(autoMarkResult.stdout, /id=monitor-cli-auto-mark/);
  assert.match(autoMarkResult.stdout, /on_real_axis=yes/);
  assert.match(autoMarkResult.stdout, /complete=yes/);
  assert.match(autoMarkResult.stdout, /event_axis=yes/);
  assert.match(autoMarkResult.stdout, /axis_mode=meeting_start_event/);
  assert.match(autoMarkResult.stdout, /axis_source=lark_http_event/);
  assert.match(autoMarkResult.stdout, /requirements=all_ok/);
  assert.match(autoMarkResult.stdout, /strict_event_axis=yes/);
  const bindings = await getJson(baseUrl, '/api/annotation-bindings?id=monitor-cli-auto-mark');
  assert.equal(bindings.found, true);
  assert.equal(bindings.item.on_real_axis, true);
  assert.equal(bindings.item.binding_state, 'real_meeting_bound');
  const report = JSON.parse(await readFile(reportFile, 'utf8'));
  assert.equal(report.type, 'real_demo_acceptance_report');
  assert.equal(report.outcome.exit_code, 0);
  assert.equal(report.outcome.completed, true);
  assert.equal(report.outcome.strict_event_axis_complete, true);
  assert.equal(report.outcome.axis_creation_mode, 'meeting_start_event');
  assert.equal(report.outcome.meeting_source, 'lark_http_event');
  assert.equal(report.outcome.real_axis_annotation_count, 1);
  assert.equal(report.outcome.requirements.every((item) => item.ok), true);

  console.log('ok real demo monitor CLI');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
