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

function startServer(port, tempDir) {
  const child = spawn(process.execPath, ['src/server.mjs'], {
    cwd: new URL('..', import.meta.url),
    env: {
      ...process.env,
      PORT: String(port),
      LARK_WS_EVENTS: '0',
      LARK_APP_ID: 'event-log-persistence-app',
      LARK_APP_SECRET: 'event-log-persistence-secret',
      LARK_VERIFICATION_TOKEN: '',
      LARK_EVENT_CALLBACK_URL: 'https://event-log-persistence.example.com/api/lark/events',
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
  child.output = () => output;
  return child;
}

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-timeline-event-log-'));
const baseUrl = `http://127.0.0.1:${port}`;
let child = null;

try {
  child = startServer(port, tempDir);
  await waitForServer(baseUrl);

  const startSeconds = 1_782_442_800;
  await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'persisted-direct-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'persisted-real-meeting',
        topic: 'Persisted real meeting',
        url: 'https://vc.feishu.cn/j/persisted',
        start_time: String(startSeconds),
      },
    },
  });
  let log = await (await fetch(`${baseUrl}/api/lark/events-log`)).json();
  assert.equal(log.count, 1);
  assert.equal(log.items[0].event_type, 'vc.meeting.all_meeting_started_v1');
  assert.equal(log.items[0].timeline_processed, true);
  assert.equal(log.items[0].timeline_started, true);

  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});

  child = startServer(port, tempDir);
  await waitForServer(baseUrl);

  log = await (await fetch(`${baseUrl}/api/lark/events-log`)).json();
  assert.equal(log.count, 1);
  assert.equal(log.items[0].event_type, 'vc.meeting.all_meeting_started_v1');
  assert.equal(log.items[0].timeline_processed, true);
  assert.equal(log.items[0].timeline_started, true);

  const report = await (await fetch(`${baseUrl}/api/lark/acceptance-report`)).json();
  assert.equal(report.current_evidence.real_event_count, 1);
  assert.equal(report.current_evidence.real_meeting_axis_active, true);
  assert.equal(report.current_validation.scoped_by_probe, false);
  const deliveryEvidence = report.open_platform_checklist.find((item) => item.id === 'event_delivery_evidence');
  assert.equal(deliveryEvidence?.status, 'ok');
  assert.match(deliveryEvidence?.evidence ?? '', /recent_public_webhook_event_count=1/);

  await postJson(baseUrl, '/api/lark/real-meeting-probe/start', {
    timeout_ms: 180_000,
    note: 'new validation should ignore historical events',
  });
  const scopedReport = await (await fetch(`${baseUrl}/api/lark/acceptance-report`)).json();
  assert.equal(scopedReport.current_validation.scoped_by_probe, true);
  assert.equal(scopedReport.current_validation.real_event_after_probe, false);
  assert.equal(scopedReport.current_validation.observed_event, null);
  assert.equal(scopedReport.acceptance_steps.find((step) => step.id === 'direct_start_meeting')?.done, false);
  assert.equal(scopedReport.open_platform_checklist.some((item) => item.id === 'event_delivery_mode'), true);
  assert.equal(scopedReport.open_platform_checklist.some((item) => item.id === 'app_availability'), true);

  console.log('ok persisted Lark event log');
} catch (error) {
  if (child?.output) console.error(child.output());
  throw error;
} finally {
  if (child && !child.killed) child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
