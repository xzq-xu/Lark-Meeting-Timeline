import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
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

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-raw-ws-diagnostics-'));
const baseUrl = `http://127.0.0.1:${port}`;
const now = new Date('2026-06-26T03:00:00.000Z').toISOString();
await writeFile(join(tempDir, 'lark-events-log.json'), JSON.stringify([
  {
    id: 'raw-candidate',
    at: now,
    transport: 'ws_long_connection',
    event_type: 'vc.meeting.unknown_payload_v1',
    timeline_candidate: true,
    timeline_processed: false,
    timeline_started: false,
    ignored_reason: 'unknown_shape',
    parsed_keys: ['header', 'event'],
    preview: { header: { event_type: 'vc.meeting.unknown_payload_v1' } },
  },
  {
    id: 'raw-ignored',
    at: now,
    transport: 'ws_long_connection',
    event_type: 'im.message.receive_v1',
    timeline_candidate: false,
    timeline_processed: false,
    timeline_started: false,
    ignored_reason: 'not_meeting_related',
    parsed_keys: ['header', 'event'],
    preview: { header: { event_type: 'im.message.receive_v1' } },
  },
], null, 2));

const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'raw-ws-diagnostics-app',
    LARK_APP_SECRET: 'raw-ws-diagnostics-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://raw-ws-diagnostics.example.com/api/lark/events',
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

  const config = await (await fetch(`${baseUrl}/api/lark/config`)).json();
  assert.equal(config.event_log_summary.ws_event_count, 2);
  assert.equal(config.event_log_summary.ws_timeline_candidate_count, 1);
  assert.equal(config.event_log_summary.ws_timeline_started_count, 0);
  assert.equal(config.event_log_summary.ws_ignored_count, 1);
  assert.equal(config.event_log_summary.last_ws_event_status, 'timeline_candidate');

  const diagnostics = await (await fetch(`${baseUrl}/api/lark/delivery-diagnostics`)).json();
  assert.equal(diagnostics.status, 'ws_events_seen_but_no_meeting_start');
  assert.equal(diagnostics.evidence.ws_event_count, 2);
  assert.equal(diagnostics.evidence.ws_timeline_candidate_count, 1);
  assert.equal(diagnostics.evidence.ws_timeline_started_count, 0);
  assert.equal(diagnostics.evidence.ws_ignored_count, 1);
  assert.match(diagnostics.summary, /像会议\/妙记候选事件/);
  assert.match(diagnostics.next_actions[0], /payload 字段/);

  console.log('ok raw WS event diagnostics');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
