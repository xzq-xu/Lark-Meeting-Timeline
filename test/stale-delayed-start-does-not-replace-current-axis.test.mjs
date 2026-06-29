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

function larkStartEvent(id, startSeconds, topic, eventId = `${id}-start-event`) {
  return {
    schema: '2.0',
    header: {
      event_id: eventId,
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id,
        topic,
        start_time: String(startSeconds),
      },
    },
  };
}

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-stale-delayed-start-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'stale-delayed-start-app',
    LARK_APP_SECRET: 'stale-delayed-start-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://stale-delayed-start.example.com/api/lark/events',
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

  const currentStartSeconds = 1_782_720_600;
  await postJson(baseUrl, '/api/lark/events', larkStartEvent(
    'current-real-meeting',
    currentStartSeconds,
    'Current real meeting',
  ));
  let state = await getJson(baseUrl, '/api/state');
  assert.equal(state.meeting.meeting_id, 'current-real-meeting');

  const stale = await postJson(baseUrl, '/api/lark/events', larkStartEvent(
    'stale-old-meeting',
    currentStartSeconds - 3600,
    'Stale old meeting',
    'stale-old-start-event',
  ));
  assert.equal(stale.ok, false);
  assert.equal(stale.ignored_reason, 'stale_meeting_start_before_current_axis');

  state = await getJson(baseUrl, '/api/state');
  assert.equal(state.meeting.meeting_id, 'current-real-meeting');
  assert.equal(state.meeting.title, 'Current real meeting');

  const log = await getJson(baseUrl, '/api/lark/events-log');
  assert.equal(log.items[0].event_type, 'vc.meeting.all_meeting_started_v1');
  assert.equal(log.items[0].ignored_reason, 'stale_meeting_start_before_current_axis');
  assert.equal(log.items[0].timeline_started, false);

  console.log('ok stale delayed start does not replace current axis');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
