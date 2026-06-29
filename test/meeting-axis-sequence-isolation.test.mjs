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

function larkStartEvent(id, startSeconds, topic) {
  return {
    schema: '2.0',
    header: {
      event_id: `${id}-start-event`,
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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-sequence-isolation-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'sequence-isolation-app',
    LARK_APP_SECRET: 'sequence-isolation-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://sequence-isolation.example.com/api/lark/events',
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

  const firstStartSeconds = 1_782_720_000;
  await postJson(baseUrl, '/api/lark/events', larkStartEvent(
    'sequence-isolation-first',
    firstStartSeconds,
    'Sequence isolation first meeting',
  ));

  const firstMark = await postJson(baseUrl, '/api/annotations', {
    id: 'first-meeting-only-mark',
    captured_at_ms: (firstStartSeconds * 1000) + 15_000,
    kind: 'handwriting_trigger',
    label: 'only belongs to the first meeting',
    text_candidates: ['only belongs to the first meeting'],
    strokes: [],
  });
  assert.equal(firstMark.state.meeting.meeting_id, 'sequence-isolation-first');
  assert.equal(firstMark.state.sequence.length, 1);

  await postJson(baseUrl, '/api/lark/events', larkStartEvent(
    'sequence-isolation-second',
    firstStartSeconds + 600,
    'Sequence isolation second meeting',
  ));

  const state = await getJson(baseUrl, '/api/state');
  assert.equal(state.meeting.meeting_id, 'sequence-isolation-second');
  assert.equal(state.sequence.length, 0, 'new real meeting must not inherit annotations from the previous real meeting');

  console.log('ok meeting axis sequence isolation');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
