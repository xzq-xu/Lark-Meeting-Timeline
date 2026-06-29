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

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-ws-parser-self-test-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'ws-parser-self-test-app',
    LARK_APP_SECRET: 'ws-parser-self-test-secret',
    LARK_VERIFICATION_TOKEN: '',
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

  const selfTest = await (await fetch(`${baseUrl}/api/lark/ws-parser-self-test`)).json();
  assert.equal(selfTest.passed, true);
  assert.equal(selfTest.mutates_state, false);
  assert.equal(selfTest.cases.direct_meeting_start.parsed_event_type, 'vc.meeting.all_meeting_started_v1');
  assert.equal(selfTest.cases.direct_meeting_start.timeline_candidate, true);
  assert.equal(selfTest.cases.direct_meeting_start.normalized_event.type, 'meeting_start');
  assert.equal(selfTest.cases.direct_meeting_start.would_start_timeline, true);
  assert.equal(selfTest.cases.meeting_context_join.parsed_event_type, 'vc.meeting.join_meeting_v1');
  assert.equal(selfTest.cases.meeting_context_join.timeline_candidate, true);
  assert.equal(selfTest.cases.meeting_context_join.normalized_event.type, 'participant_join');
  assert.equal(selfTest.cases.meeting_context_join.would_start_timeline, true);

  const diagnostics = await (await fetch(`${baseUrl}/api/lark/delivery-diagnostics`)).json();
  assert.equal(diagnostics.parser_self_test.passed, true);

  const log = await (await fetch(`${baseUrl}/api/lark/events-log`)).json();
  assert.equal(log.count, 0, 'parser self-test must not mutate event log');

  console.log('ok WS parser self test');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
