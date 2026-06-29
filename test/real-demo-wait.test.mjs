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

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-real-demo-wait-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'real-demo-wait-app',
    LARK_APP_SECRET: 'real-demo-wait-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://real-demo-wait.example.com/api/lark/events',
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
    device_simulator: true,
    passive_scan: false,
  });

  const timedOut = await postJson(baseUrl, '/api/lark/real-demo/wait', {
    timeout_ms: 1000,
    interval_ms: 100,
  });
  assert.equal(timedOut.observed, false);
  assert.equal(timedOut.timed_out, true);
  assert.notEqual(timedOut.status, 'complete');

  const waitPromise = postJson(baseUrl, '/api/lark/real-demo/wait', {
    timeout_ms: 5000,
    interval_ms: 100,
  });

  await new Promise((resolve) => setTimeout(resolve, 250));
  const startSeconds = Math.floor(Date.now() / 1000);
  await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'real-demo-wait-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'real-demo-wait-meeting',
        topic: 'Real demo wait meeting',
        url: 'https://vc.feishu.cn/j/real-demo-wait',
        start_time: String(startSeconds),
      },
    },
  });

  const observed = await waitPromise;
  assert.equal(observed.observed, true);
  assert.equal(observed.timed_out, false);
  assert.equal(observed.status, 'complete');
  assert.equal(observed.result.complete, true);
  assert.equal(observed.result.gates.real_meeting_axis_active, true);
  assert.equal(observed.result.gates.real_axis_annotation_count, 1);
  assert.equal(observed.result.evidence.annotation_binding.items[0].source, 'hanwang_epaper_simulator');

  console.log('ok real demo wait endpoint');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
