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

async function waitForSseState(baseUrl, predicate, timeoutMs = 5_000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(`${baseUrl}/api/stream`, { signal: controller.signal });
    assert.equal(response.ok, true, 'SSE stream should open');
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const chunks = buffer.split('\n\n');
      buffer = chunks.pop() ?? '';
      for (const chunk of chunks) {
        const event = chunk.split('\n').find((line) => line.startsWith('event: '))?.slice(7);
        const data = chunk.split('\n').find((line) => line.startsWith('data: '))?.slice(6);
        if (event === 'state' && data) {
          const state = JSON.parse(data);
          if (predicate(state)) {
            await reader.cancel();
            return state;
          }
        }
      }
    }
  } finally {
    clearTimeout(timer);
  }
  throw new Error('SSE state predicate was not observed');
}

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-auto-acceptance-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'auto-acceptance-app',
    LARK_APP_SECRET: 'auto-acceptance-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://auto-acceptance.example.com/api/lark/events',
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

  const initial = await (await fetch(`${baseUrl}/api/acceptance/auto-annotation`)).json();
  assert.equal(initial.enabled, false);

  const enabled = await postJson(baseUrl, '/api/acceptance/auto-annotation', { enabled: true });
  assert.equal(enabled.auto_acceptance.enabled, true);

  const startSeconds = 1_782_442_800;
  const expectedId = 'auto-acceptance-auto-real-meeting';
  const ssePromise = waitForSseState(baseUrl, (state) => (
    state.meeting?.meeting_id === 'auto-real-meeting'
      && state.sequence?.some((item) => item.id === expectedId && item.source === 'demo_auto_acceptance')
  ));

  const started = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'auto-acceptance-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'auto-real-meeting',
        topic: 'Auto acceptance real meeting',
        url: 'https://vc.feishu.cn/j/auto-acceptance',
        start_time: String(startSeconds),
      },
    },
  });

  assert.equal(started.auto_acceptance_annotation.id, expectedId);
  const sseState = await ssePromise;
  const autoItem = sseState.sequence.find((item) => item.id === expectedId);
  assert.equal(autoItem.kind, 'acceptance_mark');
  assert.equal(autoItem.payload.origin ?? autoItem.payload.raw_payload?.origin, 'auto_acceptance');

  const bindings = await (await fetch(`${baseUrl}/api/annotation-bindings?id=${expectedId}`)).json();
  assert.equal(bindings.found, true);
  assert.equal(bindings.item.binding_state, 'real_meeting_bound');
  assert.equal(bindings.item.on_real_axis, true);

  const readiness = await (await fetch(`${baseUrl}/api/readiness`)).json();
  assert.equal(readiness.checks.find((check) => check.id === 'real_annotation_seen')?.ok, true);
  assert.equal(readiness.auto_acceptance.enabled, true);
  assert.equal(readiness.auto_acceptance.last_annotation_id, expectedId);

  const duplicate = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'auto-acceptance-start-duplicate',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'auto-real-meeting',
        topic: 'Auto acceptance real meeting',
        url: 'https://vc.feishu.cn/j/auto-acceptance',
        start_time: String(startSeconds),
      },
    },
  });
  assert.equal(duplicate.auto_acceptance_annotation, null);
  assert.equal(duplicate.state.sequence.filter((item) => item.id === expectedId).length, 1);

  console.log('ok auto acceptance annotation');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
