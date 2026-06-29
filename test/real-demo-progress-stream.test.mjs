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

async function startProgressStream(baseUrl) {
  const controller = new AbortController();
  const response = await fetch(`${baseUrl}/api/lark/real-demo/progress-stream?interval_ms=500`, {
    signal: controller.signal,
  });
  assert.equal(response.ok, true);
  assert.match(response.headers.get('content-type') ?? '', /text\/event-stream/);

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  const queue = [];
  const waiters = [];
  let buffer = '';

  const publish = (payload) => {
    queue.push(payload);
    for (let index = waiters.length - 1; index >= 0; index -= 1) {
      const waiter = waiters[index];
      const matched = queue.find((item) => waiter.predicate(item));
      if (!matched) continue;
      waiters.splice(index, 1);
      clearTimeout(waiter.timer);
      waiter.resolve(matched);
    }
  };

  const parseBlock = (block) => {
    const event = block.split('\n').find((line) => line.startsWith('event: '))?.slice(7);
    const data = block
      .split('\n')
      .filter((line) => line.startsWith('data: '))
      .map((line) => line.slice(6))
      .join('\n');
    if (event === 'progress' && data) publish(JSON.parse(data));
  };

  const pump = (async () => {
    while (true) {
      const { done, value } = await reader.read();
      if (done) return;
      buffer += decoder.decode(value, { stream: true });
      let splitAt = buffer.indexOf('\n\n');
      while (splitAt >= 0) {
        parseBlock(buffer.slice(0, splitAt));
        buffer = buffer.slice(splitAt + 2);
        splitAt = buffer.indexOf('\n\n');
      }
    }
  })().catch((error) => {
    if (error.name !== 'AbortError') throw error;
  });

  return {
    nextProgress(predicate, timeoutMs = 6000) {
      const existing = queue.find((item) => predicate(item));
      if (existing) return Promise.resolve(existing);
      return new Promise((resolve, reject) => {
        const timer = setTimeout(() => {
          const index = waiters.findIndex((item) => item.timer === timer);
          if (index >= 0) waiters.splice(index, 1);
          reject(new Error('timed out waiting for progress event'));
        }, timeoutMs);
        waiters.push({ predicate, resolve, timer });
      });
    },
    async close() {
      controller.abort();
      await reader.cancel().catch(() => {});
      await pump.catch(() => {});
    },
  };
}

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-real-demo-progress-stream-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'real-demo-progress-stream-app',
    LARK_APP_SECRET: 'real-demo-progress-stream-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://real-demo-progress-stream.example.com/api/lark/events',
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

let stream = null;
try {
  await waitForServer(baseUrl);

  await postJson(baseUrl, '/api/lark/real-demo/prepare', {
    auto_annotation: false,
    device_simulator: false,
    device_stream: false,
    passive_scan: false,
  });

  stream = await startProgressStream(baseUrl);
  const waiting = await stream.nextProgress((event) => event.status === 'ready_to_open_meeting');
  assert.equal(waiting.completion_evidence.real_meeting_axis_active, false);
  assert.equal(waiting.completion_evidence.event_audit_status, 'no_event_delivery_observed');

  const startSeconds = 1_782_442_800;
  await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'progress-stream-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'progress-stream-real-meeting',
        topic: 'Real demo progress stream meeting',
        start_time: String(startSeconds),
      },
    },
  });

  const withAxis = await stream.nextProgress((event) => event.completion_evidence.real_meeting_axis_active);
  assert.equal(withAxis.status, 'real_axis_waiting_annotation');
  assert.equal(withAxis.completion_evidence.real_axis_annotation_count, 0);

  await postJson(baseUrl, '/api/annotations', {
    id: 'progress-stream-open-ann',
    source: 'future_epaper_device',
    captured_at_ms: startSeconds * 1000 + 18_000,
    kind: 'handwriting_trigger',
    label: 'why?',
    text_candidates: ['why?', 'why'],
  });

  const complete = await stream.nextProgress((event) => event.completion_evidence.real_demo_complete);
  assert.equal(complete.status, 'complete');
  assert.equal(complete.completion_evidence.real_axis_annotation_count, 1);
  assert.equal(complete.completion_evidence.event_audit_status, 'event_delivery_ok');

  console.log('ok real demo progress stream');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  if (stream) await stream.close();
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
