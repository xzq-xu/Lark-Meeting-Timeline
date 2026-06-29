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

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-annotation-end-recompute-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'annotation-end-recompute-app',
    LARK_APP_SECRET: 'annotation-end-recompute-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://annotation-end-recompute.example.com/api/lark/events',
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
  const startSeconds = 1_782_442_800;

  await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'recompute-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'recompute-real-meeting',
        topic: 'Annotation end recompute',
        start_time: String(startSeconds),
      },
    },
  });

  const beforeEnd = await postJson(baseUrl, '/api/annotations', {
    id: 'mark-before-end',
    source: 'hanwang_epaper',
    captured_at_ms: startSeconds * 1000 + 30_000,
    kind: 'handwriting_trigger',
    label: '结束前标注',
  });
  assert.equal(beforeEnd.ack.on_real_axis, true);

  const laterThanFutureEnd = await postJson(baseUrl, '/api/annotations', {
    id: 'mark-after-future-end',
    source: 'hanwang_epaper',
    captured_at_ms: startSeconds * 1000 + 90_000,
    kind: 'handwriting_trigger',
    label: '后续被结束事件判定为结束后',
  });
  assert.equal(laterThanFutureEnd.ack.on_real_axis, true);

  await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'recompute-end',
      event_type: 'vc.meeting.all_meeting_ended_v1',
      create_time: String(startSeconds + 60),
    },
    event: {
      meeting: {
        id: 'recompute-real-meeting',
        topic: 'Annotation end recompute',
        end_time: String(startSeconds + 60),
      },
    },
  });

  const bindings = await getJson(baseUrl, '/api/annotation-bindings');
  const before = bindings.items.find((item) => item.id === 'mark-before-end');
  const after = bindings.items.find((item) => item.id === 'mark-after-future-end');
  assert.equal(bindings.binding_state, 'real_meeting_bound');
  assert.equal(before.on_real_axis, true);
  assert.equal(before.after_meeting_end_ms, 0);
  assert.equal(after.on_real_axis, false);
  assert.equal(after.after_meeting_end_ms, 30_000);
  assert.equal(bindings.real_axis_count, 1);

  const lateUploadBeforeEnd = await postJson(baseUrl, '/api/annotations', {
    id: 'mark-uploaded-after-end-but-captured-before',
    source: 'hanwang_epaper',
    captured_at_ms: startSeconds * 1000 + 45_000,
    kind: 'handwriting_trigger',
    label: '结束后上传但结束前采集',
  });
  assert.equal(lateUploadBeforeEnd.item.time_ms, 45_000);
  assert.equal(lateUploadBeforeEnd.ack.on_real_axis, true);
  assert.equal(lateUploadBeforeEnd.ack.after_meeting_end_ms, 0);

  await postJson(baseUrl, '/api/import/lark-transcript', {
    meeting: { end_time: null },
    transcript: [],
  });
  const bindingsWithEventOnlyEnd = await getJson(baseUrl, '/api/annotation-bindings');
  const afterWithEventOnlyEnd = bindingsWithEventOnlyEnd.items.find((item) => item.id === 'mark-after-future-end');
  const lateUploadWithEventOnlyEnd = bindingsWithEventOnlyEnd.items.find((item) => item.id === 'mark-uploaded-after-end-but-captured-before');
  assert.equal(afterWithEventOnlyEnd.on_real_axis, false);
  assert.equal(afterWithEventOnlyEnd.after_meeting_end_ms, 30_000);
  assert.equal(lateUploadWithEventOnlyEnd.on_real_axis, true);
  assert.equal(lateUploadWithEventOnlyEnd.after_meeting_end_ms, 0);
  assert.equal(bindingsWithEventOnlyEnd.real_axis_count, 2);

  console.log('ok annotation after-end recompute');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
