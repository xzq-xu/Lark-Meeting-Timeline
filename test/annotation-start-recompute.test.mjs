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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-annotation-start-recompute-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'annotation-start-recompute-app',
    LARK_APP_SECRET: 'annotation-start-recompute-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://annotation-start-recompute.example.com/api/lark/events',
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
  const trueStartSeconds = 1_782_442_800;
  const wrongStartSeconds = trueStartSeconds - 100;

  await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'start-recompute-wrong-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(wrongStartSeconds),
    },
    event: {
      meeting: {
        id: 'start-recompute-real-meeting',
        topic: 'Start recompute meeting',
      },
    },
  });

  const mark = await postJson(baseUrl, '/api/annotations', {
    id: 'mark-before-real-end-after-start-correction',
    source: 'hanwang_epaper',
    captured_at_ms: trueStartSeconds * 1000 + 30_000,
    kind: 'handwriting_trigger',
    label: '会中真实标注',
    text_candidates: ['会中真实标注'],
  });
  assert.equal(mark.item.time_ms, 130_000);
  assert.equal(mark.ack.on_real_axis, true);

  const correctedStart = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'start-recompute-correct-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(trueStartSeconds),
    },
    event: {
      meeting: {
        id: 'start-recompute-real-meeting',
        topic: 'Start recompute meeting',
        start_time: String(trueStartSeconds),
      },
    },
  });
  const rebased = correctedStart.state.sequence.find((item) => item.id === 'mark-before-real-end-after-start-correction');
  assert.equal(correctedStart.state.meeting.start_time, '2026-06-26T03:00:00.000Z');
  assert.equal(rebased.time_ms, 30_000);
  assert.equal(rebased.time_source, 'captured_at');

  const ended = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'start-recompute-end',
      event_type: 'vc.meeting.all_meeting_ended_v1',
      create_time: String(trueStartSeconds + 60),
    },
    event: {
      meeting: {
        id: 'start-recompute-real-meeting',
        topic: 'Start recompute meeting',
        start_time: String(trueStartSeconds),
        end_time: String(trueStartSeconds + 60),
      },
    },
  });
  const finalMark = ended.state.sequence.find((item) => item.id === 'mark-before-real-end-after-start-correction');
  const endEvent = ended.state.events.find((event) => event.id === 'start-recompute-end');
  assert.equal(finalMark.time_ms, 30_000);
  assert.equal(endEvent.time_ms, 60_000);
  assert.ok(finalMark.time_ms < endEvent.time_ms, 'mark should remain before meeting end after start correction');

  const bindings = await getJson(baseUrl, '/api/annotation-bindings');
  const binding = bindings.items.find((item) => item.id === 'mark-before-real-end-after-start-correction');
  assert.equal(binding.on_real_axis, true);
  assert.equal(binding.after_meeting_end_ms, 0);
  assert.equal(bindings.real_axis_count, 1);

  console.log('ok annotation start-time recompute');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
