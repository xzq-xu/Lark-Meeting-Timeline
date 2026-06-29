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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-end-anchor-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'end-anchor-app',
    LARK_APP_SECRET: 'end-anchor-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://end-anchor.example.com/api/lark/events',
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
  const pending = await postJson(baseUrl, '/api/annotations', {
    id: 'pending-before-real-end',
    source: 'hanwang_epaper',
    captured_at_ms: startSeconds * 1000 + 30_000,
    kind: 'handwriting_trigger',
    label: '会中写下的标注',
    text_candidates: ['会中写下的标注'],
  });
  assert.equal(pending.state.meeting.source, 'annotation_fallback');
  assert.equal(pending.state.meeting.pending_binding, true);

  const ignoredEnd = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'end-without-start-time',
      event_type: 'vc.meeting.all_meeting_ended_v1',
      create_time: String(startSeconds + 60),
    },
    event: {
      meeting: {
        id: 'real-meeting-without-start-time',
        topic: 'Missing start anchor',
      },
    },
  });
  assert.equal(ignoredEnd.ok, false);
  assert.equal(ignoredEnd.ignored_reason, 'meeting_end_without_known_start_axis');
  assert.equal(ignoredEnd.state.meeting.source, 'annotation_fallback');
  assert.equal(ignoredEnd.state.events.some((event) => event.type === 'meeting_end'), false);

  const logAfterIgnored = await getJson(baseUrl, '/api/lark/events-log');
  const ignoredLog = logAfterIgnored.items.find((item) => item.event_type === 'vc.meeting.all_meeting_ended_v1');
  assert.equal(ignoredLog.timeline_processed, false);
  assert.equal(ignoredLog.ignored_reason, 'meeting_end_without_known_start_axis');

  const anchoredEnd = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'end-with-start-time',
      event_type: 'vc.meeting.all_meeting_ended_v1',
      create_time: String(startSeconds + 60),
    },
    event: {
      meeting: {
        id: 'real-meeting-with-start-time',
        topic: 'End event with start anchor',
        start_time: String(startSeconds),
        end_time: String(startSeconds + 60),
      },
    },
  });
  assert.equal(anchoredEnd.ok, true);
  assert.equal(anchoredEnd.timeline_started, true);
  assert.equal(anchoredEnd.state.meeting.source, 'lark_http_event');
  assert.equal(anchoredEnd.state.meeting.pending_binding, false);
  assert.equal(anchoredEnd.state.meeting.start_time, '2026-06-26T03:00:00.000Z');
  assert.equal(anchoredEnd.state.meeting.end_time, '2026-06-26T03:01:00.000Z');
  const startEvent = anchoredEnd.state.events.find((event) => event.type === 'meeting_start');
  const endEvent = anchoredEnd.state.events.find((event) => event.id === 'end-with-start-time');
  const mark = anchoredEnd.state.sequence.find((item) => item.id === 'pending-before-real-end');
  assert.equal(startEvent.time_ms, 0);
  assert.equal(startEvent.metadata.raw_type, 'inferred_from_meeting_end_start_time');
  assert.equal(endEvent.time_ms, 60_000);
  assert.equal(mark.time_ms, 30_000);
  assert.equal(mark.time_source, 'captured_at');
  assert.ok(mark.time_ms < endEvent.time_ms, 'pending mark should be before meeting end after rebasing');

  const lateUnreliable = await postJson(baseUrl, '/api/annotations', {
    id: 'late-without-captured-at',
    source: 'hanwang_epaper',
    kind: 'handwriting_trigger',
    label: '会后才上传且缺少采集时间',
    text_candidates: ['会后才上传且缺少采集时间'],
  });
  assert.equal(lateUnreliable.ack.binding_state, 'real_meeting_bound');
  assert.equal(lateUnreliable.ack.timing_reliable, false);
  assert.equal(lateUnreliable.ack.on_real_axis, false);
  assert.ok(lateUnreliable.ack.warnings.includes('missing_captured_at_ms'));
  assert.ok(lateUnreliable.ack.warnings.includes('normalized_after_meeting_end'));
  const bindingsAfterLate = await getJson(baseUrl, '/api/annotation-bindings');
  assert.equal(bindingsAfterLate.real_axis_count, 1);
  const lateBinding = bindingsAfterLate.items.find((entry) => entry.id === 'late-without-captured-at');
  assert.equal(lateBinding.timing_reliable, false);
  assert.equal(lateBinding.on_real_axis, false);

  console.log('ok end event start anchor guard');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
