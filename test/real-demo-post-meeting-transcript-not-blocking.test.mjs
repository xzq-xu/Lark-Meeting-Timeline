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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-real-demo-post-meeting-transcript-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'real-demo-post-meeting-transcript-app',
    LARK_APP_SECRET: 'real-demo-post-meeting-transcript-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://real-demo-post-meeting-transcript.example.com/api/lark/events',
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
    device_simulator: false,
    device_stream: false,
    passive_scan: false,
  });

  const startSeconds = 1_782_442_800;
  const endSeconds = startSeconds + 60;
  await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'post-meeting-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'post-meeting-real-meeting',
        topic: 'Post meeting transcript is not blocking',
        url: 'https://vc.feishu.cn/j/post-meeting-transcript',
        start_time: String(startSeconds),
      },
    },
  });

  const annotation = await postJson(baseUrl, '/api/annotations', {
    id: 'post-meeting-live-mark',
    source: 'future_epaper_device',
    device_id: 'hanwang-devkit-post-meeting',
    captured_at_ms: startSeconds * 1000 + 12_000,
    kind: 'handwriting_trigger',
    label: '会中实时标注',
    text_candidates: ['会中实时标注'],
  });
  assert.equal(annotation.ack.binding_state, 'real_meeting_bound');
  assert.equal(annotation.ack.on_real_axis, true);

  const completedBeforeEnd = await getJson(baseUrl, '/api/lark/real-demo/status');
  assert.equal(completedBeforeEnd.complete, true);
  assert.equal(completedBeforeEnd.gates.real_axis_annotation_count, 1);

  const ended = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'post-meeting-end',
      event_type: 'vc.meeting.all_meeting_ended_v1',
      create_time: String(endSeconds),
    },
    event: {
      meeting: {
        id: 'post-meeting-real-meeting',
        topic: 'Post meeting transcript is not blocking',
        start_time: String(startSeconds),
        end_time: String(endSeconds),
      },
    },
  });
  assert.equal(ended.ok, true);
  assert.equal(ended.state.meeting.end_time, '2026-06-26T03:01:00.000Z');
  assert.equal(ended.state.sequence.find((item) => item.id === 'post-meeting-live-mark')?.time_ms, 12_000);

  const completedAfterEnd = await getJson(baseUrl, '/api/lark/real-demo/status');
  assert.equal(completedAfterEnd.status, 'complete');
  assert.equal(completedAfterEnd.complete, true);
  assert.equal(completedAfterEnd.gates.real_meeting_axis_active, true);
  assert.equal(completedAfterEnd.gates.real_axis_annotation_count, 1);
  assert.equal(completedAfterEnd.evidence.annotation_binding.items[0].on_real_axis, true);

  const transcriptStatus = await getJson(baseUrl, '/api/transcript-status');
  assert.equal(transcriptStatus.meeting_ended, true);
  assert.equal(transcriptStatus.post_meeting, true);
  assert.equal(transcriptStatus.has_transcript, false);
  assert.equal(transcriptStatus.realtime_blocking, false);
  assert.equal(transcriptStatus.status, 'ended_no_minute_token');

  const readiness = await getJson(baseUrl, '/api/readiness');
  const transcriptCheck = readiness.checks.find((check) => check.id === 'post_meeting_transcript');
  assert.equal(transcriptCheck.ok, true);
  assert.equal(readiness.current.transcript_status.realtime_blocking, false);

  console.log('ok real demo post-meeting transcript not blocking');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
