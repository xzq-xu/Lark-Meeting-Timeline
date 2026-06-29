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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-meeting-info-shape-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'meeting-info-shape-app',
    LARK_APP_SECRET: 'meeting-info-shape-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://meeting-info-shape.example.com/api/lark/events',
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
  const result = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'meeting-info-direct-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: '1782442800',
    },
    event: {
      meeting_info: {
        open_meeting_id: 'om_info_server_001',
        meeting_no: '987654321',
        topic: 'Server meeting_info shape',
        url: 'https://vc.feishu.cn/j/meeting-info-shape',
        begin_time: '1782442800',
      },
    },
  });

  assert.equal(result.timeline_started, true);
  assert.equal(result.state.meeting.meeting_id, 'om_info_server_001');
  assert.equal(result.state.meeting.external_meeting_id, 'om_info_server_001');
  assert.equal(result.state.meeting.title, 'Server meeting_info shape');
  assert.equal(result.state.meeting.meeting_url, 'https://vc.feishu.cn/j/meeting-info-shape');
  assert.equal(result.state.meeting.start_time, '2026-06-26T03:00:00.000Z');
  assert.equal(result.state.events.find((event) => event.id === 'meeting-info-direct-start')?.time_ms, 0);

  const annotation = await postJson(baseUrl, '/api/annotations', {
    id: 'meeting-info-before-end-mark',
    captured_at_ms: 1_782_442_830_000,
    kind: 'handwriting_trigger',
    label: '结束前的标注',
  });
  assert.equal(annotation.state.sequence.find((item) => item.id === 'meeting-info-before-end-mark')?.time_ms, 30_000);

  const ended = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'meeting-info-direct-end',
      event_type: 'vc.meeting.all_meeting_ended_v1',
      create_time: '1782442860',
    },
    event: {
      meeting_info: {
        open_meeting_id: 'om_info_server_001',
        meeting_no: '987654321',
        topic: 'Server meeting_info shape',
        url: 'https://vc.feishu.cn/j/meeting-info-shape',
        end_time: '1782442860',
      },
    },
  });
  assert.equal(ended.state.meeting.start_time, '2026-06-26T03:00:00.000Z');
  assert.equal(ended.state.meeting.end_time, '2026-06-26T03:01:00.000Z');
  assert.equal(ended.state.events.find((event) => event.id === 'meeting-info-direct-end')?.time_ms, 60_000);
  assert.equal(ended.state.sequence.find((item) => item.id === 'meeting-info-before-end-mark')?.time_ms, 30_000);

  console.log('ok meeting_info event shape');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
