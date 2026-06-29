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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-realtime-current-axis-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'realtime-current-axis-app',
    LARK_APP_SECRET: 'realtime-current-axis-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://realtime-current-axis.example.com/api/lark/events',
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
  await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'realtime-ended-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'realtime-ended-meeting',
        topic: 'Realtime ended meeting',
        start_time: String(startSeconds),
      },
    },
  });
  await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'realtime-ended-end',
      event_type: 'vc.meeting.all_meeting_ended_v1',
      create_time: String(startSeconds + 60),
    },
    event: {
      meeting: {
        id: 'realtime-ended-meeting',
        topic: 'Realtime ended meeting',
        start_time: String(startSeconds),
        end_time: String(startSeconds + 60),
      },
    },
  });

  const capturedAtMs = (startSeconds * 1000) + 10 * 60_000;
  const realtime = await postJson(baseUrl, '/api/annotations', {
    id: 'realtime-after-ended-axis',
    source: 'browser_demo',
    realtime: true,
    mode: 'realtime',
    captured_at_ms: capturedAtMs,
    kind: 'live_mark',
    label: '此刻写下的标注',
    payload: {
      realtime: true,
      mode: 'realtime',
    },
  });

  assert.equal(realtime.auto_open_session_binding.status, 'started');
  assert.equal(realtime.state.meeting.source, 'open_meeting_session');
  assert.equal(realtime.state.meeting.meeting_id, `annotation-open-session-${capturedAtMs}`);
  assert.equal(realtime.item.time_ms, 0);
  assert.equal(realtime.ack.after_meeting_end_ms, 0);
  assert.equal(realtime.ack.warnings.includes('normalized_after_meeting_end'), false);

  await postJson(baseUrl, '/api/demo/reset', {});
  await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'realtime-stale-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'realtime-stale-old-meeting',
        topic: 'Realtime stale old meeting',
        start_time: String(startSeconds),
      },
    },
  });
  await postJson(baseUrl, '/api/lark/real-demo/prepare', {
    auto_annotation: false,
    device_simulator: false,
    device_stream: false,
    passive_scan: false,
  });

  const staleCapturedAtMs = (startSeconds * 1000) + 65 * 60_000;
  const staleRealtime = await postJson(baseUrl, '/api/annotations', {
    id: 'realtime-after-stale-axis',
    source: 'browser_demo',
    realtime: true,
    mode: 'realtime',
    captured_at_ms: staleCapturedAtMs,
    kind: 'live_mark',
    label: '当前写下而不是旧会议第65分钟',
    payload: {
      realtime: true,
      mode: 'realtime',
    },
  });
  assert.equal(staleRealtime.auto_open_session_binding.status, 'started');
  assert.equal(staleRealtime.state.meeting.source, 'open_meeting_session');
  assert.equal(staleRealtime.state.meeting.meeting_id, `annotation-open-session-${staleCapturedAtMs}`);
  assert.equal(staleRealtime.item.time_ms, 0);

  console.log('ok realtime annotation uses current axis');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
