import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
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
const tempDir = await mkdtemp(join(tmpdir(), 'local-detector-reconcile-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
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

  const startMs = 1_782_442_800_000;
  const startIso = new Date(startMs).toISOString();
  const meetingUrl = 'https://meet.google.com/local-demo';

  const localStart = await postJson(baseUrl, '/api/platform-events/local-detector', {
    id: 'local-detector-start-001',
    type: 'meeting_started',
    detected_platform: 'google_meet',
    meeting_id: 'local-detector-google-meet-001',
    meeting_url: meetingUrl,
    title: 'Local detector Google Meet',
    start_time_ms: startMs,
  });
  assert.equal(localStart.results[0].action, 'startMeeting');
  assert.equal(localStart.state.meeting.source, 'local_detector');
  assert.equal(localStart.state.meeting.platform, 'google_meet');
  assert.equal(localStart.state.meeting.meeting_url, meetingUrl);

  const annotation = await postJson(baseUrl, '/api/annotations', {
    id: 'local-detector-ann-001',
    source: 'hanwang_epaper',
    captured_at_ms: startMs + 12_000,
    kind: 'handwriting_trigger',
    label: 'why?',
  });
  assert.equal(annotation.ack.on_real_axis, true);
  assert.equal(annotation.ack.meeting_source, 'local_detector');
  assert.equal(annotation.ack.normalized_time_ms, 12_000);

  const officialStart = await postJson(baseUrl, '/api/platform-events/google-meet', {
    id: 'google-start-after-local-detector',
    type: 'google.workspace.meet.conference.v2.started',
    time: startIso,
    data: {
      conferenceRecord: { name: 'conferenceRecords/google-record-reconcile-001' },
      meetingUri: meetingUrl,
      title: 'Official Google Meet',
    },
  });
  assert.equal(officialStart.results[0].action, 'startMeeting');
  assert.equal(officialStart.state.meeting.source, 'google_meet_webhook');
  assert.equal(officialStart.state.meeting.platform, 'google_meet');
  assert.equal(officialStart.state.meeting.meeting_id, 'google-record-reconcile-001');
  assert.equal(officialStart.state.sequence.some((item) => (
    item.id === 'local-detector-ann-001'
      && item.time_ms === 12_000
      && item.source === 'hanwang_epaper'
  )), true);

  const status = await getJson(baseUrl, '/api/platform-events/status');
  assert.equal(status.status.local_detector.received_count, 1);
  assert.equal(status.status.google_meet.received_count, 1);

  console.log('ok local detector reconciles with delayed provider event');
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
  if (process.exitCode) {
    console.error(output);
  }
}
