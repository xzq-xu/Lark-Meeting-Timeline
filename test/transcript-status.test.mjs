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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-transcript-status-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'transcript-status-app',
    LARK_APP_SECRET: 'transcript-status-secret',
    LARK_VERIFICATION_TOKEN: '',
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

  await postJson(baseUrl, '/api/live/start-meeting', {
    title: 'Transcript status live meeting',
    meeting_id: 'transcript-status-live',
    minute_token: 'minute-status-token',
    start_time: '2026-06-26T03:00:00.000Z',
  });
  const liveStatus = await getJson(baseUrl, '/api/transcript-status');
  assert.equal(liveStatus.status, 'live_no_transcript_expected');
  assert.equal(liveStatus.realtime_blocking, false);
  assert.equal(liveStatus.minute_token_present, true);
  assert.equal(liveStatus.segment_count, 0);

  await postJson(baseUrl, '/api/live/end-meeting', {
    time_ms: 300_000,
  });
  const endedStatus = await getJson(baseUrl, '/api/transcript-status');
  assert.equal(endedStatus.status, 'ready_to_sync_minute');
  assert.equal(endedStatus.meeting_ended, true);
  assert.equal(endedStatus.sync_endpoint, '/api/lark/sync-minute');

  await postJson(baseUrl, '/api/import/lark-transcript', {
    meeting: {
      meeting_id: 'transcript-status-live',
      start_time: '2026-06-26T03:00:00.000Z',
      minute_token: 'minute-status-token',
    },
    transcript: [{
      id: 'seg-imported-1',
      start_ms: 1000,
      end_ms: 4000,
      speaker_name: 'Tester',
      text: '会后导入的转写',
    }],
  });
  const syncedStatus = await getJson(baseUrl, '/api/transcript-status');
  assert.equal(syncedStatus.status, 'synced');
  assert.equal(syncedStatus.has_transcript, true);
  assert.equal(syncedStatus.segment_count, 1);

  const readiness = await getJson(baseUrl, '/api/readiness');
  assert.equal(readiness.current.transcript_status.status, 'synced');
  const transcriptCheck = readiness.checks.find((check) => check.id === 'post_meeting_transcript');
  assert.equal(transcriptCheck.ok, true);
  assert.match(transcriptCheck.detail, /转写已导入/);

  console.log('ok transcript status');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
