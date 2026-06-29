import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-timeline-fallback-'));
const baseUrl = `http://127.0.0.1:${port}`;
const now = Date.now();
await writeFile(join(tempDir, 'lark-auth.json'), JSON.stringify({
  oauth_state: null,
  token: {
    access_token: 'test-user-token',
    refresh_token: 'test-refresh-token',
    expires_in: 7200,
    refresh_expires_in: 7200,
    scope: 'minutes:minutes.search:read minutes:minutes.basic:read minutes:minutes.transcript:export vc:meeting.search:read',
    obtained_at_ms: now,
  },
  user: {
    data: {
      user_id: 'test_user',
      open_id: 'ou_test',
      name: 'Test User',
    },
  },
  updated_at: new Date(now).toISOString(),
}, null, 2));

const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'fallback-test-app',
    LARK_APP_SECRET: 'fallback-test-secret',
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
  await postJson(baseUrl, '/api/import/sequence', { sequence: [] });
  await postJson(baseUrl, '/api/import/lark-transcript', {
    meeting: {
      platform: 'lark',
      meeting_id: 'scan-bound-meeting',
      title: 'Scan fallback meeting',
      start_time: '2026-06-26T03:00:00.000Z',
      end_time: null,
      source: 'lark_meeting_search_api',
    },
    transcript: [],
  });
  await postJson(baseUrl, '/api/annotations', {
    id: 'fallback-ann-1',
    source: 'hanwang_epaper',
    captured_at_ms: Date.parse('2026-06-26T03:01:00.000Z'),
    kind: 'handwriting_trigger',
    label: 'why?',
    text_candidates: ['why?', 'why'],
  });

  const readiness = await (await fetch(`${baseUrl}/api/readiness`)).json();
  assert.equal(readiness.ready, true);
  assert.equal(readiness.checks.find((check) => check.id === 'direct_meeting_event_seen')?.ok, false);
  assert.equal(readiness.checks.find((check) => check.id === 'real_lark_event_seen')?.ok, false);
  assert.equal(readiness.checks.find((check) => check.id === 'real_meeting_entry')?.ok, true);
  assert.equal(readiness.checks.find((check) => check.id === 'real_annotation_seen')?.ok, true);
  assert.equal(readiness.blockers.some((blocker) => blocker.id === 'direct_meeting_event_seen'), false);

  const diagnostics = await (await fetch(`${baseUrl}/api/lark/delivery-diagnostics`)).json();
  assert.equal(diagnostics.evidence.real_event_count, 0);
  assert.equal(diagnostics.evidence.real_meeting_axis_active, true);
  assert.equal(diagnostics.evidence.open_annotation_count, 1);
  assert.equal(diagnostics.open_platform_checklist.find((item) => item.id === 'event_delivery_evidence')?.status, 'blocked');

  console.log('ok scan fallback readiness');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
