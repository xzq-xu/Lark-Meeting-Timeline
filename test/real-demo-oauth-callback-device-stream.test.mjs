import assert from 'node:assert/strict';
import { createServer } from 'node:http';
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

async function waitForJson(baseUrl, path, predicate, timeoutMs = 8_000) {
  const deadline = Date.now() + timeoutMs;
  let last = null;
  while (Date.now() < deadline) {
    last = await getJson(baseUrl, path);
    if (predicate(last)) return last;
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
  throw new Error(`condition not met for ${path}: ${JSON.stringify(last)}`);
}

async function waitForSseState(baseUrl, predicate, timeoutMs = 8_000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(`${baseUrl}/api/stream`, { signal: controller.signal });
    assert.equal(response.ok, true, 'SSE stream should open');
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const chunks = buffer.split('\n\n');
      buffer = chunks.pop() ?? '';
      for (const chunk of chunks) {
        const event = chunk.split('\n').find((line) => line.startsWith('event: '))?.slice(7);
        const data = chunk.split('\n').find((line) => line.startsWith('data: '))?.slice(6);
        if (event === 'state' && data) {
          const state = JSON.parse(data);
          if (predicate(state)) {
            await reader.cancel();
            return state;
          }
        }
      }
    }
  } finally {
    clearTimeout(timer);
  }
  throw new Error('SSE state predicate was not observed');
}

function startFakeLark(baseMeeting) {
  const calls = [];
  const server = createServer(async (req, res) => {
    const url = new URL(req.url, 'http://127.0.0.1');
    calls.push({ method: req.method, pathname: url.pathname });
    res.setHeader('content-type', 'application/json; charset=utf-8');

    if (req.method === 'POST' && url.pathname === '/open-apis/authen/v2/oauth/token') {
      for await (const _chunk of req) {
        // drain request body
      }
      res.end(JSON.stringify({
        code: 0,
        data: {
          access_token: 'real-demo-device-stream-user-token',
          refresh_token: 'real-demo-device-stream-refresh-token',
          expires_in: 7200,
          refresh_expires_in: 7200,
          token_type: 'Bearer',
          scope: 'vc:meeting.search:read minutes:minutes.basic:read',
        },
      }));
      return;
    }

    if (req.method === 'GET' && url.pathname === '/open-apis/authen/v1/user_info') {
      res.end(JSON.stringify({
        code: 0,
        data: {
          user_id: 'real_demo_user',
          open_id: 'ou_real_demo',
          union_id: 'on_real_demo',
          name: 'Real Demo User',
        },
      }));
      return;
    }

    if (req.method === 'POST' && url.pathname === '/open-apis/vc/v1/meetings/search') {
      for await (const _chunk of req) {
        // drain request body
      }
      res.end(JSON.stringify({
        code: 0,
        data: { items: [baseMeeting] },
      }));
      return;
    }

    if (req.method === 'GET' && url.pathname === `/open-apis/vc/v1/meetings/${baseMeeting.id}`) {
      res.end(JSON.stringify({
        code: 0,
        data: { meeting: baseMeeting },
      }));
      return;
    }

    res.statusCode = 404;
    res.end(JSON.stringify({ code: 404, msg: `not found: ${req.method} ${url.pathname}` }));
  });
  return { server, calls };
}

const appPort = await freePort();
const fakeLarkPort = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-real-demo-oauth-device-stream-'));
const baseUrl = `http://127.0.0.1:${appPort}`;
const startSeconds = Math.floor(Date.now() / 1000) - 25;
const fakeMeeting = {
  id: 'oauth-device-stream-real-meeting',
  topic: 'OAuth device stream real meeting',
  url: 'https://vc.feishu.cn/j/oauth-device-stream-real',
  start_time: String(startSeconds),
};
const fakeLark = startFakeLark(fakeMeeting);
fakeLark.server.listen(fakeLarkPort, '127.0.0.1');
await once(fakeLark.server, 'listening');

const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(appPort),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'real-demo-oauth-device-stream-app',
    LARK_APP_SECRET: 'real-demo-oauth-device-stream-secret',
    LARK_BASE_URL: `http://127.0.0.1:${fakeLarkPort}`,
    LARK_REDIRECT_URI: `${baseUrl}/api/auth/lark/callback`,
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

  await postJson(baseUrl, '/api/lark/real-demo/prepare', {
    auto_annotation: false,
    device_simulator: true,
    device_stream: true,
    passive_scan: true,
    device_stream_interval_ms: 250,
    device_stream_max_count: 2,
  });

  await waitForJson(baseUrl, '/api/device-simulator/stream', (stream) => (
    stream.enabled === true && stream.status === 'waiting_for_real_axis'
  ));

  const ssePromise = waitForSseState(baseUrl, (state) => (
    state.meeting?.meeting_id === 'oauth-device-stream-real-meeting'
      && state.meeting?.source === 'lark_passive_meeting_scan'
      && state.sequence?.some((item) => item.id === 'device-stream-oauth-device-stream-real-meeting-1')
      && state.presentation?.real_axis_active === true
  ));

  const start = await getJson(baseUrl, '/api/auth/lark/start?scope=vc%3Ameeting%3Areadonly');
  assert.equal(start.callback_url, `${baseUrl}/api/auth/lark/callback`);
  assert.equal(start.scope_present['vc:meeting.search:read'], true);

  const callback = await fetch(`${baseUrl}/api/auth/lark/callback?code=test-code&state=${encodeURIComponent(start.state)}`);
  assert.equal(callback.ok, true);
  const callbackHtml = await callback.text();
  assert.match(callbackHtml, /lark-auth-complete/);
  assert.match(callbackHtml, /oauth-device-stream-real-meeting/);

  const sseState = await ssePromise;
  const firstStreamMark = sseState.sequence.find((item) => item.id === 'device-stream-oauth-device-stream-real-meeting-1');
  assert.equal(firstStreamMark.source, 'hanwang_epaper_simulator');
  assert.equal(firstStreamMark.time_source, 'captured_at');

  const status = await waitForJson(baseUrl, '/api/lark/real-demo/status', (body) => (
    body.gates.real_meeting_axis_active === true
      && body.gates.real_axis_annotation_count > 0
  ));
  assert.equal(status.evidence.annotation_binding.binding_state, 'real_meeting_bound');

  const progress = await getJson(baseUrl, '/api/lark/real-demo/progress');
  assert.equal(progress.completion_evidence.real_meeting_axis_active, true);
  assert.equal(progress.completion_evidence.real_axis_annotation_count > 0, true);
  assert.equal(
    progress.completion_evidence.requirements.find((item) => item.id === 'transcript_post_meeting_only')?.ok,
    true,
  );
  assert.equal(progress.completion_evidence.event_axis_built, false);
  assert.equal(progress.completion_evidence.axis_creation_mode, 'lark_passive_meeting_scan');

  const passive = await getJson(baseUrl, '/api/lark/passive-meeting-scan');
  assert.equal(passive.last_result.status, 'bound');
  assert.equal(passive.last_result.selected_meeting_id, 'oauth-device-stream-real-meeting');
  assert.equal(fakeLark.calls.some((call) => call.pathname === '/open-apis/vc/v1/meetings/search'), true);

  console.log('ok real demo OAuth callback binds axis and streams device annotations');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  fakeLark.server.close();
  await once(fakeLark.server, 'close').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
