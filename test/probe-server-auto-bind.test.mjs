import assert from 'node:assert/strict';
import { createServer } from 'node:http';
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

async function waitForState(baseUrl, predicate, timeoutMs = 6000) {
  const deadline = Date.now() + timeoutMs;
  let last = null;
  while (Date.now() < deadline) {
    last = await (await fetch(`${baseUrl}/api/state`)).json();
    if (predicate(last)) return last;
    await new Promise((resolve) => setTimeout(resolve, 120));
  }
  throw new Error(`state predicate not met: ${JSON.stringify(last?.meeting ?? null)}`);
}

function startFakeLark(baseMeeting) {
  const calls = [];
  const server = createServer(async (req, res) => {
    const url = new URL(req.url, 'http://127.0.0.1');
    calls.push({ method: req.method, pathname: url.pathname });
    res.setHeader('content-type', 'application/json; charset=utf-8');
    if (req.method === 'POST' && url.pathname === '/open-apis/vc/v1/meetings/search') {
      for await (const _chunk of req) {
        // drain
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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-probe-server-auto-bind-'));
const baseUrl = `http://127.0.0.1:${appPort}`;
const startSeconds = Math.floor(Date.now() / 1000);
const fakeMeeting = {
  id: 'server-loop-real-meeting',
  topic: 'Server loop real meeting',
  url: 'https://vc.feishu.cn/j/server-loop-real',
  start_time: String(startSeconds),
};
const fakeLark = startFakeLark(fakeMeeting);
fakeLark.server.listen(fakeLarkPort, '127.0.0.1');
await once(fakeLark.server, 'listening');

await writeFile(join(tempDir, 'lark-auth.json'), JSON.stringify({
  oauth_state: null,
  token: {
    access_token: 'test-user-token',
    refresh_token: 'test-refresh-token',
    expires_in: 7200,
    refresh_expires_in: 7200,
    scope: 'vc:meeting.search:read',
    obtained_at_ms: Date.now(),
  },
  user: {
    data: {
      user_id: 'test_user',
      open_id: 'ou_test',
      union_id: 'on_test',
      name: 'Test User',
    },
  },
  updated_at: new Date().toISOString(),
}, null, 2));

const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(appPort),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'probe-server-auto-bind-app',
    LARK_APP_SECRET: 'probe-server-auto-bind-secret',
    LARK_BASE_URL: `http://127.0.0.1:${fakeLarkPort}`,
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
  const started = await postJson(baseUrl, '/api/lark/real-meeting-probe/start', {
    timeout_ms: 180_000,
    note: 'server loop auto bind test',
    reset_temporary_axis: true,
    auto_search: true,
    auto_search_interval_ms: 1000,
  });
  assert.equal(started.status, 'waiting');
  assert.equal(started.auto_search.server_loop.scheduled, true);

  await postJson(baseUrl, '/api/annotations', {
    id: 'server-loop-ann',
    source: 'hanwang_epaper',
    captured_at_ms: startSeconds * 1000 + 30_000,
    kind: 'handwriting_trigger',
    label: 'server loop mark',
    text_candidates: ['server loop mark'],
  });

  const boundState = await waitForState(baseUrl, (state) => (
    state.meeting?.meeting_id === 'server-loop-real-meeting'
      && state.meeting?.source === 'lark_probe_auto_search'
  ));
  const mark = boundState.sequence.find((item) => item.id === 'server-loop-ann');
  assert.equal(mark.time_ms, 30_000);

  const probe = await (await fetch(`${baseUrl}/api/lark/real-meeting-probe`)).json();
  assert.equal(probe.auto_search.last_result.status, 'bound');
  assert.equal(probe.auto_search.server_loop.scheduled, false);
  assert.equal(fakeLark.calls.some((call) => call.pathname === '/open-apis/vc/v1/meetings/search'), true);

  console.log('ok server-side probe auto bind loop');
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
