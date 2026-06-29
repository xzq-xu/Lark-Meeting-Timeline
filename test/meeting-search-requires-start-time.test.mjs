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

function startFakeLark() {
  const calls = [];
  const server = createServer(async (req, res) => {
    const url = new URL(req.url, 'http://127.0.0.1');
    calls.push({ method: req.method, pathname: url.pathname });
    res.setHeader('content-type', 'application/json; charset=utf-8');
    if (req.method === 'POST' && url.pathname === '/open-apis/auth/v3/tenant_access_token/internal') {
      for await (const _chunk of req) {
        // drain request body
      }
      res.end(JSON.stringify({
        code: 0,
        tenant_access_token: 'fake-tenant-token',
        expire: 7200,
      }));
      return;
    }
    if (req.method === 'POST' && url.pathname === '/open-apis/vc/v1/meetings/search') {
      for await (const _chunk of req) {
        // drain request body
      }
      res.end(JSON.stringify({
        code: 0,
        data: {
          items: [{
            id: 'meeting-without-start-time',
            topic: 'Missing start time meeting',
            url: 'https://vc.feishu.cn/j/missing-start-time',
          }],
        },
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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-meeting-search-requires-start-'));
const baseUrl = `http://127.0.0.1:${appPort}`;
const fakeLark = startFakeLark();
fakeLark.server.listen(fakeLarkPort, '127.0.0.1');
await once(fakeLark.server, 'listening');

const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(appPort),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'meeting-search-requires-start-app',
    LARK_APP_SECRET: 'meeting-search-requires-start-secret',
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

  const response = await fetch(`${baseUrl}/api/lark/bind-tenant-latest-meeting`, {
    method: 'POST',
    headers: { 'content-type': 'application/json; charset=utf-8' },
    body: JSON.stringify({ lookback_seconds: 120, lookahead_seconds: 120 }),
  });
  const json = await response.json();
  assert.equal(response.status, 422);
  assert.match(json.error, /does not include start_time/);
  assert.equal(fakeLark.calls.some((call) => call.pathname === '/open-apis/vc/v1/meetings/search'), true);

  const state = await (await fetch(`${baseUrl}/api/state`)).json();
  assert.notEqual(state.meeting.meeting_id, 'meeting-without-start-time');
  assert.notEqual(state.meeting.source, 'lark_tenant_meeting_search_api');

  console.log('ok meeting search requires start time');
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
