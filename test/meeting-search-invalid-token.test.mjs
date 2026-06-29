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
        // drain
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
        // drain
      }
      res.statusCode = 401;
      res.end(JSON.stringify({
        code: 99991663,
        msg: 'Invalid access token for authorization. Please make a request with token attached.',
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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-search-invalid-token-'));
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
    LARK_APP_ID: 'invalid-token-app',
    LARK_APP_SECRET: 'invalid-token-secret',
    LARK_BASE_URL: `http://127.0.0.1:${fakeLarkPort}`,
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
  const response = await fetch(`${baseUrl}/api/lark/search-tenant-meetings`, {
    method: 'POST',
    headers: { 'content-type': 'application/json; charset=utf-8' },
    body: JSON.stringify({ page_size: 1 }),
  });
  const json = await response.json();
  assert.equal(response.status, 401);
  assert.equal(json.search_status, 'invalid_token');
  assert.match(json.next_action, /当前用户 OAuth/);

  const diagnostics = await (await fetch(`${baseUrl}/api/lark/delivery-diagnostics`)).json();
  assert.equal(diagnostics.active_search.status, 'oauth_login_required');
  assert.equal(diagnostics.active_search.api_status, 'invalid_token');
  assert.equal(diagnostics.active_search.tenant_search_supported, false);
  assert.match(diagnostics.active_search.next_action, /重新登录飞书账号/);

  const readiness = await (await fetch(`${baseUrl}/api/readiness`)).json();
  assert.match(readiness.checks.find((check) => check.id === 'meeting_search_permission')?.detail, /尚未登录飞书账号/);

  console.log('ok meeting search invalid token diagnostics');
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
