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

async function waitForJson(baseUrl, path, predicate, timeoutMs = 10_000) {
  const deadline = Date.now() + timeoutMs;
  let last = null;
  while (Date.now() < deadline) {
    const response = await fetch(`${baseUrl}${path}`);
    last = await response.json();
    assert.equal(response.ok, true, `${path} failed: ${JSON.stringify(last)}`);
    if (predicate(last)) return last;
    await new Promise((resolve) => setTimeout(resolve, 120));
  }
  throw new Error(`condition not observed for ${path}: ${JSON.stringify(last)}`);
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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-passive-tenant-cooldown-'));
const baseUrl = `http://127.0.0.1:${appPort}`;
await writeFile(join(tempDir, 'passive-meeting-scan.json'), JSON.stringify({
  enabled: true,
  interval_ms: 5000,
  tenant_fallback_cooldown_ms: 60000,
  lookback_seconds: 600,
  lookahead_seconds: 120,
}, null, 2));

const fakeLark = startFakeLark();
fakeLark.server.listen(fakeLarkPort, '127.0.0.1');
await once(fakeLark.server, 'listening');

const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(appPort),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'passive-tenant-cooldown-app',
    LARK_APP_SECRET: 'passive-tenant-cooldown-secret',
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

  const failed = await waitForJson(baseUrl, '/api/lark/passive-meeting-scan', (body) => (
    body.last_result?.reason === 'tenant_search_failed'
  ));
  assert.equal(failed.last_result.tenant_fallback, 'failed');
  assert.equal(fakeLark.calls.filter((call) => call.pathname === '/open-apis/vc/v1/meetings/search').length, 1);

  const cooled = await waitForJson(baseUrl, '/api/lark/passive-meeting-scan', (body) => (
    body.last_result?.reason === 'tenant_search_cooldown'
  ), 9000);
  assert.equal(cooled.last_result.tenant_fallback, 'cooldown');
  assert.equal(cooled.last_result.auth_mode, 'tenant_access_token');
  assert.equal(typeof cooled.last_result.next_allowed_at, 'string');
  assert.equal(fakeLark.calls.filter((call) => call.pathname === '/open-apis/vc/v1/meetings/search').length, 1);

  await new Promise((resolve) => setTimeout(resolve, 6200));
  const stillCooled = await waitForJson(baseUrl, '/api/lark/passive-meeting-scan', (body) => (
    body.last_result?.reason === 'tenant_search_cooldown'
      && body.last_result?.at !== cooled.last_result.at
  ), 9000);
  assert.equal(stillCooled.last_result.tenant_fallback, 'cooldown');
  assert.equal(fakeLark.calls.filter((call) => call.pathname === '/open-apis/vc/v1/meetings/search').length, 1);

  console.log('ok passive tenant scan cooldown');
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
