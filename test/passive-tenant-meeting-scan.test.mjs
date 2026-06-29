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

function startFakeLark(baseMeeting) {
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
        data: { items: [baseMeeting] },
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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-passive-tenant-meeting-scan-'));
const baseUrl = `http://127.0.0.1:${appPort}`;
const startSeconds = Math.floor(Date.now() / 1000) - 20;
const fakeMeeting = {
  id: 'tenant-passive-real-meeting',
  topic: 'Tenant passive real meeting',
  url: 'https://vc.feishu.cn/j/tenant-passive-real',
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
    LARK_APP_ID: 'passive-tenant-scan-app',
    LARK_APP_SECRET: 'passive-tenant-scan-secret',
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

  const scan = await postJson(baseUrl, '/api/lark/passive-meeting-scan', {
    trigger_now: true,
    lookback_seconds: 120,
    lookahead_seconds: 120,
  });

  assert.equal(scan.trigger.status, 'bound');
  assert.equal(scan.trigger.auth_mode, 'tenant_access_token');
  assert.equal(scan.trigger.tenant_fallback, 'bound');
  assert.equal(scan.trigger.user_oauth_reason, 'oauth_login_required');
  assert.equal(scan.trigger.selected_meeting_id, 'tenant-passive-real-meeting');

  const state = await (await fetch(`${baseUrl}/api/state`)).json();
  assert.equal(state.meeting.meeting_id, 'tenant-passive-real-meeting');
  assert.equal(state.meeting.source, 'lark_tenant_passive_meeting_scan');
  assert.equal(state.events[0].label, '租户被动扫描建轴');

  const passiveStatus = await (await fetch(`${baseUrl}/api/lark/passive-meeting-scan`)).json();
  assert.equal(passiveStatus.last_result.status, 'bound');
  assert.equal(passiveStatus.last_result.auth_mode, 'tenant_access_token');
  assert.equal(passiveStatus.last_result.tenant_fallback, 'bound');

  const readiness = await (await fetch(`${baseUrl}/api/readiness`)).json();
  assert.equal(readiness.current.annotation_binding.binding_state, 'real_meeting_bound');
  assert.equal(readiness.current.meeting.source, 'lark_tenant_passive_meeting_scan');
  assert.equal(fakeLark.calls.some((call) => call.pathname === '/open-apis/vc/v1/meetings/search'), true);

  console.log('ok passive tenant meeting scan fallback');
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
