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
          access_token: 'refreshed-user-token',
          refresh_token: 'refreshed-refresh-token',
          expires_in: 7200,
          refresh_expires_in: 7200,
          scope: 'vc:meeting.search:read minutes:minutes.basic:read',
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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-real-demo-prepare-'));
const baseUrl = `http://127.0.0.1:${appPort}`;
const startSeconds = Math.floor(Date.now() / 1000) - 15;
const fakeMeeting = {
  id: 'real-demo-prepare-meeting',
  topic: 'Real demo prepare meeting',
  url: 'https://vc.feishu.cn/j/real-demo-prepare',
  start_time: String(startSeconds),
};
const fakeLark = startFakeLark(fakeMeeting);
fakeLark.server.listen(fakeLarkPort, '127.0.0.1');
await once(fakeLark.server, 'listening');

await writeFile(join(tempDir, 'lark-auth.json'), JSON.stringify({
  oauth_state: null,
  token: {
    access_token: 'expired-user-token',
    refresh_token: 'test-refresh-token',
    expires_in: 1,
    refresh_expires_in: 7200,
    scope: 'vc:meeting.search:read minutes:minutes.basic:read',
    obtained_at_ms: Date.now() - 2000,
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
    LARK_APP_ID: 'real-demo-prepare-app',
    LARK_APP_SECRET: 'real-demo-prepare-secret',
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

  const prepared = await postJson(baseUrl, '/api/lark/real-demo/prepare', {
    auto_annotation: true,
    passive_scan: true,
  });
  assert.equal(prepared.prepared, true);
  assert.equal(prepared.auth_required, false);
  assert.equal(prepared.auth_refresh.refreshed, true);
  assert.equal(prepared.auto_acceptance.enabled, true);
  assert.equal(prepared.device_simulator.enabled, true);
  assert.equal(prepared.passive_meeting_scan.enabled, true);
  assert.equal(prepared.real_demo_session.active, true);
  assert.ok(prepared.real_demo_session.prepared_at);
  assert.equal(prepared.temporary_axis_reset.reset, true);
  assert.equal(prepared.trigger.status, 'bound');
  assert.equal(prepared.trigger.selected_meeting_id, 'real-demo-prepare-meeting');

  const state = await (await fetch(`${baseUrl}/api/state`)).json();
  assert.equal(state.meeting.meeting_id, 'real-demo-prepare-meeting');
  assert.equal(state.meeting.source, 'lark_passive_meeting_scan');
  assert.equal(state.sequence.length, 2);
  assert.equal(state.sequence.some((item) => /^auto-acceptance-/.test(item.id) && item.source === 'demo_auto_acceptance'), true);
  assert.equal(state.sequence.some((item) => /^device-simulator-/.test(item.id) && item.source === 'hanwang_epaper_simulator'), true);

  const readiness = await (await fetch(`${baseUrl}/api/readiness`)).json();
  assert.equal(readiness.current.annotation_binding.real_axis_count, 2);
  assert.equal(readiness.real_demo_session.active, true);
  assert.equal(readiness.real_demo_session.last_real_axis_source, 'lark_passive_meeting_scan');
  assert.equal(readiness.device_simulator.last_meeting_id, 'real-demo-prepare-meeting');

  const realDemoStatus = await (await fetch(`${baseUrl}/api/lark/real-demo/status`)).json();
  assert.equal(realDemoStatus.status, 'complete');
  assert.equal(realDemoStatus.complete, true);
  assert.equal(realDemoStatus.gates.real_meeting_axis_active, true);
  assert.equal(realDemoStatus.gates.real_axis_annotation_count, 2);
  assert.equal(fakeLark.calls.some((call) => call.pathname === '/open-apis/authen/v2/oauth/token'), true);
  assert.equal(fakeLark.calls.some((call) => call.pathname === '/open-apis/vc/v1/meetings/search'), true);

  console.log('ok real demo prepare binds meeting and writes auto annotation');
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
