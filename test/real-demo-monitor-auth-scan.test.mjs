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

async function getJson(baseUrl, path) {
  const response = await fetch(`${baseUrl}${path}`);
  const json = await response.json();
  assert.equal(response.ok, true, `${path} failed: ${JSON.stringify(json)}`);
  return json;
}

async function runMonitor(baseUrl, extraArgs = []) {
  const child = spawn(process.execPath, ['scripts/monitor-real-demo.mjs', `--url=${baseUrl}`, '--once', ...extraArgs], {
    cwd: new URL('..', import.meta.url),
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  let stdout = '';
  let stderr = '';
  child.stdout.on('data', (chunk) => {
    stdout += chunk.toString();
  });
  child.stderr.on('data', (chunk) => {
    stderr += chunk.toString();
  });
  const [code] = await once(child, 'exit');
  return { code, stdout, stderr };
}

function startFakeLark(baseMeeting) {
  const calls = [];
  const server = createServer(async (req, res) => {
    const url = new URL(req.url, 'http://127.0.0.1');
    calls.push({ method: req.method, pathname: url.pathname });
    res.setHeader('content-type', 'application/json; charset=utf-8');
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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-real-demo-monitor-auth-scan-'));
const baseUrl = `http://127.0.0.1:${appPort}`;
const startSeconds = Math.floor(Date.now() / 1000) - 20;
const fakeMeeting = {
  id: 'monitor-auth-scan-real-meeting',
  topic: 'Monitor auth scan real meeting',
  url: 'https://vc.feishu.cn/j/monitor-auth-scan',
  start_time: String(startSeconds),
};
const fakeLark = startFakeLark(fakeMeeting);
fakeLark.server.listen(fakeLarkPort, '127.0.0.1');
await once(fakeLark.server, 'listening');

await writeFile(join(tempDir, 'lark-auth.json'), JSON.stringify({
  oauth_state: null,
  token: {
    access_token: 'valid-user-token',
    refresh_token: 'valid-refresh-token',
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

await writeFile(join(tempDir, 'passive-meeting-scan.json'), JSON.stringify({
  enabled: false,
  interval_ms: 300_000,
  lookback_seconds: 600,
  lookahead_seconds: 120,
  last_attempt_at: null,
  last_result: null,
  updated_at: new Date().toISOString(),
}, null, 2));

const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(appPort),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'real-demo-monitor-auth-scan-app',
    LARK_APP_SECRET: 'real-demo-monitor-auth-scan-secret',
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

  const result = await runMonitor(baseUrl, ['--wait-auth', '--auth-timeout-ms=1000']);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /auth_wait \| ready=yes/);
  assert.match(result.stdout, /auth_scan_triggered=true/);
  assert.match(result.stdout, /status=bound/);
  assert.match(result.stdout, /meeting=monitor-auth-scan-real-meeting/);
  assert.match(result.stdout, /source=lark_passive_meeting_scan/);
  assert.match(result.stdout, /real_axis=yes/);
  assert.match(result.stdout, /axis_source=lark_passive_meeting_scan/);
  assert.match(result.stdout, /annotations=0/);

  const state = await getJson(baseUrl, '/api/state');
  assert.equal(state.meeting.meeting_id, 'monitor-auth-scan-real-meeting');
  assert.equal(state.meeting.source, 'lark_passive_meeting_scan');
  assert.equal(state.presentation.real_axis_active, true);
  assert.equal(fakeLark.calls.some((call) => call.pathname === '/open-apis/vc/v1/meetings/search'), true);

  console.log('ok real demo monitor auth scan');
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
