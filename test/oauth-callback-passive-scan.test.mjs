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

    if (req.method === 'POST' && url.pathname === '/open-apis/authen/v2/oauth/token') {
      for await (const _chunk of req) {
        // drain request body
      }
      res.end(JSON.stringify({
        code: 0,
        data: {
          access_token: 'new-user-token',
          refresh_token: 'new-refresh-token',
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
          user_id: 'test_user',
          open_id: 'ou_test',
          union_id: 'on_test',
          name: 'OAuth User',
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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-oauth-callback-passive-scan-'));
const baseUrl = `http://127.0.0.1:${appPort}`;
const startSeconds = Math.floor(Date.now() / 1000) - 20;
const fakeMeeting = {
  id: 'oauth-callback-real-meeting',
  topic: 'OAuth callback real meeting',
  url: 'https://vc.feishu.cn/j/oauth-callback-real',
  start_time: String(startSeconds),
};
const fakeLark = startFakeLark(fakeMeeting);
fakeLark.server.listen(fakeLarkPort, '127.0.0.1');
await once(fakeLark.server, 'listening');

await writeFile(join(tempDir, 'passive-meeting-scan.json'), JSON.stringify({
  enabled: false,
  tenant_fallback_enabled: false,
  interval_ms: 300_000,
  lookback_seconds: 600,
  lookahead_seconds: 120,
  last_attempt_at: new Date().toISOString(),
  last_result: {
    status: 'skipped',
    reason: 'oauth_token_expired',
    at: new Date().toISOString(),
  },
  updated_at: new Date().toISOString(),
}, null, 2));

const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(appPort),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'oauth-callback-scan-app',
    LARK_APP_SECRET: 'oauth-callback-scan-secret',
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

  const start = await (await fetch(`${baseUrl}/api/auth/lark/start?scope=vc%3Ameeting%3Areadonly`)).json();
  const laterStart = await (await fetch(`${baseUrl}/api/auth/lark/start?scope=vc%3Ameeting%3Areadonly`)).json();
  assert.equal(start.scope_present['vc:meeting.search:read'], true);
  assert.notEqual(start.state, laterStart.state);

  const bogusCallback = await fetch(`${baseUrl}/api/auth/lark/callback?code=test-code&state=not-a-real-state`);
  assert.equal(bogusCallback.status, 400);
  assert.match(await bogusCallback.text(), /state 不匹配/);

  const callback = await fetch(`${baseUrl}/api/auth/lark/callback?code=test-code&state=${encodeURIComponent(start.state)}`);
  assert.equal(callback.ok, true);
  const callbackHtml = await callback.text();
  assert.match(callbackHtml, /lark-auth-complete/);
  assert.match(callbackHtml, /passive_scan/);
  assert.match(callbackHtml, /oauth-callback-real-meeting/);

  const boundState = await waitForState(baseUrl, (state) => (
    state.meeting?.meeting_id === 'oauth-callback-real-meeting'
      && state.meeting?.source === 'lark_passive_meeting_scan'
  ));
  assert.equal(boundState.meeting.title, 'OAuth callback real meeting');

  const passive = await (await fetch(`${baseUrl}/api/lark/passive-meeting-scan`)).json();
  assert.equal(passive.enabled, true);
  assert.equal(passive.tenant_fallback_enabled, false);
  assert.equal(passive.last_result.status, 'bound');
  assert.equal(passive.last_result.selected_meeting_id, 'oauth-callback-real-meeting');
  assert.equal(fakeLark.calls.some((call) => call.pathname === '/open-apis/vc/v1/meetings/search'), true);

  console.log('ok OAuth callback triggers passive meeting scan');
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
