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

async function waitForJson(baseUrl, path, predicate, timeoutMs = 5_000) {
  const deadline = Date.now() + timeoutMs;
  let last = null;
  while (Date.now() < deadline) {
    const response = await fetch(`${baseUrl}${path}`);
    last = await response.json();
    assert.equal(response.ok, true, `${path} failed: ${JSON.stringify(last)}`);
    if (predicate(last)) return last;
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
  throw new Error(`condition not met for ${path}: ${JSON.stringify(last)}`);
}

function startFakeLark(getItems) {
  const calls = [];
  const server = createServer(async (req, res) => {
    const url = new URL(req.url, 'http://127.0.0.1');
    const call = {
      method: req.method,
      pathname: url.pathname,
      search: Object.fromEntries(url.searchParams.entries()),
      at: Date.now(),
    };
    res.setHeader('content-type', 'application/json; charset=utf-8');
    if (req.method === 'POST' && url.pathname === '/open-apis/vc/v1/meetings/search') {
      let text = '';
      for await (const _chunk of req) {
        text += _chunk;
      }
      call.body = text ? JSON.parse(text) : {};
      calls.push(call);
      res.end(JSON.stringify({ code: 0, data: { items: getItems() } }));
      return;
    }
    calls.push(call);
    const meetingId = decodeURIComponent(url.pathname.split('/').at(-1) || '');
    const meeting = getItems().find((item) => String(item.id) === meetingId);
    if (req.method === 'GET' && url.pathname === `/open-apis/vc/v1/meetings/${meetingId}` && meeting) {
      res.end(JSON.stringify({ code: 0, data: { meeting } }));
      return;
    }
    res.statusCode = 404;
    res.end(JSON.stringify({ code: 404, msg: `not found: ${req.method} ${url.pathname}` }));
  });
  return { server, calls };
}

const appPort = await freePort();
const fakeLarkPort = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-startup-fast-passive-scan-'));
const baseUrl = `http://127.0.0.1:${appPort}`;
const startSeconds = Math.floor(Date.now() / 1000) + 1;
let meetingVisible = false;
const fakeMeeting = {
  id: 'startup-fast-passive-scan-meeting',
  topic: 'Startup fast passive scan meeting',
  url: 'https://vc.feishu.cn/j/startup-fast-passive',
  start_time: String(startSeconds),
};
const fakeLark = startFakeLark(() => (meetingVisible ? [fakeMeeting] : []));
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
  user: { data: { user_id: 'test_user', open_id: 'ou_test', union_id: 'on_test', name: 'Test User' } },
  updated_at: new Date().toISOString(),
}, null, 2));

const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(appPort),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'startup-fast-passive-scan-app',
    LARK_APP_SECRET: 'startup-fast-passive-scan-secret',
    LARK_BASE_URL: `http://127.0.0.1:${fakeLarkPort}`,
    LARK_VERIFICATION_TOKEN: '',
    TIMELINE_DATA_DIR: tempDir,
    REAL_DEMO_AUTO_ARM: '1',
    REAL_DEMO_AUTO_ANNOTATION: '0',
    REAL_DEMO_DEVICE_SIMULATOR: '0',
    REAL_DEMO_DEVICE_STREAM: '0',
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

  const status = await waitForJson(baseUrl, '/api/lark/real-demo/status', (json) => (
    json.real_demo_session?.active === true && json.passive_meeting_scan?.enabled === true
  ));
  assert.equal(status.passive_meeting_scan.server_loop.interval_ms, 2000);

  meetingVisible = true;
  const state = await waitForJson(baseUrl, '/api/state', (json) => (
    json.meeting?.meeting_id === 'startup-fast-passive-scan-meeting'
      && json.meeting?.source === 'lark_passive_meeting_scan'
  ), 5_000);
  assert.equal(state.meeting.title, 'Startup fast passive scan meeting');
  const searchCalls = fakeLark.calls.filter((call) => call.pathname === '/open-apis/vc/v1/meetings/search');
  assert.equal(searchCalls.length >= 2, true);
  assert.equal(searchCalls[0].search.user_id_type, undefined);
  assert.deepEqual(searchCalls[0].body.meeting_filter.participant_ids, ['ou_test']);
  assert.match(searchCalls[0].body.meeting_filter.start_time.start_time, /T.*Z$/);
  assert.match(searchCalls[0].body.meeting_filter.start_time.end_time, /T.*Z$/);

  const scan = await (await fetch(`${baseUrl}/api/lark/passive-meeting-scan`)).json();
  assert.equal(scan.last_result.status, 'bound');
  assert.equal(scan.last_result.selected_meeting_id, 'startup-fast-passive-scan-meeting');

  console.log('ok startup fast passive scan fallback');
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
