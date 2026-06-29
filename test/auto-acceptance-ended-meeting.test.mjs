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

function startFakeLark(meeting) {
  const server = createServer(async (req, res) => {
    const url = new URL(req.url, 'http://127.0.0.1');
    res.setHeader('content-type', 'application/json; charset=utf-8');
    if (req.method === 'POST' && url.pathname === '/open-apis/vc/v1/meetings/search') {
      for await (const _chunk of req) {
        // drain body
      }
      res.end(JSON.stringify({ code: 0, data: { items: [meeting] } }));
      return;
    }
    if (req.method === 'GET' && url.pathname === `/open-apis/vc/v1/meetings/${meeting.id}`) {
      res.end(JSON.stringify({ code: 0, data: { meeting } }));
      return;
    }
    res.statusCode = 404;
    res.end(JSON.stringify({ code: 404, msg: `not found: ${req.method} ${url.pathname}` }));
  });
  return server;
}

const appPort = await freePort();
const fakeLarkPort = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-auto-acceptance-ended-'));
const baseUrl = `http://127.0.0.1:${appPort}`;
const nowSeconds = Math.floor(Date.now() / 1000);
const endedMeeting = {
  id: 'ended-real-demo-meeting',
  topic: 'Ended real demo meeting',
  url: 'https://vc.feishu.cn/j/ended-real-demo',
  start_time: String(nowSeconds - 60),
  end_time: String(nowSeconds - 5),
};
const fakeLark = startFakeLark(endedMeeting);
fakeLark.listen(fakeLarkPort, '127.0.0.1');
await once(fakeLark, 'listening');

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
  user: { data: { user_id: 'test_user', open_id: 'ou_test', union_id: 'on_test' } },
  updated_at: new Date().toISOString(),
}, null, 2));

const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(appPort),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'auto-acceptance-ended-app',
    LARK_APP_SECRET: 'auto-acceptance-ended-secret',
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
  assert.equal(prepared.trigger.status, 'bound');

  const state = await (await fetch(`${baseUrl}/api/state`)).json();
  assert.equal(state.meeting.meeting_id, 'ended-real-demo-meeting');
  assert.equal(state.meeting.source, 'lark_passive_meeting_scan');
  assert.ok(state.meeting.end_time);
  assert.equal(state.sequence.length, 0);

  const auto = await (await fetch(`${baseUrl}/api/acceptance/auto-annotation`)).json();
  assert.equal(auto.count, 0);
  assert.equal(auto.last_annotation_id, null);

  console.log('ok auto acceptance skips ended meeting');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  fakeLark.close();
  await once(fakeLark, 'close').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
