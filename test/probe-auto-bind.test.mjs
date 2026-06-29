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
    if (req.method === 'POST' && url.pathname === '/open-apis/vc/v1/meetings/search') {
      for await (const _chunk of req) {
        // drain
      }
      res.end(JSON.stringify({
        code: 0,
        data: {
          items: [baseMeeting],
        },
      }));
      return;
    }
    if (req.method === 'GET' && url.pathname === `/open-apis/vc/v1/meetings/${baseMeeting.id}`) {
      res.end(JSON.stringify({
        code: 0,
        data: {
          meeting: baseMeeting,
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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-probe-auto-bind-'));
const baseUrl = `http://127.0.0.1:${appPort}`;
const startSeconds = Math.floor(Date.now() / 1000);
const fakeMeeting = {
  id: 'auto-probe-real-meeting',
  topic: 'Auto probe real meeting',
  url: 'https://vc.feishu.cn/j/auto-probe-real',
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
    LARK_APP_ID: 'probe-auto-bind-app',
    LARK_APP_SECRET: 'probe-auto-bind-secret',
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
  await postJson(baseUrl, '/api/lark/real-meeting-probe/start', {
    timeout_ms: 180_000,
    note: 'probe auto bind test',
    reset_temporary_axis: true,
    auto_search: true,
  });
  await postJson(baseUrl, '/api/annotations', {
    id: 'probe-auto-bind-ann',
    source: 'hanwang_epaper',
    captured_at_ms: startSeconds * 1000 + 30_000,
    kind: 'handwriting_trigger',
    label: 'why?',
    text_candidates: ['why?', 'why'],
  });

  const bound = await postJson(baseUrl, '/api/lark/real-meeting-probe/auto-bind', {});
  assert.equal(bound.status, 'bound');
  assert.equal(bound.state.meeting.meeting_id, 'auto-probe-real-meeting');
  assert.equal(bound.state.meeting.source, 'lark_probe_auto_search');
  const mark = bound.state.sequence.find((item) => item.id === 'probe-auto-bind-ann');
  assert.equal(mark.time_ms, 30_000);

  const report = await (await fetch(`${baseUrl}/api/lark/acceptance-report`)).json();
  assert.equal(report.current_validation.auto_search_binding_after_probe, true);
  assert.equal(report.current_validation.real_entry_after_probe, true);
  assert.equal(report.current_validation.annotation_on_real_axis, true);
  assert.equal(report.current_validation.ready, true);
  const entryStep = report.acceptance_steps.find((step) => step.id === 'direct_start_meeting');
  assert.equal(entryStep.done, true);
  assert.equal(entryStep.evidence, 'probe_auto_search');

  assert.equal(fakeLark.calls.some((call) => call.pathname === '/open-apis/vc/v1/meetings/search'), true);
  console.log('ok probe auto bind fallback');
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
