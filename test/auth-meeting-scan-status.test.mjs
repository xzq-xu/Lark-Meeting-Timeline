import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { createServer } from 'node:http';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

function sendJson(response, body) {
  response.writeHead(200, { 'content-type': 'application/json; charset=utf-8' });
  response.end(`${JSON.stringify(body)}\n`);
}

async function listen(server) {
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  const address = server.address();
  return `http://127.0.0.1:${address.port}`;
}

async function runAuthStatus(baseUrl, args = []) {
  const child = spawn(process.execPath, [
    'scripts/auth-meeting-scan-status.mjs',
    `--url=${baseUrl}`,
    '--json',
    ...args,
  ], {
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
  return {
    code,
    stdout: stdout.trim(),
    stderr: stderr.trim(),
    json: stdout.trim() ? JSON.parse(stdout) : null,
  };
}

let authReady = false;
let scanCalls = 0;
const tempDir = await mkdtemp(join(tmpdir(), 'lark-auth-meeting-scan-status-'));
const server = createServer(async (request, response) => {
  const url = new URL(request.url, 'http://127.0.0.1');
  if (url.pathname === '/api/auth/lark/start') {
    const callbackUrl = `http://${request.headers.host}/api/auth/lark/callback`;
    sendJson(response, {
      auth_url: `http://${request.headers.host}/fake-auth?app_id=test&state=secret-state&redirect_uri=${encodeURIComponent(callbackUrl)}&scope=vc%3Ameeting%3Areadonly`,
      redirect_uri: callbackUrl,
      requested_scope: 'vc:meeting.search:read',
      scopes: ['vc:meeting.search:read'],
      scope_present: { 'vc:meeting.search:read': true },
      state_created_at: '2026-06-27T00:00:00.000Z',
      state_expires_at: '2026-06-27T00:15:00.000Z',
    });
    return;
  }
  if (url.pathname === '/fake-auth') {
    response.writeHead(302, {
      location: `/fake-login?redirect_uri=${encodeURIComponent(`http://${request.headers.host}${url.pathname}${url.search}`)}`,
    });
    response.end();
    return;
  }
  if (url.pathname === '/fake-login') {
    response.writeHead(200, { 'content-type': 'text/html; charset=utf-8' });
    response.end('<!doctype html><title>fake login</title>');
    return;
  }
  if (url.pathname === '/api/auth/lark/status') {
    sendJson(response, authReady
      ? {
        authenticated: true,
        token_present: true,
        expired: false,
        refresh_token_present: true,
        meeting_search: {
          usable: true,
          scope_present: true,
          needs_reauth: false,
          required_scope: 'vc:meeting.search:read',
        },
      }
      : {
        authenticated: false,
        token_present: true,
        expired: true,
        refresh_token_present: false,
        meeting_search: {
          usable: false,
          scope_present: false,
          needs_reauth: true,
          reason: 'oauth_token_expired',
          required_scope: 'vc:meeting.search:read',
          next_action: '重新登录飞书账号，并授予 vc:meeting.search:read',
        },
      });
    return;
  }
  if (url.pathname === '/api/lark/passive-meeting-scan') {
    if (request.method === 'POST') {
      for await (const _chunk of request) {
        // drain
      }
      scanCalls += 1;
      sendJson(response, {
        passive_meeting_scan: {
          enabled: true,
          tenant_fallback_enabled: false,
          last_result: {
            status: 'bound',
            selected_meeting_id: 'auth-status-meeting',
          },
        },
        trigger: {
          status: 'bound',
          selected_meeting_id: 'auth-status-meeting',
        },
      });
      return;
    }
    sendJson(response, {
      enabled: authReady,
      tenant_fallback_enabled: false,
      server_loop: { enabled: authReady, scheduled: authReady },
      last_result: authReady
        ? { status: 'bound', selected_meeting_id: 'auth-status-meeting' }
        : { status: 'skipped', reason: 'tenant_fallback_disabled', user_oauth_reason: 'oauth_token_expired' },
    });
    return;
  }
  if (url.pathname === '/api/lark/real-demo/status') {
    sendJson(response, {
      status: authReady ? 'real_axis_waiting_annotation' : 'ready_to_open_meeting',
      ready_to_open_meeting: true,
      gates: {
        real_meeting_axis_active: authReady,
        real_axis_annotation_count: 0,
      },
    });
    return;
  }
  response.writeHead(404, { 'content-type': 'application/json; charset=utf-8' });
  response.end('{"error":"not found"}\n');
});

try {
  const baseUrl = await listen(server);
  const reportFile = join(tempDir, 'auth-status-report.json');

  const expired = await runAuthStatus(baseUrl, [`--report-file=${reportFile}`]);
  assert.equal(expired.code, 0);
  assert.equal(expired.json.type, 'auth_meeting_scan_status');
  assert.equal(expired.json.ok, false);
  assert.equal(expired.json.status, 'oauth_expired');
  assert.equal(expired.json.auth.reason, 'oauth_token_expired');
  assert.equal(expired.json.auth.scope_present, false);
  assert.equal(expired.json.auth_start.scope_present['vc:meeting.search:read'], true);
  assert.equal(expired.json.auth_start.auth_url.includes('secret-state'), false);
  assert.match(expired.json.auth_start.auth_url, /state=%3Credacted%3E|state=<redacted>/);
  assert.equal(expired.json.auth_url_check.callback_preserved, true);
  assert.equal(expired.json.auth_url_check.status, 'callback_preserved');
  assert.equal(expired.json.auth_url_check.redirect_count, 1);

  const report = JSON.parse(await readFile(reportFile, 'utf8'));
  assert.equal(report.type, 'auth_meeting_scan_status');
  assert.equal(report.report.file, reportFile);
  assert.equal(report.report.command, 'scripts/auth-meeting-scan-status.mjs');
  assert.equal(report.report.exit_code, 0);

  const strictExpired = await runAuthStatus(baseUrl, ['--require-ready=true']);
  assert.equal(strictExpired.code, 2);
  assert.equal(strictExpired.json.status, 'oauth_expired');

  authReady = true;
  const ready = await runAuthStatus(baseUrl);
  assert.equal(ready.code, 0, ready.stderr);
  assert.equal(ready.json.ok, true);
  assert.equal(ready.json.status, 'auth_ready');
  assert.equal(ready.json.auth.meeting_search_usable, true);
  assert.equal(ready.json.scan_after_auth.trigger.status, 'bound');
  assert.equal(scanCalls, 1);

  console.log('ok auth meeting scan status');
} finally {
  const closed = once(server, 'close').catch(() => {});
  server.close();
  await closed;
  await rm(tempDir, { recursive: true, force: true });
}
