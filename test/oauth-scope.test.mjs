import assert from 'node:assert/strict';
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

function scopesFromAuthUrl(authUrl) {
  return new URL(authUrl).searchParams.get('scope')?.split(/\s+/).filter(Boolean) ?? [];
}

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-timeline-oauth-scope-'));
const baseUrl = `http://127.0.0.1:${port}`;
await writeFile(join(tempDir, 'lark-auth.json'), JSON.stringify({
  oauth_state: null,
  token: {
    access_token: 'old-token',
    refresh_token: null,
    expires_in: 1,
    refresh_expires_in: 0,
    scope: 'minutes:minutes.search:read minutes:minutes.basic:read',
    obtained_at_ms: Date.now() - 60_000,
  },
  user: {
    data: {
      open_id: 'ou_old',
      name: 'Old Token User',
    },
  },
  updated_at: new Date().toISOString(),
}, null, 2));
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'oauth-scope-app',
    LARK_APP_SECRET: 'oauth-scope-secret',
    LARK_REDIRECT_URI: `${baseUrl}/api/auth/lark/callback`,
    LARK_OAUTH_SCOPES: 'minutes:minutes.search:read minutes:minutes.basic:read vc:meeting.search:read',
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

  const normal = await (await fetch(`${baseUrl}/api/auth/lark/start`)).json();
  assert.deepEqual(scopesFromAuthUrl(normal.auth_url), [
    'minutes:minutes.search:read',
    'minutes:minutes.basic:read',
    'vc:meeting.search:read',
  ]);

  const minutesOnly = await (await fetch(`${baseUrl}/api/auth/lark/start?purpose=minutes`)).json();
  assert.deepEqual(scopesFromAuthUrl(minutesOnly.auth_url), [
    'minutes:minutes.search:read',
    'minutes:minutes.basic:read',
    'minutes:minutes.transcript:export',
  ]);
  assert.equal(minutesOnly.scope_present['vc:meeting.search:read'], false);
  assert.equal(minutesOnly.scope_mode, 'minutes');

  const expanded = await (await fetch(`${baseUrl}/api/auth/lark/start?scope=vc%3Ameeting.search%3Aread%20minutes%3Aminutes.basic%3Aread`)).json();
  assert.deepEqual(scopesFromAuthUrl(expanded.auth_url), [
    'minutes:minutes.search:read',
    'minutes:minutes.basic:read',
    'vc:meeting.search:read',
  ]);
  assert.deepEqual(expanded.scopes, [
    'minutes:minutes.search:read',
    'minutes:minutes.basic:read',
    'vc:meeting.search:read',
  ]);
  assert.equal(expanded.requested_scope, 'vc:meeting.search:read minutes:minutes.basic:read');
  assert.equal(expanded.redirect_uri, `${baseUrl}/api/auth/lark/callback`);
  assert.equal(expanded.callback_url, `${baseUrl}/api/auth/lark/callback`);
  assert.equal(typeof expanded.state_created_at, 'string');
  assert.equal(typeof expanded.state_expires_at, 'string');
  assert.equal(expanded.scope_present['vc:meeting.search:read'], true);

  const redirect = await fetch(`${baseUrl}/api/auth/lark/start?scope=vc%3Ameeting.search%3Aread&redirect=1`, {
    redirect: 'manual',
  });
  assert.equal(redirect.status, 302);
  const location = redirect.headers.get('location');
  assert.ok(location);
  assert.deepEqual(scopesFromAuthUrl(location), [
    'minutes:minutes.search:read',
    'minutes:minutes.basic:read',
    'vc:meeting.search:read',
  ]);

  const status = await (await fetch(`${baseUrl}/api/auth/lark/status`)).json();
  assert.equal(status.authenticated, false);
  assert.equal(status.expired, true);
  assert.equal(status.meeting_search.usable, false);
  assert.equal(status.meeting_search.scope_present, false);
  assert.equal(status.meeting_search.needs_reauth, true);
  assert.equal(status.meeting_search.reason, 'oauth_token_expired');

  console.log('ok OAuth dynamic scope');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
