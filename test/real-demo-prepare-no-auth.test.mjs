import assert from 'node:assert/strict';
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

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-real-demo-prepare-no-auth-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'real-demo-no-auth-app',
    LARK_APP_SECRET: 'real-demo-no-auth-secret',
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

  const before = await (await fetch(`${baseUrl}/api/state`)).json();
  const prepared = await postJson(baseUrl, '/api/lark/real-demo/prepare', {
    auto_annotation: true,
    passive_scan: true,
  });
  const after = await (await fetch(`${baseUrl}/api/state`)).json();
  const readiness = await (await fetch(`${baseUrl}/api/readiness`)).json();

  assert.equal(prepared.prepared, true);
  assert.equal(prepared.auth_required, true);
  assert.equal(prepared.trigger, null);
  assert.equal(prepared.auth_start.scopes.includes('vc:meeting.search:read'), true);
  assert.match(prepared.auth_start.redirect_url, /\/api\/auth\/lark\/start\?/);
  assert.match(prepared.auth_start.redirect_url, /redirect=1/);
  assert.equal(prepared.auth_start.method, 'GET');
  assert.equal(prepared.real_demo_session.active, true);
  assert.equal(prepared.temporary_axis_reset.reset, true);
  assert.equal(prepared.device_simulator.enabled, true);
  assert.equal(after.meeting.meeting_id, before.meeting.meeting_id);
  assert.equal(after.meeting.source, before.meeting.source);
  assert.equal(after.sequence.length, before.sequence.length);
  assert.equal(after.presentation.hide_timeline, true);
  assert.equal(after.presentation.axis_status, 'waiting_real_meeting_start');
  assert.equal(after.presentation.hidden_reason, 'demo_sample_axis_is_not_current_real_demo');
  assert.equal(after.real_demo_session.active, true);
  assert.equal(readiness.real_demo_session.active, true);
  assert.equal(readiness.current.annotation_binding.binding_state, 'demo_ignored');

  const realDemoStatus = await (await fetch(`${baseUrl}/api/lark/real-demo/status`)).json();
  assert.equal(realDemoStatus.status, 'needs_setup');
	  assert.equal(realDemoStatus.complete, false);
	  assert.equal(realDemoStatus.gates.real_demo_prepared, true);
	  assert.equal(realDemoStatus.gates.device_simulator_enabled, true);
	  assert.equal(realDemoStatus.gates.real_event_seen, false);
	  assert.equal(realDemoStatus.auth_start.scopes.includes('vc:meeting.search:read'), true);
	  assert.match(realDemoStatus.auth_start.redirect_url, /\/api\/auth\/lark\/start\?/);
	  assert.match(realDemoStatus.auth_start.redirect_url, /redirect=1/);
	  assert.equal(realDemoStatus.operator_runbook.oauth_scan_required_for_main_path, false);
	  assert.equal(realDemoStatus.blockers.some((item) => item.id === 'event_receiver'), true);
  const scanOauthBlocker = realDemoStatus.blockers.find((item) => item.id === 'meeting_scan_oauth');
  assert.ok(scanOauthBlocker);
  assert.match(scanOauthBlocker.action_url, /\/api\/auth\/lark\/start\?/);
  assert.match(scanOauthBlocker.oauth_redirect_url, /\/api\/auth\/lark\/start\?/);
  assert.match(scanOauthBlocker.oauth_json_url, /\/api\/auth\/lark\/start\?/);
  assert.match(scanOauthBlocker.permission_url, /open\.feishu\.cn\/app\//);
  const transcriptStatus = await (await fetch(`${baseUrl}/api/transcript-status`)).json();
  assert.equal(transcriptStatus.status, 'waiting_real_meeting');
  assert.equal(transcriptStatus.realtime_blocking, false);
  assert.equal(transcriptStatus.has_transcript, false);
  assert.equal(transcriptStatus.segment_count, 0);
  assert.equal(transcriptStatus.ignored_sample_segment_count > 0, true);
  assert.match(transcriptStatus.next_action, /样例转写不计入本次验收/);

  console.log('ok real demo prepare without auth does not create axis');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
