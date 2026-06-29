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

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-real-event-audit-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'real-event-audit-app',
    LARK_APP_SECRET: 'real-event-audit-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://real-event-audit.example.com/api/lark/events',
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

  const config = await (await fetch(`${baseUrl}/api/lark/config`)).json();
  assert.equal(config.real_meeting_event_audit.status, 'no_event_delivery_observed');
  assert.equal(config.real_meeting_event_audit.receiver_ready, true);
  assert.equal(config.real_meeting_event_audit.parser_self_test_passed, true);
  assert.equal(config.real_meeting_event_audit.local_handlers_ready, true);
  assert.equal(config.real_meeting_event_audit.event_delivery_seen, false);
  assert.match(config.real_meeting_event_audit.next_action, /直接开启一次飞书会议/);

  const diagnostics = await (await fetch(`${baseUrl}/api/lark/delivery-diagnostics`)).json();
  assert.equal(diagnostics.real_meeting_event_audit.status, 'no_event_delivery_observed');
  assert.equal(diagnostics.audit_status, 'no_event_delivery_observed');
  assert.equal(diagnostics.root_cause, 'none_until_a_real_meeting_is_opened');
  assert.equal(diagnostics.real_meeting_event_audit.required_open_platform_checks.find((item) => item.id === 'platform_delivery')?.ok, false);
  assert.equal(diagnostics.real_meeting_event_audit.required_open_platform_checks.find((item) => item.id === 'subscription_saved_and_published')?.ok, false);
  assert.match(diagnostics.real_meeting_event_audit.next_action, /事件订阅方式已保存为长连接/);
  assert.equal(diagnostics.open_platform_checklist.some((item) => item.id === 'event_subscription_publish'), true);

  console.log('ok real event audit');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
