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

async function runOnsiteStatus(baseUrl, args = []) {
  const child = spawn(process.execPath, [
    'scripts/onsite-status.mjs',
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

const tempDir = await mkdtemp(join(tmpdir(), 'lark-onsite-status-'));
const server = createServer((request, response) => {
  const path = new URL(request.url, 'http://127.0.0.1').pathname;
  if (path === '/api/lark/real-demo/status') {
    sendJson(response, {
      complete: false,
      next_action: '重新授权扫描权限。',
      gates: {
        event_receiver_ready: true,
        open_annotation_ready: true,
        realtime_stream_ready: true,
        device_stream_enabled: true,
        real_meeting_axis_active: false,
        real_axis_annotation_count: 0,
      },
      device_stream_simulator: {
        enabled: true,
        status: 'waiting_for_real_axis',
      },
    });
    return;
  }
  if (path === '/api/lark/delivery-diagnostics') {
    sendJson(response, {
      status: 'waiting_for_lark_delivery',
      receiver: {
        ws_state: 'connected',
      },
      evidence: {
        ws_event_count: 0,
        real_event_count: 0,
        public_webhook_event_count: 0,
        ws_timeline_started_count: 0,
      },
      real_meeting_event_audit: {
        status: 'no_event_delivery_observed',
        local_handlers_ready: true,
        parser_self_test_passed: true,
        required_open_platform_checks: [
          { id: 'platform_delivery', ok: false },
        ],
      },
    });
    return;
  }
  if (path === '/api/lark/passive-meeting-scan') {
    sendJson(response, {
      enabled: true,
      tenant_fallback_enabled: false,
      server_loop: { enabled: true, scheduled: true },
      last_result: {
        status: 'skipped',
        reason: 'tenant_fallback_disabled',
        user_oauth_reason: 'oauth_token_expired',
      },
    });
    return;
  }
  if (path === '/api/auth/lark/status') {
    sendJson(response, {
      meeting_search: {
        usable: false,
        needs_reauth: true,
        reason: 'oauth_token_expired',
        required_scope: 'vc:meeting.search:read',
        scope_present: false,
      },
    });
    return;
  }
  if (path === '/api/stream/status') {
    sendJson(response, {
      current_clients: 1,
      broadcast_count: 2,
    });
    return;
  }
  if (path === '/api/annotation-ingest-info') {
    sendJson(response, {
      endpoint: `${request.headers.host}/api/annotations`,
      batch_endpoint: `${request.headers.host}/api/annotations/batch`,
      stream_url: `${request.headers.host}/api/stream`,
      real_demo_acceptance_url: `${request.headers.host}/api/lark/real-demo/acceptance`,
    });
    return;
  }
  response.writeHead(404, { 'content-type': 'application/json; charset=utf-8' });
  response.end('{"error":"not found"}\n');
});

try {
  const baseUrl = await listen(server);
  const reportFile = join(tempDir, 'onsite-status-report.json');

  const result = await runOnsiteStatus(baseUrl, [`--report-file=${reportFile}`]);
  assert.equal(result.code, 0, result.stderr);
  assert.equal(result.json.type, 'onsite_status');
  assert.equal(result.json.ok, true);
  assert.equal(result.json.status, 'ready_but_event_delivery_unproven');
  assert.equal(result.json.minimum_start_ready, true);
  assert.equal(result.json.practical_start_ready, false);
  assert.equal(result.json.gates.event_path_ready, true);
  assert.equal(result.json.gates.scan_fallback_enabled, true);
  assert.equal(result.json.gates.scan_fallback_ready, false);
  assert.equal(result.json.auth.reason, 'oauth_token_expired');
  assert.equal(result.json.event_delivery.failed_open_platform_checks[0], 'platform_delivery');

  const report = JSON.parse(await readFile(reportFile, 'utf8'));
  assert.equal(report.type, 'onsite_status');
  assert.equal(report.report.file, reportFile);
  assert.equal(report.report.command, 'scripts/onsite-status.mjs');
  assert.equal(report.report.exit_code, 0);

  const strict = await runOnsiteStatus(baseUrl, ['--require-practical-ready=true']);
  assert.equal(strict.code, 2);
  assert.equal(strict.json.ok, false);
  assert.equal(strict.json.minimum_start_ready, true);
  assert.equal(strict.json.practical_start_ready, false);

  console.log('ok onsite status');
} finally {
  const closed = once(server, 'close').catch(() => {});
  server.close();
  await closed;
  await rm(tempDir, { recursive: true, force: true });
}
