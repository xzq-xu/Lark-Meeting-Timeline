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

async function runEventReport(baseUrl, args = []) {
  const child = spawn(process.execPath, [
    'scripts/event-delivery-report.mjs',
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

const tempDir = await mkdtemp(join(tmpdir(), 'lark-event-delivery-report-'));
const server = createServer((request, response) => {
  const path = new URL(request.url, 'http://127.0.0.1').pathname;
  if (path === '/api/lark/delivery-diagnostics') {
    sendJson(response, {
      status: 'waiting_for_lark_delivery',
      receiver: {
        ws_state: 'connected',
        ws_enabled: true,
        registered_event_types: [
          'vc.meeting.all_meeting_started_v1',
          'vc.meeting.all_meeting_ended_v1',
        ],
        required_direct_events: ['vc.meeting.all_meeting_started_v1'],
        fallback_context_events: ['vc.meeting.join_meeting_v1'],
        missing_registered_events: [],
        missing_fallback_events: [],
        http_callback_public_https: false,
      },
      evidence: {
        ws_event_count: 0,
        real_event_count: 0,
        public_webhook_event_count: 0,
        ws_timeline_candidate_count: 0,
        ws_timeline_processed_count: 0,
        ws_timeline_started_count: 0,
      },
      real_meeting_event_audit: {
        status: 'no_event_delivery_observed',
        local_handlers_ready: true,
        parser_self_test_passed: true,
        next_action: 'Open a real Lark meeting and wait for event delivery.',
        required_open_platform_checks: [
          { id: 'platform_delivery', ok: false, label: 'meeting event delivered' },
        ],
      },
    });
    return;
  }
  if (path === '/api/lark/events-log') {
    sendJson(response, { items: [] });
    return;
  }
  if (path === '/api/lark/real-demo/progress') {
    sendJson(response, {
      completion_evidence: {
        real_meeting_axis_active: false,
        event_axis_built: false,
        real_axis_annotation_count: 0,
        last_broadcast_real_axis_annotation_count: 0,
      },
      result: {
        operator_runbook: {
          primary_next_action: 'Keep the page open, then start a real meeting.',
        },
      },
    });
    return;
  }
  if (path === '/api/lark/real-demo/acceptance') {
    sendJson(response, {
      product_acceptance_complete: false,
      strict_event_acceptance_complete: false,
    });
    return;
  }
  response.writeHead(404, { 'content-type': 'application/json; charset=utf-8' });
  response.end('{"error":"not found"}\n');
});

try {
  const baseUrl = await listen(server);
  const reportFile = join(tempDir, 'event-delivery-report.json');

  const result = await runEventReport(baseUrl, [`--report-file=${reportFile}`]);
  assert.equal(result.code, 0, result.stderr);
  assert.equal(result.json.type, 'event_delivery_report');
  assert.equal(result.json.ok, true);
  assert.equal(result.json.status, 'waiting_for_lark_delivery');
  assert.equal(result.json.local_receiver_ready, true);
  assert.equal(result.json.local_handlers_ready, true);
  assert.equal(result.json.parser_self_test_passed, true);
  assert.equal(result.json.event_delivery_seen, false);
  assert.equal(result.json.real_axis_active, false);
  assert.equal(result.json.receiver.ws_state, 'connected');
  assert.equal(result.json.evidence.ws_event_count, 0);
  assert.equal(result.json.open_platform_checks[0].id, 'platform_delivery');

  const report = JSON.parse(await readFile(reportFile, 'utf8'));
  assert.equal(report.type, 'event_delivery_report');
  assert.equal(report.ok, true);
  assert.equal(report.report.file, reportFile);
  assert.equal(report.report.command, 'scripts/event-delivery-report.mjs');
  assert.equal(report.report.exit_code, 0);
  assert.equal(report.report.error, null);

  const strictResult = await runEventReport(baseUrl, ['--fail-on-no-delivery=true']);
  assert.equal(strictResult.code, 2);
  assert.equal(strictResult.json.ok, false);
  assert.equal(strictResult.json.status, 'waiting_for_lark_delivery');
  assert.equal(strictResult.json.event_delivery_seen, false);

  console.log('ok event delivery report');
} finally {
  const closed = once(server, 'close').catch(() => {});
  server.close();
  await closed;
  await rm(tempDir, { recursive: true, force: true });
}
