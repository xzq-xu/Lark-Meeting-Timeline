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

async function getJson(baseUrl, path) {
  const response = await fetch(`${baseUrl}${path}`);
  const json = await response.json();
  assert.equal(response.ok, true, `${path} failed: ${JSON.stringify(json)}`);
  return json;
}

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-real-demo-progress-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'real-demo-progress-app',
    LARK_APP_SECRET: 'real-demo-progress-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://real-demo-progress.example.com/api/lark/events',
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

  await postJson(baseUrl, '/api/lark/real-demo/prepare', {
    auto_annotation: false,
    device_simulator: false,
    device_stream: false,
    passive_scan: false,
  });

  const waiting = await getJson(baseUrl, '/api/lark/real-demo/progress');
  assert.equal(waiting.observed, false);
  assert.equal(waiting.timed_out, false);
  assert.equal(waiting.completion_evidence.real_demo_complete, false);
  assert.equal(waiting.completion_evidence.real_meeting_axis_active, false);
  assert.equal(waiting.completion_evidence.meeting_source, null);
  assert.equal(waiting.completion_evidence.event_axis_built, false);
  assert.equal(waiting.completion_evidence.axis_creation_mode, 'not_built');
  assert.equal(waiting.completion_evidence.event_audit_status, 'no_event_delivery_observed');
  assert.equal(waiting.completion_evidence.ws_event_count, 0);
  assert.equal(waiting.result.operator_runbook.phase, 'open_lark_meeting');
  assert.equal(waiting.result.operator_runbook.can_open_meeting_now, true);
  assert.equal(waiting.result.operator_runbook.oauth_scan_required_for_main_path, false);
  assert.match(waiting.result.operator_runbook.annotation_endpoint, /\/api\/annotations$/);
  assert.equal(waiting.result.operator_runbook.success_condition.real_axis_annotation_count_min, 1);
  assert.equal(waiting.result.next_action, waiting.result.operator_runbook.primary_next_action);
  assert.equal(waiting.completion_evidence.next_action, waiting.result.operator_runbook.primary_next_action);
  assert.match(waiting.result.next_action, /直接开启飞书会议|开启飞书会议等待事件投递/);
  const waitingRequirements = Object.fromEntries(waiting.completion_evidence.requirements.map((item) => [item.id, item]));
  assert.equal(waitingRequirements.event_receiver_ready.ok, true);
  assert.equal(waitingRequirements.real_meeting_axis_active.ok, false);
  assert.equal(waitingRequirements.event_meeting_axis_built.ok, false);
  assert.equal(waitingRequirements.open_annotation_on_real_axis.ok, false);
  assert.equal(waitingRequirements.realtime_state_broadcast.ok, false);
  assert.equal(waitingRequirements.transcript_post_meeting_only.ok, true);

  const waitingAcceptance = await getJson(baseUrl, '/api/lark/real-demo/acceptance');
  assert.equal(waitingAcceptance.verdict, 'ready_to_open_meeting');
  assert.equal(waitingAcceptance.product_acceptance_complete, false);
  assert.equal(waitingAcceptance.strict_event_acceptance_complete, false);
  assert.equal(waitingAcceptance.event_delivery_required_for_product_acceptance, false);
  assert.equal(waitingAcceptance.can_open_meeting_now, true);
  assert.match(waitingAcceptance.recommended_command, /accept:real-meeting/);
  assert.equal(waitingAcceptance.meeting_axis.active, false);
  assert.equal(waitingAcceptance.missing_product_requirements.some((item) => item.id === 'real_meeting_axis_active'), true);
  assert.equal(waitingAcceptance.missing_product_requirements.some((item) => item.id === 'event_meeting_axis_built'), false);

  const startSeconds = 1_782_442_800;
  await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'progress-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'progress-real-meeting',
        topic: 'Real demo progress meeting',
        start_time: String(startSeconds),
      },
    },
  });

  const withAxis = await getJson(baseUrl, '/api/lark/real-demo/progress');
  assert.equal(withAxis.observed, false);
  assert.equal(withAxis.status, 'real_axis_waiting_annotation');
  assert.equal(withAxis.completion_evidence.real_meeting_axis_active, true);
  assert.equal(withAxis.completion_evidence.meeting_source, 'lark_http_event');
  assert.equal(withAxis.completion_evidence.event_axis_built, true);
  assert.equal(withAxis.completion_evidence.strict_event_axis_complete, false);
  assert.equal(withAxis.completion_evidence.axis_creation_mode, 'meeting_start_event');
  assert.equal(withAxis.completion_evidence.real_axis_annotation_count, 0);
  assert.equal(withAxis.completion_evidence.event_audit_status, 'event_delivery_ok');
  assert.equal(withAxis.result.operator_runbook.phase, 'write_open_annotation');
  const withAxisRequirements = Object.fromEntries(withAxis.completion_evidence.requirements.map((item) => [item.id, item]));
  assert.equal(withAxisRequirements.event_meeting_axis_built.ok, true);

  const axisAcceptance = await getJson(baseUrl, '/api/lark/real-demo/acceptance');
  assert.equal(axisAcceptance.verdict, 'waiting_for_open_annotation');
  assert.equal(axisAcceptance.meeting_axis.active, true);
  assert.equal(axisAcceptance.meeting_axis.event_axis, true);
  assert.equal(axisAcceptance.missing_product_requirements.some((item) => item.id === 'open_annotation_on_real_axis'), true);

  await postJson(baseUrl, '/api/annotations', {
    id: 'progress-open-ann',
    source: 'future_epaper_device',
    captured_at_ms: startSeconds * 1000 + 18_000,
    kind: 'handwriting_trigger',
    label: 'why?',
    text_candidates: ['why?', 'why'],
  });

  const complete = await getJson(baseUrl, '/api/lark/real-demo/progress');
  assert.equal(complete.observed, true);
  assert.equal(complete.status, 'complete');
  assert.equal(complete.completion_evidence.real_demo_complete, true);
  assert.equal(complete.completion_evidence.strict_event_axis_complete, true);
  assert.equal(complete.completion_evidence.real_axis_annotation_count, 1);
  assert.equal(complete.completion_evidence.real_axis_annotation_count_total, 1);
  assert.equal(complete.completion_evidence.event_audit_status, 'event_delivery_ok');
  assert.equal(complete.completion_evidence.last_broadcast_real_axis_annotation_count, 1);
  assert.equal(complete.result.operator_runbook.phase, 'complete');
  const completedRequirements = Object.fromEntries(complete.completion_evidence.requirements.map((item) => [item.id, item]));
  assert.equal(Object.values(completedRequirements).every((item) => item.ok), true);

  const completeAcceptance = await getJson(baseUrl, '/api/lark/real-demo/acceptance');
  assert.equal(completeAcceptance.verdict, 'pass');
  assert.equal(completeAcceptance.product_acceptance_complete, true);
  assert.equal(completeAcceptance.strict_event_acceptance_complete, true);
  assert.equal(completeAcceptance.recommended_command, null);
  assert.equal(completeAcceptance.realtime_annotation.real_axis_annotation_count, 1);

  console.log('ok real demo progress endpoint');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
