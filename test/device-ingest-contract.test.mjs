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

async function postJson(baseUrl, path, body, headers = {}) {
  const response = await fetch(`${baseUrl}${path}`, {
    method: 'POST',
    headers: { 'content-type': 'application/json; charset=utf-8', ...headers },
    body: JSON.stringify(body),
  });
  const json = await response.json();
  assert.equal(response.ok, true, `${path} failed: ${JSON.stringify(json)}`);
  return json;
}

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-device-ingest-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'device-contract-app',
    LARK_APP_SECRET: 'device-contract-secret',
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

  const preflight = await fetch(`${baseUrl}/api/annotations`, {
    method: 'OPTIONS',
    headers: {
      origin: 'http://device.local',
      'access-control-request-method': 'POST',
      'access-control-request-headers': 'content-type,x-hmp-device-id,x-hmp-device-type',
    },
  });
  assert.equal(preflight.status, 204);
  assert.match(preflight.headers.get('access-control-allow-headers') ?? '', /x-hmp-device-id/);
  assert.match(preflight.headers.get('access-control-allow-headers') ?? '', /x-hmp-device-type/);

  const initialInfo = await (await fetch(`${baseUrl}/api/annotation-ingest-info`)).json();
  assert.equal(initialInfo.annotation_route.mode, 'create_pending_on_first_annotation');
  assert.equal(initialInfo.annotation_route.will_create_pending_on_first_annotation, true);
  assert.match(initialInfo.binding_lookup_url, /\/api\/annotation-bindings\?id=\{annotation_id\}$/);
  assert.match(initialInfo.real_demo_status_url, /\/api\/lark\/real-demo\/status$/);
  assert.match(initialInfo.real_demo_acceptance_url, /\/api\/lark\/real-demo\/acceptance$/);
  assert.match(initialInfo.real_demo_monitor_url, /\/api\/lark\/real-demo\/monitor$/);
  assert.equal(initialInfo.time_sync_url, `${baseUrl}/api/time`);
  assert.equal(initialInfo.server_time_url, `${baseUrl}/api/time`);
  assert.equal(initialInfo.time_url, `${baseUrl}/api/time`);
  assert.equal(initialInfo.stream_status_url, `${baseUrl}/api/stream/status`);
  assert.equal(initialInfo.stream_status_alias_url, `${baseUrl}/api/stream-status`);
  const streamStatus = await (await fetch(initialInfo.stream_status_url)).json();
  const streamStatusAlias = await (await fetch(initialInfo.stream_status_alias_url)).json();
  assert.equal(streamStatus.current_clients, streamStatusAlias.current_clients);
  assert.equal(streamStatus.broadcast_count, streamStatusAlias.broadcast_count);
  assert.equal(initialInfo.clock_sync.required, true);
  assert.equal(initialInfo.clock_sync.endpoint, `${baseUrl}/api/time`);
  assert.match(initialInfo.clock_sync.recommended_algorithm, /clock_offset_ms/);
  assert.equal(initialInfo.clock_sync.max_recommended_skew_ms, 500);
  const clientSendAtMs = Date.now();
  const timeSync = await (await fetch(`${initialInfo.time_sync_url}?client_send_at_ms=${clientSendAtMs}`)).json();
  assert.equal(typeof timeSync.server_time_ms, 'number');
  assert.equal(typeof timeSync.server_time, 'string');
  assert.equal(timeSync.client_send_at_ms, clientSendAtMs);
  assert.equal(typeof timeSync.server_received_at_ms, 'number');
  assert.equal(typeof timeSync.server_sent_at_ms, 'number');
  assert.equal(typeof timeSync.estimated_offset_at_receive_ms, 'number');
  assert.match(timeSync.device_midpoint_formula, /client_receive_at_ms/);
  assert.match(timeSync.captured_at_formula, /captured_at_ms/);
  assert.equal(initialInfo.device_client_contract.can_send_before_real_axis, true);
  assert.equal(initialInfo.device_client_contract.can_send_during_meeting, true);
  assert.equal(initialInfo.device_client_contract.can_send_after_meeting, true);
  assert.equal(initialInfo.device_client_contract.stable_id_required_for_idempotency, true);
  assert.equal(initialInfo.device_client_contract.clock_sync_required, true);
  assert.equal(initialInfo.device_client_contract.preferred_time_field, 'captured_at_ms');
  assert.match(initialInfo.device_client_contract.production_success_condition, /ack\.on_real_axis/);
  assert.match(initialInfo.device_client_contract.final_success_condition, /annotation-bindings/);
  assert.match(initialInfo.device_client_contract.product_acceptance_condition, /product_acceptance_complete/);
  assert.match(initialInfo.device_client_contract.product_acceptance_condition, /strict_event_acceptance_complete/);
  assert.ok(initialInfo.device_client_contract.should_alert_operator_when.includes('ack.requires_device_captured_at === true'));
  assert.equal(initialInfo.response_ack.object_path, 'ack');
  assert.equal(initialInfo.response_ack.fields.binding_state, 'real_meeting_bound | pending_real_meeting | local_simulation | demo_ignored | unbound');
  assert.equal(initialInfo.response_ack.fields.idempotent, 'boolean');
  assert.match(initialInfo.response_ack.fields.timing_reliable, /boolean/);
  assert.match(initialInfo.response_ack.fields.requires_device_captured_at, /boolean/);
  assert.match(initialInfo.response_ack.sibling_fields.passive_binding, /bind a real meeting/);
  assert.equal(initialInfo.real_demo_session.active, false);
  assert.equal(initialInfo.device_simulator.enabled, false);
  assert.equal(initialInfo.meeting_search_oauth.usable, false);
  assert.match(initialInfo.curl_examples.single, /curl -sS -X POST/);
  assert.match(initialInfo.curl_examples.single, /\/api\/annotations/);
  assert.match(initialInfo.curl_examples.single, /x-hmp-device-id: hanwang-alpha-001/);
  assert.match(initialInfo.curl_examples.single, /captured_at_ms/);
  assert.match(initialInfo.curl_examples.verify_status, /\/api\/annotation-status\?id=device-mark-001/);
  assert.match(initialInfo.curl_examples.watch_stream, /curl -N .*\/api\/stream/);
  assert.match(initialInfo.curl_examples.full_roundtrip, /ANNOTATION_ID='device-mark-001'/);
  assert.match(initialInfo.curl_examples.full_roundtrip, /pending_real_meeting \/ real_axis_bound/);
  assert.match(initialInfo.curl_examples.full_roundtrip, /curl -N .*\/api\/stream/);
  assert.equal(initialInfo.current_meeting.meeting_id, null);
  assert.equal(initialInfo.ignored_current_meeting.meeting_id, 'demo-lark-meeting-001');
  assert.equal(initialInfo.ignored_current_meeting.reason, 'demo_sample_axis_is_not_a_real_meeting');
  assert.equal(initialInfo.annotation_route.requires_captured_at_to_create_pending, true);

  const missingCapturedAtFirst = await postJson(baseUrl, '/api/annotations', {
    id: 'device-contract-no-captured-at-first',
    kind: 'handwriting_trigger',
    label: 'missing captured_at before pending',
    text_candidates: ['missing captured_at before pending'],
  }, {
    'x-hmp-device-id': 'hanwang-alpha-001',
  });
  assert.equal(missingCapturedAtFirst.state.meeting.meeting_id, 'demo-lark-meeting-001');
  assert.equal(missingCapturedAtFirst.state.meeting.pending_binding, false);
  assert.equal(missingCapturedAtFirst.ack.created_pending_timeline, false);
  assert.equal(missingCapturedAtFirst.ack.binding_state, 'demo_ignored');
  assert.equal(missingCapturedAtFirst.ack.timing_reliable, false);
  assert.deepEqual(missingCapturedAtFirst.ack.warnings, [
    'missing_captured_at_ms',
    'unbound_time_uses_server_receive_time',
    'normalized_after_meeting_end',
  ]);

  const capturedAt = Date.parse('2026-06-26T03:00:10.000Z');
  const result = await postJson(baseUrl, '/api/annotations', {
    id: 'device-contract-ann-1',
    captured_at_ms: capturedAt,
    kind: 'handwriting_trigger',
    label: 'why?',
    text_candidates: ['why?', 'why'],
  }, {
    'x-hmp-device-id': 'hanwang-alpha-001',
  });
  assert.equal(result.state.meeting.source, 'annotation_fallback');
  assert.equal(result.state.meeting.pending_binding, true);
  assert.equal(result.state.meeting.minute_token, null);
  assert.equal(result.ack.accepted, true);
  assert.equal(result.ack.annotation_id, 'device-contract-ann-1');
  assert.equal(result.ack.operation, 'created');
  assert.equal(result.ack.idempotent, true);
  assert.equal(result.ack.idempotency_key_source, 'id');
  assert.deepEqual(result.ack.warnings, []);
  assert.equal(result.ack.timing_reliable, true);
  assert.equal(result.ack.requires_device_captured_at, false);
  assert.equal(result.ack.binding_state, 'pending_real_meeting');
  assert.equal(result.ack.pending_real_meeting, true);
  assert.equal(result.ack.created_pending_timeline, true);
  assert.equal(result.ack.replaced_existing, false);
  assert.equal(result.ack.meeting_id, result.state.meeting.meeting_id);
  assert.equal(result.ack.normalized_time_ms, 0);
  assert.equal(result.ack.time_source, 'captured_at');
  assert.equal(result.item.source, 'hanwang-alpha-001');
  assert.equal(result.item.payload.device.id, 'hanwang-alpha-001');
  assert.equal(result.item.time_source, 'captured_at');
  assert.equal(result.item.time_ms, 0);

  const pendingBindings = await (await fetch(`${baseUrl}/api/annotation-bindings`)).json();
  assert.equal(pendingBindings.binding_state, 'pending_real_meeting');
  assert.equal(pendingBindings.total, 1);
  assert.equal(pendingBindings.pending_count, 1);
  assert.equal(pendingBindings.real_axis_count, 0);
  assert.equal(pendingBindings.items[0].id, 'device-contract-ann-1');
  assert.equal(pendingBindings.items[0].binding_state, 'pending_real_meeting');

  const pendingOne = await (await fetch(`${baseUrl}/api/annotation-bindings?id=device-contract-ann-1`)).json();
  assert.equal(pendingOne.query.id, 'device-contract-ann-1');
  assert.equal(pendingOne.found, true);
  assert.equal(pendingOne.total, 1);
  assert.equal(pendingOne.item.id, 'device-contract-ann-1');
  assert.equal(pendingOne.item.binding_state, 'pending_real_meeting');

  const missingOne = await (await fetch(`${baseUrl}/api/annotation-bindings?id=missing-device-ann`)).json();
  assert.equal(missingOne.query.id, 'missing-device-ann');
  assert.equal(missingOne.found, false);
  assert.equal(missingOne.total, 0);
  assert.equal(missingOne.item, null);

  const retry = await postJson(baseUrl, '/api/annotations', {
    id: 'device-contract-ann-1',
    captured_at_ms: capturedAt + 5000,
    kind: 'handwriting_trigger',
    label: 'why? retry',
    text_candidates: ['why? retry', 'why?'],
  }, {
    'x-hmp-device-id': 'hanwang-alpha-001',
  });
  assert.equal(retry.ack.accepted, true);
  assert.equal(retry.ack.annotation_id, 'device-contract-ann-1');
  assert.equal(retry.ack.operation, 'updated');
  assert.equal(retry.ack.replaced_existing, true);
  assert.equal(retry.ack.idempotent, true);
  assert.equal(retry.ack.created_pending_timeline, false);
  assert.equal(retry.ack.binding_state, 'pending_real_meeting');
  assert.equal(retry.state.sequence.filter((item) => item.id === 'device-contract-ann-1').length, 1);
  assert.equal(retry.state.sequence.find((item) => item.id === 'device-contract-ann-1')?.label, 'why? retry');

  const generatedId = await postJson(baseUrl, '/api/annotations', {
    captured_at_ms: capturedAt + 10_000,
    kind: 'handwriting_trigger',
    label: 'no stable id',
    text_candidates: ['no stable id'],
  }, {
    'x-hmp-device-id': 'hanwang-alpha-001',
  });
  assert.equal(generatedId.ack.accepted, true);
  assert.equal(generatedId.ack.operation, 'created');
  assert.equal(generatedId.ack.idempotent, false);
  assert.equal(generatedId.ack.idempotency_key_source, null);
  assert.deepEqual(generatedId.ack.warnings, ['missing_stable_id']);

  const missingCapturedAt = await postJson(baseUrl, '/api/annotations', {
    id: 'device-contract-no-captured-at',
    kind: 'handwriting_trigger',
    label: 'missing captured_at',
    text_candidates: ['missing captured_at'],
  }, {
    'x-hmp-device-id': 'hanwang-alpha-001',
  });
  assert.equal(missingCapturedAt.ack.accepted, true);
  assert.equal(missingCapturedAt.ack.time_source, 'server_received_at');
  assert.equal(missingCapturedAt.ack.timing_reliable, false);
  assert.equal(missingCapturedAt.ack.requires_device_captured_at, true);
  assert.deepEqual(missingCapturedAt.ack.warnings, [
    'missing_captured_at_ms',
    'pending_time_uses_server_receive_time',
  ]);

  const pendingInfo = await (await fetch(`${baseUrl}/api/annotation-ingest-info`)).json();
  assert.equal(pendingInfo.annotation_route.mode, 'append_to_pending_meeting');
  assert.equal(pendingInfo.current_meeting.pending_binding, true);
  assert.equal(pendingInfo.current_meeting.source, 'annotation_fallback');

  console.log('ok device ingest contract');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
