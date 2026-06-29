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

async function waitForSseState(baseUrl, predicate, timeoutMs = 5_000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(`${baseUrl}/api/stream`, { signal: controller.signal });
    assert.equal(response.ok, true, 'SSE stream should open');
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const chunks = buffer.split('\n\n');
      buffer = chunks.pop() ?? '';
      for (const chunk of chunks) {
        const event = chunk.split('\n').find((line) => line.startsWith('event: '))?.slice(7);
        const data = chunk.split('\n').find((line) => line.startsWith('data: '))?.slice(6);
        if (event === 'state' && data) {
          const state = JSON.parse(data);
          if (predicate(state)) {
            await reader.cancel();
            return state;
          }
        }
      }
    }
  } finally {
    clearTimeout(timer);
  }
  throw new Error('SSE state predicate was not observed');
}

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-timeline-runtime-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'runtime-smoke-app',
    LARK_APP_SECRET: 'runtime-smoke-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://runtime-smoke.example.com/api/lark/events',
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
      'access-control-request-headers': 'content-type,x-device-id',
    },
  });
  assert.equal(preflight.status, 204);
  assert.equal(preflight.headers.get('access-control-allow-origin'), '*');

  const ingestInfo = await (await fetch(`${baseUrl}/api/annotation-ingest-info`)).json();
  assert.equal(ingestInfo.endpoint, `${baseUrl}/api/annotations`);
  assert.equal(ingestInfo.schema_url, `${baseUrl}/annotation-schema.json`);
  assert.equal(ingestInfo.cors.enabled, true);
  assert.equal(ingestInfo.timing.preferred_field, 'captured_at_ms');

  const uiStateModule = await fetch(`${baseUrl}/uiState.mjs`);
  assert.equal(uiStateModule.ok, true);
  assert.match(uiStateModule.headers.get('content-type') ?? '', /text\/javascript/);

  const startSeconds = 1_782_442_800;
  await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'runtime-direct-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(startSeconds),
    },
    event: {
      meeting: {
        id: 'runtime-direct-meeting',
        topic: 'Runtime direct meeting',
        url: 'https://vc.feishu.cn/j/runtime-direct',
        start_time: String(startSeconds),
      },
    },
  });

  const annotationPromise = waitForSseState(baseUrl, (state) => (
    state.meeting?.meeting_id === 'runtime-direct-meeting'
      && state.sequence?.some((item) => item.id === 'runtime-ann-1' && item.time_ms === 90_000)
  ));

  const annotationResult = await postJson(baseUrl, '/api/annotations', {
    id: 'runtime-ann-1',
    source: 'runtime_smoke_epaper',
    captured_at_ms: startSeconds * 1000 + 90_000,
    kind: 'handwriting_trigger',
    label: '手写 why?',
    text_candidates: ['why?', 'why'],
    intent: 'question',
    strokes: [[{ x: 0.1, y: 0.2, t: 0 }]],
  });
  assert.equal(annotationResult.item.time_ms, 90_000);
  assert.equal(annotationResult.item.source, 'runtime_smoke_epaper');
  assert.equal(annotationResult.ack.accepted, true);
  assert.equal(annotationResult.ack.operation, 'created');
  assert.equal(annotationResult.ack.idempotent, true);
  assert.equal(annotationResult.ack.binding_state, 'real_meeting_bound');
  assert.equal(annotationResult.ack.on_real_axis, true);
  assert.equal(annotationResult.ack.pending_binding, false);
  assert.equal(annotationResult.ack.meeting_id, 'runtime-direct-meeting');
  assert.equal(annotationResult.ack.normalized_time_ms, 90_000);

  const sseState = await annotationPromise;
  const sseAnnotation = sseState.sequence.find((item) => item.id === 'runtime-ann-1');
  assert.equal(sseAnnotation.time_ms, 90_000);

  const readiness = await (await fetch(`${baseUrl}/api/readiness`)).json();
  assert.equal(readiness.ready, true);
  assert.equal(readiness.checks.find((check) => check.id === 'lark_user_oauth')?.ok, false);
  assert.equal(readiness.checks.find((check) => check.id === 'passive_real_meeting_listener')?.ok, true);
  const larkReadiness = await (await fetch(`${baseUrl}/api/lark/readiness`)).json();
  assert.equal(larkReadiness.ready, readiness.ready);
  assert.equal(larkReadiness.current.meeting.meeting_id, readiness.current.meeting.meeting_id);
  const realAnnotationCheck = readiness.checks.find((check) => check.id === 'real_annotation_seen');
  assert.equal(realAnnotationCheck?.ok, true);
  assert.match(realAnnotationCheck.detail, /1 条/);

  const endResult = await postJson(baseUrl, '/api/lark/events', {
    schema: '2.0',
    header: {
      event_id: 'runtime-direct-end',
      event_type: 'vc.meeting.all_meeting_ended_v1',
    },
    event: {
      meeting: {
        id: 'runtime-direct-meeting',
        topic: 'Runtime direct meeting',
        url: 'https://vc.feishu.cn/j/runtime-direct',
        start_time: String(startSeconds),
        end_time: String(startSeconds + 300),
      },
    },
  });
  assert.equal(endResult.state.meeting.end_time, '2026-06-26T03:05:00.000Z');
  const endEvent = endResult.state.events.find((event) => event.type === 'meeting_end');
  assert.equal(endEvent.time_ms, 300_000);
  const finalAnnotation = endResult.state.sequence.find((item) => item.id === 'runtime-ann-1');
  assert.ok(finalAnnotation.time_ms < endEvent.time_ms, 'annotation should remain before meeting end');

  const cachedAfterEndResult = await postJson(baseUrl, '/api/annotations', {
    id: 'runtime-ann-cached-after-end',
    source: 'runtime_smoke_epaper',
    captured_at_ms: startSeconds * 1000 + 120_000,
    kind: 'handwriting_trigger',
    label: '会中采集会后上传',
    text_candidates: ['会中采集会后上传'],
    strokes: [],
  });
  assert.equal(cachedAfterEndResult.state.meeting.meeting_id, 'runtime-direct-meeting');
  const cachedAfterEnd = cachedAfterEndResult.state.sequence.find((item) => item.id === 'runtime-ann-cached-after-end');
  assert.equal(cachedAfterEnd.time_ms, 120_000);
  assert.equal(cachedAfterEnd.time_source, 'captured_at');
  assert.ok(cachedAfterEnd.time_ms < endEvent.time_ms, 'cached annotation should stay before meeting end');

  const diagnostics = await (await fetch(`${baseUrl}/api/lark/delivery-diagnostics`)).json();
  assert.equal(diagnostics.evidence.real_event_count >= 1, true);
  assert.equal(diagnostics.evidence.open_annotation_count, 2);
  assert.equal(diagnostics.evidence.real_meeting_axis_active, true);
  assert.equal(diagnostics.evidence.annotation_binding.binding_state, 'real_meeting_bound');
  assert.equal(diagnostics.evidence.annotation_binding.real_axis_count, 2);
  assert.equal(diagnostics.open_platform_checklist.some((item) => item.id === 'event_delivery_mode'), true);
  assert.equal(diagnostics.open_platform_checklist.some((item) => item.id === 'direct_meeting_permission'), true);
  assert.equal(diagnostics.open_platform_checklist.find((item) => item.id === 'event_delivery_evidence')?.status, 'ok');

  const annotationBindings = await (await fetch(`${baseUrl}/api/annotation-bindings`)).json();
  assert.equal(annotationBindings.binding_state, 'real_meeting_bound');
  assert.equal(annotationBindings.real_axis_count, 2);

  const acceptanceReport = await (await fetch(`${baseUrl}/api/lark/acceptance-report`)).json();
  assert.equal(acceptanceReport.ready, true);
  assert.equal(acceptanceReport.status, 'ready');
  assert.equal(acceptanceReport.current_evidence.annotation_binding.real_axis_count, 2);
  assert.equal(acceptanceReport.open_platform_checklist.some((item) => item.id === 'app_availability'), true);
  assert.equal(acceptanceReport.required_main_path.event_subscription.includes('vc.meeting.all_meeting_started_v1'), true);
  assert.equal(acceptanceReport.acceptance_steps.find((step) => step.id === 'passive_meeting_listener')?.done, true);
  assert.equal(acceptanceReport.acceptance_steps.find((step) => step.id === 'optional_probe')?.optional, true);
  assert.equal(acceptanceReport.acceptance_steps.find((step) => step.id === 'direct_start_meeting')?.done, true);
  assert.equal(acceptanceReport.acceptance_steps.find((step) => step.id === 'write_open_annotation')?.done, true);

  console.log('ok runtime direct meeting + annotation SSE');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
