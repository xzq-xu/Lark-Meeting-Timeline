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
const tempDir = await mkdtemp(join(tmpdir(), 'lark-timeline-tenant-search-'));
const baseUrl = `http://127.0.0.1:${port}`;
const now = Date.now();
await writeFile(join(tempDir, 'state.json'), JSON.stringify({
  meeting: {
    platform: 'lark',
    meeting_id: 'tenant-bound-meeting',
    title: 'Tenant fallback meeting',
    start_time: '2026-06-26T03:00:00.000Z',
    end_time: null,
    timezone: 'Asia/Shanghai',
    pending_binding: false,
    source: 'lark_tenant_meeting_search_api',
  },
  segments: [],
  events: [{
    id: 'evt-search-tenant-bound-meeting-start',
    time_ms: 0,
    type: 'meeting_start',
    label: '应用身份会议搜索建轴',
    source: 'lark_tenant_meeting_search_api',
    metadata: { raw_type: 'vc.meeting.search', auth_mode: 'tenant_access_token' },
  }],
  sequence: [],
  duration_ms: 600_000,
  alignments: [],
  updated_at: new Date(now).toISOString(),
}, null, 2));

const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'tenant-search-test-app',
    LARK_APP_SECRET: 'tenant-search-test-secret',
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
  await postJson(baseUrl, '/api/annotations', {
    id: 'tenant-search-ann-1',
    source: 'hanwang_epaper',
    captured_at_ms: Date.parse('2026-06-26T03:01:00.000Z'),
    kind: 'handwriting_trigger',
    label: 'why?',
    text_candidates: ['why?', 'why'],
  });

  const readiness = await (await fetch(`${baseUrl}/api/readiness`)).json();
  assert.equal(readiness.ready, true);
  assert.equal(readiness.checks.find((check) => check.id === 'real_meeting_entry')?.ok, true);
  assert.equal(readiness.checks.find((check) => check.id === 'real_annotation_seen')?.ok, true);

  const report = await (await fetch(`${baseUrl}/api/lark/acceptance-report`)).json();
  assert.equal(report.ready, true);
  assert.equal(report.status, 'ready');
  assert.equal(report.current_evidence.real_meeting_axis_active, true);
  assert.equal(report.current_evidence.open_annotation_count, 1);

  console.log('ok tenant search fallback readiness');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
