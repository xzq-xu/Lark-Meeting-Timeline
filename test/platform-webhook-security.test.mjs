import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createHmac } from 'node:crypto';
import { once } from 'node:events';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
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

async function postRaw(baseUrl, path, rawBody, headers = {}) {
  const response = await fetch(`${baseUrl}${path}`, {
    method: 'POST',
    headers: { 'content-type': 'application/json; charset=utf-8', ...headers },
    body: rawBody,
  });
  const text = await response.text();
  let json = null;
  try {
    json = text ? JSON.parse(text) : null;
  } catch {
    json = text;
  }
  return { response, json, text };
}

function zoomSignature(secret, timestamp, rawBody) {
  return `v0=${createHmac('sha256', secret).update(`v0:${timestamp}:${rawBody}`).digest('hex')}`;
}

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'meeting-platform-webhook-security-'));
const baseUrl = `http://127.0.0.1:${port}`;
const zoomSecret = 'zoom-platform-secret';
const graphClientState = 'graph-client-state';
const googleBearer = 'google-pubsub-token';
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    TIMELINE_DATA_DIR: tempDir,
    ZOOM_WEBHOOK_SECRET_TOKEN: zoomSecret,
    MICROSOFT_GRAPH_CLIENT_STATE: graphClientState,
    GOOGLE_PUBSUB_BEARER_TOKEN: googleBearer,
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

  const graphValidation = await fetch(`${baseUrl}/api/platform-events/teams?validationToken=hello%20graph`);
  assert.equal(graphValidation.status, 200);
  assert.equal(await graphValidation.text(), 'hello graph');
  assert.match(graphValidation.headers.get('content-type') ?? '', /text\/plain/);

  const plainToken = 'zoom-url-validation-token';
  const zoomValidationRaw = JSON.stringify({
    event: 'endpoint.url_validation',
    payload: { plainToken },
  });
  const zoomValidation = await postRaw(baseUrl, '/api/platform-events/zoom', zoomValidationRaw);
  assert.equal(zoomValidation.response.status, 200);
  assert.deepEqual(zoomValidation.json, {
    plainToken,
    encryptedToken: createHmac('sha256', zoomSecret).update(plainToken).digest('hex'),
  });

  const invalidZoomRaw = JSON.stringify({
    event: 'meeting.started',
    event_ts: 1_782_442_800_000,
    payload: { object: { uuid: 'zoom-secure-001', id: 111, start_time: '2026-06-26T02:00:00.000Z' } },
  });
  const invalidZoom = await postRaw(baseUrl, '/api/platform-events/zoom', invalidZoomRaw, {
    'x-zm-request-timestamp': String(Math.floor(Date.now() / 1000)),
    'x-zm-signature': 'v0=bad',
  });
  assert.equal(invalidZoom.response.status, 401);
  assert.equal(invalidZoom.json.verification.reason, 'zoom_signature_mismatch');

  const validZoomTimestamp = String(Math.floor(Date.now() / 1000));
  const validZoom = await postRaw(baseUrl, '/api/platform-events/zoom', invalidZoomRaw, {
    'x-zm-request-timestamp': validZoomTimestamp,
    'x-zm-signature': zoomSignature(zoomSecret, validZoomTimestamp, invalidZoomRaw),
  });
  assert.equal(validZoom.response.status, 200);
  assert.equal(validZoom.json.ok, true);
  assert.equal(validZoom.json.status.last_verification.reason, 'verified');

  const invalidTeams = await postRaw(baseUrl, '/api/platform-events/teams', JSON.stringify({
    value: [{
      clientState: 'wrong-state',
      resourceData: {
        eventType: 'callStarted',
        eventDateTime: '2026-06-26T02:00:00.000Z',
        onlineMeetingId: 'teams-secure-001',
      },
    }],
  }));
  assert.equal(invalidTeams.response.status, 401);
  assert.equal(invalidTeams.json.verification.reason, 'microsoft_graph_client_state_mismatch');

  const googleWithoutBearer = await postRaw(baseUrl, '/api/platform-events/google-meet', JSON.stringify({
    id: 'google-secure-001',
    type: 'google.workspace.meet.conference.v2.started',
    time: '2026-06-26T02:00:00.000Z',
    data: { conferenceRecord: { name: 'conferenceRecords/google-secure-001' } },
  }));
  assert.equal(googleWithoutBearer.response.status, 401);
  assert.equal(googleWithoutBearer.json.verification.reason, 'google_pubsub_bearer_mismatch');

  console.log('ok platform webhook security endpoints');
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
  if (process.exitCode) {
    console.error(output);
  }
}
