import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createHmac, createSign, generateKeyPairSync } from 'node:crypto';
import { once } from 'node:events';
import { mkdtemp, rm } from 'node:fs/promises';
import { createServer } from 'node:http';
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

function base64Url(input) {
  return Buffer.from(typeof input === 'string' ? input : JSON.stringify(input), 'utf8')
    .toString('base64url');
}

function signJwt({ privateKey, kid, payload }) {
  const header = { alg: 'RS256', typ: 'JWT', kid };
  const signingInput = `${base64Url(header)}.${base64Url(payload)}`;
  const signer = createSign('RSA-SHA256');
  signer.update(signingInput);
  signer.end();
  return `${signingInput}.${signer.sign(privateKey).toString('base64url')}`;
}

const port = await freePort();
const jwksPort = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'meeting-platform-webhook-security-'));
const baseUrl = `http://127.0.0.1:${port}`;
const jwksUrl = `http://127.0.0.1:${jwksPort}/certs`;
const zoomSecret = 'zoom-platform-secret';
const graphClientState = 'graph-client-state';
const googleAudience = `${baseUrl}/api/platform-events/google-meet`;
const googleServiceAccountEmail = 'pubsub-pusher@demo-project.iam.gserviceaccount.com';
const googleKid = 'google-local-test-key';
const { publicKey: googlePublicKey, privateKey: googlePrivateKey } = generateKeyPairSync('rsa', { modulusLength: 2048 });
const googlePublicJwk = { ...googlePublicKey.export({ format: 'jwk' }), kid: googleKid, alg: 'RS256', use: 'sig' };
const jwksServer = createServer((req, res) => {
  if (req.url === '/certs') {
    res.writeHead(200, {
      'content-type': 'application/json; charset=utf-8',
      'cache-control': 'public, max-age=300',
    });
    res.end(JSON.stringify({ keys: [googlePublicJwk] }));
    return;
  }
  res.writeHead(404);
  res.end();
});
jwksServer.listen(jwksPort, '127.0.0.1');
await once(jwksServer, 'listening');
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    TIMELINE_DATA_DIR: tempDir,
    ZOOM_WEBHOOK_SECRET_TOKEN: zoomSecret,
    MICROSOFT_GRAPH_CLIENT_STATE: graphClientState,
    GOOGLE_PUBSUB_OIDC_AUDIENCE: googleAudience,
    GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL: googleServiceAccountEmail,
    GOOGLE_PUBSUB_JWKS_URL: jwksUrl,
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
    force: true,
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

  const googleEventRaw = JSON.stringify({
    force: true,
    id: 'google-secure-001',
    type: 'google.workspace.meet.conference.v2.started',
    time: '2026-06-26T02:00:00.000Z',
    data: { conferenceRecord: { name: 'conferenceRecords/google-secure-001' } },
  });
  const googleWithoutJwt = await postRaw(baseUrl, '/api/platform-events/google-meet', googleEventRaw);
  assert.equal(googleWithoutJwt.response.status, 401);
  assert.equal(googleWithoutJwt.json.verification.reason, 'google_pubsub_oidc_token_missing');

  const nowSec = Math.floor(Date.now() / 1000);
  const googleJwt = signJwt({
    privateKey: googlePrivateKey,
    kid: googleKid,
    payload: {
      iss: 'https://accounts.google.com',
      sub: '1234567890',
      aud: googleAudience,
      email: googleServiceAccountEmail,
      email_verified: true,
      iat: nowSec - 30,
      exp: nowSec + 3600,
    },
  });
  const validGoogle = await postRaw(baseUrl, '/api/platform-events/google-meet', googleEventRaw, {
    authorization: `Bearer ${googleJwt}`,
  });
  assert.equal(validGoogle.response.status, 200);
  assert.equal(validGoogle.json.ok, true);
  assert.equal(validGoogle.json.status.last_verification.reason, 'verified');
  assert.equal(validGoogle.json.status.last_verification.email, googleServiceAccountEmail);

  console.log('ok platform webhook security endpoints');
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  jwksServer.close();
  await once(jwksServer, 'close').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
  if (process.exitCode) {
    console.error(output);
  }
}
