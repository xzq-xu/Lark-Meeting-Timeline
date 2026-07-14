import assert from 'node:assert/strict';
import { createHmac, createSign, generateKeyPairSync } from 'node:crypto';

import {
  buildZoomUrlValidationResponse,
  microsoftGraphValidationResponse,
  platformWebhookVerificationStatus,
  verifyGooglePubSubBearer,
  verifyGooglePubSubOidcJwt,
  verifyMicrosoftGraphClientState,
  verifyWebexWebhookEvent,
  verifyZoomWebhookEvent,
} from '../packages/meeting-timeline-sdk/adapters/webhook-security.mjs';

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

const zoomSecret = 'zoom-secret-token';
const plainToken = 'zoom-plain-token';
const zoomValidation = buildZoomUrlValidationResponse({ plainToken }, { secretToken: zoomSecret });
assert.deepEqual(zoomValidation, {
  plainToken,
  encryptedToken: createHmac('sha256', zoomSecret).update(plainToken).digest('hex'),
});

const zoomBody = JSON.stringify({ event: 'meeting.started', payload: { object: { uuid: 'z-1' } } });
const zoomTimestamp = String(Math.floor(Date.now() / 1000));
const zoomSignature = `v0=${createHmac('sha256', zoomSecret).update(`v0:${zoomTimestamp}:${zoomBody}`).digest('hex')}`;
assert.equal(verifyZoomWebhookEvent({
  headers: {
    'x-zm-request-timestamp': zoomTimestamp,
    'x-zm-signature': zoomSignature,
  },
  rawBody: zoomBody,
  secretToken: zoomSecret,
}).ok, true);
assert.equal(verifyZoomWebhookEvent({
  headers: {
    'x-zm-request-timestamp': zoomTimestamp,
    'x-zm-signature': 'v0=bad',
  },
  rawBody: zoomBody,
  secretToken: zoomSecret,
}).reason, 'zoom_signature_mismatch');

const webexSecret = 'webex-shared-secret';
const webexBody = JSON.stringify({ resource: 'meetings', event: 'started', data: { id: 'webex-1' } });
const webexSignature = createHmac('sha1', webexSecret).update(webexBody).digest('hex');
assert.equal(verifyWebexWebhookEvent({
  headers: { 'x-spark-signature': webexSignature },
  rawBody: webexBody,
  secret: webexSecret,
}).ok, true);
assert.equal(verifyWebexWebhookEvent({
  headers: { 'x-spark-signature': 'bad' },
  rawBody: webexBody,
  secret: webexSecret,
}).reason, 'webex_signature_mismatch');

const graphUrl = new URL('https://timeline.example.com/api/platform-events/teams?validationToken=hello%20graph');
assert.equal(microsoftGraphValidationResponse(graphUrl), 'hello graph');
assert.equal(verifyMicrosoftGraphClientState({
  value: [
    { clientState: 'expected-state', resourceData: { id: '1' } },
    { clientState: 'expected-state', resourceData: { id: '2' } },
  ],
}, { clientState: 'expected-state' }).ok, true);
assert.equal(verifyMicrosoftGraphClientState({
  value: [{ clientState: 'wrong-state' }],
}, { clientState: 'expected-state' }).reason, 'microsoft_graph_client_state_mismatch');

assert.equal(verifyGooglePubSubBearer({
  headers: { authorization: 'Bearer expected-token' },
  expectedToken: 'expected-token',
}).ok, true);
assert.equal(verifyGooglePubSubBearer({
  headers: { authorization: 'Bearer wrong-token' },
  expectedToken: 'expected-token',
}).reason, 'google_pubsub_bearer_mismatch');

const { publicKey, privateKey } = generateKeyPairSync('rsa', { modulusLength: 2048 });
const kid = 'google-test-kid';
const publicJwk = { ...publicKey.export({ format: 'jwk' }), kid, alg: 'RS256', use: 'sig' };
const oidcAudience = 'https://timeline.example.com/api/platform-events/google-meet';
const serviceAccountEmail = 'pubsub-pusher@demo-project.iam.gserviceaccount.com';
const nowSec = Math.floor(Date.now() / 1000);
const oidcJwt = signJwt({
  privateKey,
  kid,
  payload: {
    iss: 'https://accounts.google.com',
    sub: '1234567890',
    aud: oidcAudience,
    email: serviceAccountEmail,
    email_verified: true,
    iat: nowSec - 30,
    exp: nowSec + 3600,
  },
});
const oidcResult = await verifyGooglePubSubOidcJwt({
  headers: { authorization: `Bearer ${oidcJwt}` },
  expectedAudience: oidcAudience,
  serviceAccountEmail,
  jwks: { keys: [publicJwk] },
});
assert.equal(oidcResult.ok, true);
assert.equal(oidcResult.email, serviceAccountEmail);
assert.equal(oidcResult.audience, oidcAudience);
assert.equal((await verifyGooglePubSubOidcJwt({
  token: oidcJwt,
  expectedAudience: 'https://wrong.example.com',
  serviceAccountEmail,
  jwks: { keys: [publicJwk] },
})).reason, 'google_pubsub_oidc_audience_mismatch');
assert.equal((await verifyGooglePubSubOidcJwt({
  token: oidcJwt,
  expectedAudience: oidcAudience,
  serviceAccountEmail: 'other@demo-project.iam.gserviceaccount.com',
  jwks: { keys: [publicJwk] },
})).reason, 'google_pubsub_oidc_email_mismatch');

const publicStatus = platformWebhookVerificationStatus('zoom', { ok: true, skipped: false, reason: 'verified' });
assert.deepEqual(publicStatus, { platform: 'zoom', ok: true, skipped: false, reason: 'verified' });

console.log('ok platform webhook security helpers');
