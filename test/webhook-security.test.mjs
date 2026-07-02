import assert from 'node:assert/strict';
import { createHmac } from 'node:crypto';

import {
  buildZoomUrlValidationResponse,
  microsoftGraphValidationResponse,
  platformWebhookVerificationStatus,
  verifyGooglePubSubBearer,
  verifyMicrosoftGraphClientState,
  verifyZoomWebhookEvent,
} from '../packages/meeting-timeline-sdk/adapters/webhook-security.mjs';

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

const publicStatus = platformWebhookVerificationStatus('zoom', { ok: true, skipped: false, reason: 'verified' });
assert.deepEqual(publicStatus, { platform: 'zoom', ok: true, skipped: false, reason: 'verified' });

console.log('ok platform webhook security helpers');
