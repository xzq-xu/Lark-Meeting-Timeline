import assert from 'node:assert/strict';
import { createHmac } from 'node:crypto';

import {
  createPlatformWebhookHandler,
  handlePlatformWebhookRequest,
  verifyPlatformWebhook,
} from '../packages/meeting-timeline-sdk/adapters/platform-webhook-handler.mjs';

function zoomSignature(secret, timestamp, rawBody) {
  return `v0=${createHmac('sha256', secret).update(`v0:${timestamp}:${rawBody}`).digest('hex')}`;
}

function webexSignature(secret, rawBody) {
  return createHmac('sha1', secret).update(rawBody).digest('hex');
}

const startMs = 1_782_442_800_000;
const startIso = new Date(startMs).toISOString();
const calls = [];
const client = {
  async startMeeting(input) {
    calls.push({ method: 'startMeeting', input });
    return { ok: true, input };
  },
  async endMeeting(input) {
    calls.push({ method: 'endMeeting', input });
    return { ok: true, input };
  },
  async insertMark(input) {
    calls.push({ method: 'insertMark', input });
    return { ok: true, input };
  },
};

const handler = createPlatformWebhookHandler(client);

const teamsValidation = await handler({
  platform: 'teams',
  method: 'GET',
  url: 'https://timeline.example.com/api/platform-events/teams?validationToken=hello%20graph',
});
assert.equal(teamsValidation.status, 200);
assert.equal(teamsValidation.body, 'hello graph');
assert.match(teamsValidation.headers['content-type'], /text\/plain/);

const zoomSecret = 'zoom-handler-secret';
const plainToken = 'zoom-handler-validation-token';
const zoomValidation = await handler({
  platform: 'zoom',
  method: 'POST',
  body: {
    event: 'endpoint.url_validation',
    payload: { plainToken },
  },
}, { zoom: { secretToken: zoomSecret } });
assert.equal(zoomValidation.status, 200);
assert.deepEqual(zoomValidation.body, {
  plainToken,
  encryptedToken: createHmac('sha256', zoomSecret).update(plainToken).digest('hex'),
});

const zoomRaw = JSON.stringify({
  event: 'meeting.started',
  event_ts: startMs,
  payload: {
    object: {
      uuid: 'zoom-handler-001',
      id: 123456789,
      topic: 'SDK handler Zoom',
      join_url: 'https://zoom.us/j/123456789',
      start_time: startIso,
    },
  },
});
const zoomTimestamp = String(Math.floor(Date.now() / 1000));
const zoomStart = await handler({
  platform: 'zoom',
  method: 'POST',
  headers: {
    'x-zm-request-timestamp': zoomTimestamp,
    'x-zm-signature': zoomSignature(zoomSecret, zoomTimestamp, zoomRaw),
  },
  body: JSON.parse(zoomRaw),
  rawBody: zoomRaw,
}, { zoom: { secretToken: zoomSecret } });
assert.equal(zoomStart.status, 200);
assert.equal(zoomStart.body.ok, true);
assert.equal(zoomStart.body.verification.reason, 'verified');
assert.equal(zoomStart.body.results[0].action, 'startMeeting');
assert.equal(calls.at(-1).input.platform, 'zoom');
assert.equal(calls.at(-1).input.meeting_id, 'zoom-handler-001');

const invalidTeams = await handler({
  platform: 'microsoft-teams',
  method: 'POST',
  body: {
    value: [{
      clientState: 'wrong-state',
      resourceData: {
        eventType: 'callStarted',
        eventDateTime: startIso,
        onlineMeetingId: 'teams-handler-001',
      },
    }],
  },
}, { microsoftTeams: { clientState: 'expected-state' } });
assert.equal(invalidTeams.status, 401);
assert.equal(invalidTeams.body.verification.reason, 'microsoft_graph_client_state_mismatch');

const googleEvent = {
  id: 'google-handler-001',
  type: 'google.workspace.meet.conference.v2.started',
  time: startIso,
  data: {
    conferenceRecord: { name: 'conferenceRecords/google-handler-001' },
    meetingUri: 'https://meet.google.com/abc-defg-hij',
    title: 'SDK handler Google Meet',
  },
};
const googleWrapped = {
  message: {
    messageId: 'pubsub-message-handler-001',
    publishTime: startIso,
    data: Buffer.from(JSON.stringify(googleEvent), 'utf8').toString('base64'),
  },
  subscription: 'projects/demo/subscriptions/meet-events',
};
const googleStart = await handlePlatformWebhookRequest(client, {
  platform: 'google-meet',
  method: 'POST',
  headers: { authorization: 'Bearer google-handler-token' },
  body: googleWrapped,
}, {
  googleMeet: {
    preferBearer: true,
    bearerToken: 'google-handler-token',
  },
});
assert.equal(googleStart.status, 200);
assert.equal(googleStart.body.verification.reason, 'verified');
assert.equal(googleStart.body.results[0].action, 'startMeeting');
assert.equal(calls.at(-1).input.platform, 'google_meet');
assert.equal(calls.at(-1).input.meeting_id, 'google-handler-001');

const localDetector = await handler({
  platform: 'local-detector',
  method: 'POST',
  body: {
    type: 'active_speaker',
    detected_platform: 'google_meet',
    meeting_id: 'local-handler-001',
    occurred_at_ms: startMs + 10_000,
    speaker: { id: 'speaker-1', name: 'Ada' },
  },
}, {
  applyOptions: { speakerAsAnnotation: true },
});
assert.equal(localDetector.status, 200);
assert.equal(localDetector.body.verification.reason, 'platform_verification_not_configured');
assert.equal(localDetector.body.signals[0].type, 'speaker_started');
assert.equal(localDetector.body.results[0].action, 'insertSpeakerMark');
assert.equal(calls.at(-1).input.kind, 'speaker_started');
assert.equal(calls.at(-1).input.label, 'Ada speaking');

const webexSecret = 'webex-handler-secret';
const webexRaw = JSON.stringify({
  id: 'webex-handler-001',
  resource: 'meetings',
  event: 'started',
  data: {
    id: 'webex-handler-001',
    startTime: startIso,
  },
});
const webexVerification = await verifyPlatformWebhook('webex', {
  headers: { 'x-spark-signature': webexSignature(webexSecret, webexRaw) },
  rawBody: webexRaw,
  body: JSON.parse(webexRaw),
}, { webex: { secret: webexSecret } });
assert.equal(webexVerification.ok, true);
assert.equal(webexVerification.reason, 'verified');

const reconciledCalls = [];
const reconciledClient = {
  async startMeeting(input) {
    reconciledCalls.push({ method: 'startMeeting', input });
    return { ok: true, input };
  },
  async endMeeting(input) {
    reconciledCalls.push({ method: 'endMeeting', input });
    return { ok: true, input };
  },
  async insertMark(input) {
    reconciledCalls.push({ method: 'insertMark', input });
    return { ok: true, input };
  },
};
const reconciledHandler = createPlatformWebhookHandler(reconciledClient, {
  reconcile: true,
  verify: false,
});
const reconciledUrl = 'https://meet.google.com/xyz-abcd-efg';
const localStartBody = {
  type: 'meeting_started',
  detected_platform: 'google_meet',
  meeting_url: reconciledUrl,
  start_time_ms: startMs + 20_000,
  title: 'SDK handler reconciled Google Meet',
};
const reconciledLocalStart = await reconciledHandler({
  platform: 'local-detector',
  method: 'POST',
  body: localStartBody,
});
assert.equal(reconciledLocalStart.status, 200);
assert.equal(reconciledLocalStart.body.raw_signal_count, 1);
assert.equal(reconciledLocalStart.body.signal_count, 1);
assert.equal(reconciledLocalStart.body.results[0].action, 'startMeeting');
assert.equal(reconciledCalls.at(-1).input.platform, 'google_meet');
assert.equal(reconciledCalls.at(-1).input.meeting_url, reconciledUrl);

const duplicateLocalStart = await reconciledHandler({
  platform: 'local-detector',
  method: 'POST',
  body: localStartBody,
});
assert.equal(duplicateLocalStart.status, 200);
assert.equal(duplicateLocalStart.body.raw_signal_count, 1);
assert.equal(duplicateLocalStart.body.signal_count, 0);
assert.equal(duplicateLocalStart.body.results.length, 0);
assert.equal(duplicateLocalStart.body.reconciliation.skipped[0].reason, 'duplicate_signal_fingerprint');
assert.equal(reconciledCalls.length, 1);

const providerReconcileEvent = {
  id: 'google-handler-reconcile-001',
  type: 'google.workspace.meet.conference.v2.started',
  time: new Date(startMs + 21_000).toISOString(),
  data: {
    conferenceRecord: { name: 'conferenceRecords/google-handler-reconcile-001' },
    meetingUri: reconciledUrl,
    title: 'SDK handler reconciled Google Meet',
  },
};
const providerReconcile = await reconciledHandler({
  platform: 'google-meet',
  method: 'POST',
  body: providerReconcileEvent,
});
assert.equal(providerReconcile.status, 200);
assert.equal(providerReconcile.body.raw_signal_count, 1);
assert.equal(providerReconcile.body.signal_count, 1);
assert.equal(providerReconcile.body.reconciliation.decisions[0].reason, 'provider_reconciles_local_axis');
assert.equal(providerReconcile.body.results[0].action, 'startMeeting');
assert.equal(reconciledCalls.length, 2);
assert.equal(reconciledHandler.getReconciliationState().active_meetings.length, 1);
assert.equal(reconciledHandler.resetReconciliationState().seen.length, 0);

console.log('ok meeting timeline SDK platform webhook handler');
