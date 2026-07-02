import assert from 'node:assert/strict';
import { createHmac } from 'node:crypto';

import {
  createMeetingPlatformFetchHandler,
  platformWebhookRequestFromWebRequest,
  platformWebhookResponseToWebResponse,
} from '../packages/meeting-timeline-sdk/adapters/platform-http.mjs';
import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import { platformCapabilityContract } from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';

function zoomSignature(secret, timestamp, rawBody) {
  return `v0=${createHmac('sha256', secret).update(`v0:${timestamp}:${rawBody}`).digest('hex')}`;
}

const baseUrl = 'https://timeline.example.com';
const basePath = '/hooks/meeting-events';
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

assert.equal(
  platformCapabilityContract('google-meet', { baseUrl }).sdk_modules.platform_http,
  '@ai-annotation/meeting-timeline-sdk/adapters/platform-http',
);

const googleBody = buildPlatformFixtureEvent('google-meet', 'meeting_start');
const googleRequest = new Request(`${baseUrl}${basePath}/google-meet`, {
  method: 'POST',
  headers: { 'content-type': 'application/json' },
  body: JSON.stringify(googleBody),
});
const parsedGoogleRequest = await platformWebhookRequestFromWebRequest(googleRequest);
assert.equal(parsedGoogleRequest.method, 'POST');
assert.equal(parsedGoogleRequest.rawBody, JSON.stringify(googleBody));
assert.equal(parsedGoogleRequest.body.type, 'google.workspace.meet.conference.v2.started');

const fetchHandler = createMeetingPlatformFetchHandler(client, {
  baseUrl,
  basePath,
  verify: false,
  reconcile: true,
});

const googleResponse = await fetchHandler(new Request(`${baseUrl}${basePath}/google-meet`, {
  method: 'POST',
  headers: { 'content-type': 'application/json' },
  body: JSON.stringify(googleBody),
}));
assert.equal(googleResponse.status, 200);
assert.match(googleResponse.headers.get('content-type'), /application\/json/);
const googleJson = await googleResponse.json();
assert.equal(googleJson.ok, true);
assert.equal(googleJson.results[0].action, 'startMeeting');
assert.equal(calls.at(-1).input.platform, 'google_meet');
assert.equal(calls.at(-1).input.meeting_id, 'google-fixture-001');

const statusResponse = await fetchHandler(new Request(`${baseUrl}${basePath}/status`, { method: 'GET' }));
assert.equal(statusResponse.status, 200);
const statusJson = await statusResponse.json();
assert.equal(statusJson.base_path, basePath);
assert.equal(statusJson.routes.length, 6);

const teamsValidation = await fetchHandler(new Request(`${baseUrl}${basePath}/teams?validationToken=hello%20graph`, {
  method: 'GET',
}));
assert.equal(teamsValidation.status, 200);
assert.match(teamsValidation.headers.get('content-type'), /text\/plain/);
assert.equal(await teamsValidation.text(), 'hello graph');

const zoomSecret = 'zoom-http-secret';
const zoomRaw = JSON.stringify(buildPlatformFixtureEvent('zoom', 'meeting_start'));
const zoomTimestamp = String(Math.floor(Date.now() / 1000));
const zoomHandler = createMeetingPlatformFetchHandler(client, {
  baseUrl,
  basePath,
  reconcile: false,
  zoom: { secretToken: zoomSecret },
});
const zoomResponse = await zoomHandler(new Request(`${baseUrl}${basePath}/zoom`, {
  method: 'POST',
  headers: {
    'content-type': 'application/json',
    'x-zm-request-timestamp': zoomTimestamp,
    'x-zm-signature': zoomSignature(zoomSecret, zoomTimestamp, zoomRaw),
  },
  body: zoomRaw,
}));
assert.equal(zoomResponse.status, 200);
const zoomJson = await zoomResponse.json();
assert.equal(zoomJson.verification.reason, 'verified');
assert.equal(zoomJson.results[0].action, 'startMeeting');
assert.equal(calls.at(-1).input.platform, 'zoom');

const invalidJson = await fetchHandler(new Request(`${baseUrl}${basePath}/google-meet`, {
  method: 'POST',
  headers: { 'content-type': 'application/json' },
  body: '{bad-json',
}));
assert.equal(invalidJson.status, 400);
assert.equal((await invalidJson.json()).details.reason, 'invalid_json_body');

const textResponse = platformWebhookResponseToWebResponse({
  status: 202,
  headers: { 'content-type': 'text/plain; charset=utf-8' },
  body: 'accepted',
});
assert.equal(textResponse.status, 202);
assert.equal(await textResponse.text(), 'accepted');

const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl,
  basePath,
  verify: false,
});
const kitResponse = await kit.handleFetchRequest(new Request(`${baseUrl}${basePath}/webex`, {
  method: 'POST',
  headers: { 'content-type': 'application/json' },
  body: JSON.stringify(buildPlatformFixtureEvent('webex', 'meeting_start')),
}));
assert.equal(kitResponse.status, 200);
assert.equal((await kitResponse.json()).results[0].action, 'startMeeting');
assert.equal(calls.at(-1).input.platform, 'webex');

console.log('ok meeting platform HTTP adapter');
