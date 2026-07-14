import assert from 'node:assert/strict';
import { createHmac } from 'node:crypto';
import { Readable } from 'node:stream';

import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import {
  createMeetingPlatformExpressMiddleware,
  createMeetingPlatformNodeHandler,
  platformWebhookRequestFromNodeRequest,
  writePlatformWebhookNodeResponse,
} from '../packages/meeting-timeline-sdk/adapters/platform-node.mjs';
import { platformCapabilityContract } from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';

function requestFromBody(body, {
  method = 'POST',
  url = '/',
  headers = { 'content-type': 'application/json' },
} = {}) {
  const req = Readable.from(body ? [body] : []);
  req.method = method;
  req.url = url;
  req.headers = headers;
  return req;
}

function responseMock() {
  const res = {
    statusCode: 200,
    headers: {},
    body: '',
    setHeader(key, value) {
      this.headers[String(key).toLowerCase()] = value;
    },
    end(body = '') {
      this.body = String(body);
      this.ended = true;
    },
  };
  return res;
}

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
  platformCapabilityContract('google-meet', { baseUrl }).sdk_modules.platform_node,
  '@ai-annotation/meeting-timeline-sdk/adapters/platform-node',
);

const googleRaw = JSON.stringify(buildPlatformFixtureEvent('google-meet', 'meeting_start'));
const parsed = await platformWebhookRequestFromNodeRequest(requestFromBody(googleRaw, {
  url: `${basePath}/google-meet`,
}), { basePath });
assert.equal(parsed.method, 'POST');
assert.equal(parsed.rawBody, googleRaw);
assert.equal(parsed.body.type, 'google.workspace.meet.conference.v2.started');

const nodeHandler = createMeetingPlatformNodeHandler(client, {
  baseUrl,
  basePath,
  verify: false,
  reconcile: true,
});
const googleRes = responseMock();
await nodeHandler(requestFromBody(googleRaw, { url: `${basePath}/google-meet` }), googleRes);
assert.equal(googleRes.statusCode, 200);
assert.match(googleRes.headers['content-type'], /application\/json/);
const googleJson = JSON.parse(googleRes.body);
assert.equal(googleJson.ok, true);
assert.equal(googleJson.results[0].action, 'startMeeting');
assert.equal(calls.at(-1).input.platform, 'google_meet');
assert.equal(calls.at(-1).input.meeting_id, 'google-fixture-001');

const teamsRes = responseMock();
await nodeHandler(requestFromBody('', {
  method: 'GET',
  url: `${basePath}/teams?validationToken=hello%20teams`,
}), teamsRes);
assert.equal(teamsRes.statusCode, 200);
assert.match(teamsRes.headers['content-type'], /text\/plain/);
assert.equal(teamsRes.body, 'hello teams');

const zoomSecret = 'zoom-node-secret';
const zoomRaw = JSON.stringify(buildPlatformFixtureEvent('zoom', 'meeting_start'));
const zoomTimestamp = String(Math.floor(Date.now() / 1000));
const zoomRes = responseMock();
await createMeetingPlatformNodeHandler(client, {
  baseUrl,
  basePath,
  zoom: { secretToken: zoomSecret },
})(requestFromBody(zoomRaw, {
  url: `${basePath}/zoom`,
  headers: {
    'content-type': 'application/json',
    'x-zm-request-timestamp': zoomTimestamp,
    'x-zm-signature': zoomSignature(zoomSecret, zoomTimestamp, zoomRaw),
  },
}), zoomRes);
assert.equal(zoomRes.statusCode, 200);
const zoomJson = JSON.parse(zoomRes.body);
assert.equal(zoomJson.verification.reason, 'verified');
assert.equal(zoomJson.results[0].action, 'startMeeting');
assert.equal(calls.at(-1).input.platform, 'zoom');

const invalidJsonRes = responseMock();
await nodeHandler(requestFromBody('{bad-json', {
  url: `${basePath}/google-meet`,
  headers: { 'content-type': 'application/json' },
}), invalidJsonRes);
assert.equal(invalidJsonRes.statusCode, 400);
assert.equal(JSON.parse(invalidJsonRes.body).details.reason, 'invalid_json_body');

const manualRes = responseMock();
writePlatformWebhookNodeResponse(manualRes, {
  status: 202,
  headers: { 'content-type': 'text/plain; charset=utf-8' },
  body: 'accepted',
});
assert.equal(manualRes.statusCode, 202);
assert.equal(manualRes.body, 'accepted');

const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl,
  basePath,
  verify: false,
});
const webexRes = responseMock();
await kit.handleNodeRequest(requestFromBody(JSON.stringify(buildPlatformFixtureEvent('webex', 'meeting_start')), {
  url: `${basePath}/webex`,
}), webexRes);
assert.equal(webexRes.statusCode, 200);
assert.equal(JSON.parse(webexRes.body).results[0].action, 'startMeeting');
assert.equal(calls.at(-1).input.platform, 'webex');

const expressRes = responseMock();
const expressMiddleware = createMeetingPlatformExpressMiddleware(client, {
  baseUrl,
  basePath,
  verify: false,
});
await expressMiddleware(requestFromBody(JSON.stringify(buildPlatformFixtureEvent('lark', 'meeting_start')), {
  url: `${basePath}/lark`,
}), expressRes);
assert.equal(expressRes.statusCode, 200);
assert.equal(JSON.parse(expressRes.body).results[0].action, 'startMeeting');
assert.equal(calls.at(-1).input.platform, 'lark');

console.log('ok meeting platform Node adapter');
