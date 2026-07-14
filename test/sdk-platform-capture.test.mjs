import assert from 'node:assert/strict';

import {
  buildMeetingPlatformCaptureAcceptanceSummary,
  buildPlatformCaptureAcceptanceReport,
  buildPlatformCaptureSamples,
  capturePlatformWebRequest,
  capturePlatformWebhookEvent,
  parsePlatformCaptureJsonl,
  parsePlatformCaptureRecord,
  serializePlatformCaptureRecord,
} from '../packages/meeting-timeline-sdk/adapters/platform-capture.mjs';
import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import { platformCapabilityContract } from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';

const baseUrl = 'https://timeline.example.com';
const basePath = '/hooks/meeting-events';
const capturedAtMs = 1_782_442_900_000;

assert.equal(
  platformCapabilityContract('google-meet', { baseUrl }).sdk_modules.platform_capture,
  '@ai-annotation/meeting-timeline-sdk/adapters/platform-capture',
);

const googleStart = buildPlatformFixtureEvent('google-meet', 'meeting_start');
const googleEnd = buildPlatformFixtureEvent('google-meet', 'meeting_end');
const googleStartRecord = capturePlatformWebhookEvent({
  method: 'POST',
  url: `${baseUrl}${basePath}/google-meet`,
  headers: {
    authorization: 'Bearer secret',
    'content-type': 'application/json',
    'x-custom-id': 'keep-me',
  },
  body: googleStart,
}, undefined, {
  basePath,
  capturedAtMs,
  label: 'google start capture',
});

assert.equal(googleStartRecord.platform, 'google_meet');
assert.equal(googleStartRecord.route.platform, 'google_meet');
assert.equal(googleStartRecord.headers.authorization, '[redacted]');
assert.equal(googleStartRecord.headers['x-custom-id'], 'keep-me');
assert.equal(typeof googleStartRecord.raw_body_sha256, 'string');
assert.equal(googleStartRecord.raw_body, undefined);
assert.equal(googleStartRecord.body.type, 'google.workspace.meet.conference.v2.started');

const googleEndRecord = capturePlatformWebhookEvent('google-meet', googleEnd, {
  capturedAtMs: capturedAtMs + 1_000,
});
const googleSamples = buildPlatformCaptureSamples([googleStartRecord, googleEndRecord]);
assert.equal(googleSamples.google_meet.length, 2);
assert.equal(googleSamples.google_meet[0].label, 'google start capture');

const googleReport = buildPlatformCaptureAcceptanceReport('google-meet', [
  googleStartRecord,
  googleEndRecord,
], {
  baseUrl,
  env: {
    GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
  },
  requireEndEvent: true,
});
assert.equal(googleReport.accepted, true);
assert.equal(googleReport.coverage.meeting_start, true);
assert.equal(googleReport.coverage.meeting_end, true);

const jsonl = [
  serializePlatformCaptureRecord(googleStartRecord),
  serializePlatformCaptureRecord(googleEndRecord),
].join('\n');
const parsed = parsePlatformCaptureJsonl(jsonl);
assert.equal(parsed.length, 2);
assert.equal(parsed[0].id, googleStartRecord.id);
assert.throws(
  () => parsePlatformCaptureRecord(JSON.stringify({ schema: 'wrong' })),
  /Invalid platform capture record schema/,
);

const webRequestRecord = await capturePlatformWebRequest(new Request(`${baseUrl}${basePath}/zoom`, {
  method: 'POST',
  headers: {
    'content-type': 'application/json',
    'x-zm-signature': 'secret-signature',
  },
  body: JSON.stringify(buildPlatformFixtureEvent('zoom', 'meeting_start')),
}), {
  basePath,
  includeRawBody: true,
});
assert.equal(webRequestRecord.platform, 'zoom');
assert.equal(webRequestRecord.headers['x-zm-signature'], '[redacted]');
assert.match(webRequestRecord.raw_body, /meeting.started/);

const summary = buildMeetingPlatformCaptureAcceptanceSummary([
  googleStartRecord,
  googleEndRecord,
  capturePlatformWebhookEvent('zoom', buildPlatformFixtureEvent('zoom', 'meeting_start')),
], {
  baseUrl,
  env: {
    GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
    ZOOM_WEBHOOK_SECRET_TOKEN: 'zoom-secret',
  },
  requireEndEvent: true,
});
assert.equal(summary.reports.find((item) => item.platform === 'google_meet').accepted, true);
assert.equal(summary.reports.find((item) => item.platform === 'zoom').status, 'missing_required_coverage');

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
const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl,
  basePath,
  verify: false,
});
const kitRecord = kit.captureWebhook({
  method: 'POST',
  url: `${baseUrl}${basePath}/webex`,
  body: buildPlatformFixtureEvent('webex', 'meeting_start'),
});
assert.equal(kitRecord.platform, 'webex');
assert.equal(kit.capturedSamples([kitRecord]).webex.length, 1);
assert.equal(kit.capturedAcceptance('webex', [kitRecord]).coverage.meeting_start, true);

const kitWebRecord = await kit.captureFetchRequest(new Request(`${baseUrl}${basePath}/lark`, {
  method: 'POST',
  headers: { 'content-type': 'application/json' },
  body: JSON.stringify(buildPlatformFixtureEvent('lark', 'meeting_start')),
}));
assert.equal(kitWebRecord.platform, 'lark');

console.log('ok meeting platform capture samples');
