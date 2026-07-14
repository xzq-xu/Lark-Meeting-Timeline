import assert from 'node:assert/strict';

import {
  buildPlatformWebhookRouteTable,
  buildPlatformWebhookRouterStatus,
  buildPlatformWebhookRouterSetup,
  createMeetingPlatformWebhookRouter,
  matchPlatformWebhookRoute,
  platformWebhookRoutePath,
} from '../packages/meeting-timeline-sdk/adapters/platform-webhook-router.mjs';
import { platformCapabilityContract } from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';

const baseUrl = 'https://timeline.example.com';
const basePath = '/hooks/meeting-events';
const startMs = 1_782_442_800_000;
const startIso = new Date(startMs).toISOString();

assert.equal(platformWebhookRoutePath('google-meet', { basePath }), '/hooks/meeting-events/google-meet');
assert.deepEqual(matchPlatformWebhookRoute({
  method: 'POST',
  url: 'https://timeline.example.com/hooks/meeting-events/google-meet?diagnose=1',
}, { basePath }), {
  matched: true,
  kind: 'webhook',
  platform: 'google_meet',
  slug: 'google-meet',
  path: '/hooks/meeting-events/google-meet',
});
assert.equal(matchPlatformWebhookRoute({ path: '/not-platform-events/google-meet' }, { basePath }), null);
assert.equal(
  platformCapabilityContract('google-meet', { baseUrl }).sdk_modules.webhook_router,
  '@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-router',
);

const routeTable = buildPlatformWebhookRouteTable({
  baseUrl,
  basePath,
  env: {
    GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/hooks/meeting-events/google-meet`,
    ZOOM_WEBHOOK_SECRET_TOKEN: 'zoom-secret',
  },
});
assert.equal(routeTable.length, 6);
assert.equal(routeTable.find((item) => item.platform === 'microsoft_teams').methods.includes('GET'), true);
assert.equal(routeTable.find((item) => item.platform === 'google_meet').path, '/hooks/meeting-events/google-meet');
assert.equal(routeTable.find((item) => item.platform === 'zoom').missing_security_env.length, 0);

const status = buildPlatformWebhookRouterStatus({
  baseUrl,
  basePath,
  env: {
    GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/hooks/meeting-events/google-meet`,
    ZOOM_WEBHOOK_SECRET_TOKEN: 'zoom-secret',
  },
});
assert.equal(status.base_path, basePath);
assert.equal(status.platforms.find((item) => item.platform === 'google_meet').route, '/hooks/meeting-events/google-meet');
assert.equal(status.platforms.find((item) => item.platform === 'zoom').ready, true);

const setup = buildPlatformWebhookRouterSetup({ baseUrl, basePath });
assert.equal(setup.routes.length, 6);
assert.equal(setup.setup.find((item) => item.platform === 'webex').endpoint, `${baseUrl}/api/platform-events/webex`);

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

const router = createMeetingPlatformWebhookRouter(client, {
  baseUrl,
  basePath,
  verify: false,
  reconcile: true,
});

assert.equal(router.routeTable().find((item) => item.platform === 'google_meet').path, '/hooks/meeting-events/google-meet');
assert.equal(router.match({ path: '/hooks/meeting-events/status' }).kind, 'status');

const routerStatus = await router({
  method: 'GET',
  url: 'https://timeline.example.com/hooks/meeting-events/status',
});
assert.equal(routerStatus.status, 200);
assert.equal(routerStatus.body.routes.length, 6);

const routerSetup = await router({
  method: 'GET',
  url: 'https://timeline.example.com/hooks/meeting-events/google-meet/setup',
});
assert.equal(routerSetup.status, 200);
assert.equal(routerSetup.body.platform, 'google_meet');

const teamsValidation = await router({
  method: 'GET',
  url: 'https://timeline.example.com/hooks/meeting-events/teams?validationToken=hello%20router',
});
assert.equal(teamsValidation.status, 200);
assert.equal(teamsValidation.body, 'hello router');

const googleEvent = {
  id: 'google-router-001',
  type: 'google.workspace.meet.conference.v2.started',
  time: startIso,
  data: {
    conferenceRecord: { name: 'conferenceRecords/google-router-001' },
    meetingUri: 'https://meet.google.com/abc-defg-hij',
    title: 'SDK router Google Meet',
  },
};
const googleStart = await router({
  method: 'POST',
  url: 'https://timeline.example.com/hooks/meeting-events/google-meet',
  body: googleEvent,
});
assert.equal(googleStart.status, 200);
assert.equal(googleStart.body.ok, true);
assert.equal(googleStart.body.results[0].action, 'startMeeting');
assert.equal(calls.at(-1).input.platform, 'google_meet');
assert.equal(calls.at(-1).input.meeting_id, 'google-router-001');

const notFound = await router({
  method: 'POST',
  url: 'https://timeline.example.com/hooks/unknown/google-meet',
});
assert.equal(notFound.status, 404);
assert.equal(notFound.body.error, 'platform_webhook_route_not_found');

assert.equal(router.getReconciliationState().active_meetings.length, 1);
assert.equal(router.resetReconciliationState().seen.length, 0);

console.log('ok meeting platform webhook router');
