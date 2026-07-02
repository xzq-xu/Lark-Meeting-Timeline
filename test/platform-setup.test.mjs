import assert from 'node:assert/strict';

import {
  GOOGLE_MEET_EVENT_TYPES,
  GOOGLE_WORKSPACE_SUBSCRIPTION_LIFECYCLE_EVENT_TYPES,
  LARK_MEETING_EVENT_TYPES,
  LOCAL_DETECTOR_EVENT_TYPES,
  MICROSOFT_GRAPH_LIFECYCLE_EVENTS,
  WEBEX_WEBHOOK_RESOURCES,
  ZOOM_MEETING_EVENT_TYPES,
  allPlatformCapabilityContracts,
  allPlatformSetupManifests,
  buildGoogleMeetWorkspaceSubscriptionRequest,
  buildGoogleWorkspaceSubscriptionRenewalRequest,
  buildMicrosoftTeamsMeetingCallSubscriptionRequest,
  buildMicrosoftGraphSubscriptionRenewalRequest,
  buildPlatformSetup,
  buildWebexWebhookRequests,
  buildZoomEventSubscriptionRequest,
  evaluateAllPlatformSetupReadiness,
  evaluateAllPlatformSubscriptionMaintenance,
  evaluatePlatformSetupReadiness,
  evaluatePlatformSubscriptionMaintenance,
  platformCapabilityContract,
  platformEventEndpoint,
  platformSetupManifest,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';

const baseUrl = 'https://timeline.example.com';

assert.equal(platformEventEndpoint(baseUrl, 'local-detector'), 'https://timeline.example.com/api/platform-events/local-detector');
assert.equal(platformEventEndpoint(baseUrl, 'lark'), 'https://timeline.example.com/api/platform-events/lark');
assert.equal(platformEventEndpoint(baseUrl, 'google-meet'), 'https://timeline.example.com/api/platform-events/google-meet');
assert.equal(platformEventEndpoint(baseUrl, 'teams'), 'https://timeline.example.com/api/platform-events/teams');
assert.equal(platformEventEndpoint(baseUrl, 'zoom'), 'https://timeline.example.com/api/platform-events/zoom');
assert.equal(platformEventEndpoint(baseUrl, 'webex'), 'https://timeline.example.com/api/platform-events/webex');

const localDetectorManifest = platformSetupManifest('desktop-observer', { baseUrl });
assert.equal(localDetectorManifest.endpoint, 'https://timeline.example.com/api/platform-events/local-detector');
assert.equal(localDetectorManifest.default_event_types.includes('meeting_started'), true);
assert.equal(LOCAL_DETECTOR_EVENT_TYPES.includes('meeting_ended'), true);
assert.equal(localDetectorManifest.capabilities.realtime_axis.status, 'supported');
assert.equal(localDetectorManifest.capabilities.sdk_modules.events, '@ai-annotation/meeting-timeline-sdk/adapters/local-detector');
const localDetectorMaintenance = evaluatePlatformSubscriptionMaintenance('local-detector', {});
assert.equal(localDetectorMaintenance.renewal_supported, false);
assert.match(localDetectorMaintenance.detail, /no_remote_subscription/);

const larkManifest = platformSetupManifest('feishu', { baseUrl });
assert.equal(larkManifest.endpoint, 'https://timeline.example.com/api/platform-events/lark');
assert.equal(larkManifest.default_event_types.includes('vc.meeting.all_meeting_started_v1'), true);
assert.equal(LARK_MEETING_EVENT_TYPES.includes('vc.meeting.all_meeting_ended_v1'), true);
assert.equal(larkManifest.capabilities.realtime_axis.status, 'supported');
assert.equal(larkManifest.capabilities.sdk_modules.events, '@ai-annotation/meeting-timeline-sdk/adapters/lark');
assert.equal(larkManifest.capabilities.sdk_modules.ingest, '@ai-annotation/meeting-timeline-sdk/adapters/platform-ingest');
const larkMaintenance = evaluatePlatformSubscriptionMaintenance('lark', {});
assert.equal(larkMaintenance.renewal_supported, false);
assert.match(larkMaintenance.detail, /does_not_use_short_cycle/);

const googleManifest = platformSetupManifest('google_meet', { baseUrl });
assert.equal(googleManifest.endpoint, 'https://timeline.example.com/api/platform-events/google-meet');
assert.equal(googleManifest.default_event_types.includes('google.workspace.meet.conference.v2.started'), true);
assert.equal(googleManifest.lifecycle_event_types.includes('google.workspace.events.subscription.v1.expirationReminder'), true);
assert.equal(googleManifest.required_security_env.includes('GOOGLE_PUBSUB_OIDC_AUDIENCE'), true);
assert.equal(GOOGLE_WORKSPACE_SUBSCRIPTION_LIFECYCLE_EVENT_TYPES.includes('google.workspace.events.subscription.v1.expired'), true);
assert.equal(googleManifest.capabilities.realtime_axis.status, 'supported_best_effort');
assert.equal(googleManifest.capabilities.post_meeting_transcript.import_endpoint, '/api/import/transcript');
assert.equal(googleManifest.capabilities.sdk_modules.events, '@ai-annotation/meeting-timeline-sdk/adapters/google-meet');

const googleCapabilities = platformCapabilityContract('google-meet', { baseUrl });
assert.equal(googleCapabilities.endpoints.platform_events, 'https://timeline.example.com/api/platform-events/google-meet');
assert.equal(googleCapabilities.realtime_transcript.status, 'not_supported');
assert.equal(googleCapabilities.limitations.includes('transcript_entries_may_differ_from_google_docs_transcript'), true);

const googleRequest = buildGoogleMeetWorkspaceSubscriptionRequest({
  targetResource: '//cloudidentity.googleapis.com/users/me',
  pubsubTopic: 'projects/demo-project/topics/meet-events',
});
assert.equal(googleRequest.targetResource, '//cloudidentity.googleapis.com/users/me');
assert.deepEqual(googleRequest.eventTypes, GOOGLE_MEET_EVENT_TYPES);
assert.equal(googleRequest.notificationEndpoint.pubsubTopic, 'projects/demo-project/topics/meet-events');
assert.equal(googleRequest.payloadOptions.includeResource, true);

const teamsRequest = buildMicrosoftTeamsMeetingCallSubscriptionRequest({
  joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/example?context={}',
  notificationUrl: 'https://timeline.example.com/api/platform-events/teams',
  clientState: 'client-state',
  now: '2026-06-26T02:00:00.000Z',
});
assert.equal(teamsRequest.changeType, 'created,updated');
assert.equal(teamsRequest.notificationUrl, 'https://timeline.example.com/api/platform-events/teams');
assert.equal(teamsRequest.lifecycleNotificationUrl, 'https://timeline.example.com/api/platform-events/teams');
assert.match(teamsRequest.resource, /^\/communications\/onlineMeetings\(joinWebUrl='/);
assert.equal(teamsRequest.clientState, 'client-state');
assert.equal(teamsRequest.expirationDateTime, '2026-06-28T02:00:00.000Z');

const teamsManifest = platformSetupManifest('teams', { baseUrl });
assert.equal(teamsManifest.lifecycle_events.includes('reauthorizationRequired'), true);
assert.equal(MICROSOFT_GRAPH_LIFECYCLE_EVENTS.includes('subscriptionRemoved'), true);
assert.equal(teamsManifest.capabilities.post_meeting_transcript.sdk_normalizer, 'normalizeMicrosoftTeamsTranscript');
assert.equal(teamsManifest.capabilities.subscription_lifecycle.status, 'supported');

const zoomRequest = buildZoomEventSubscriptionRequest({
  webhookUrl: 'https://timeline.example.com/api/platform-events/zoom',
  accountId: 'zoom-account-1',
});
assert.equal(zoomRequest.event_webhook_url, 'https://timeline.example.com/api/platform-events/zoom');
assert.deepEqual(zoomRequest.events, ZOOM_MEETING_EVENT_TYPES);
assert.equal(zoomRequest.subscription_scope, 'account');
assert.equal(zoomRequest.account_id, 'zoom-account-1');
const zoomCapabilities = platformCapabilityContract('zoom', { baseUrl });
assert.equal(zoomCapabilities.subscription_lifecycle.status, 'not_applicable');
assert.equal(zoomCapabilities.post_meeting_transcript.sdk_normalizer, 'normalizeZoomTranscript');

const zoomSetup = buildPlatformSetup('zoom', {
  baseUrl,
  zoomSubscription: {
    webhookUrl: 'https://timeline.example.com/api/platform-events/zoom',
    name: 'Timeline Zoom Events',
  },
});
assert.equal(zoomSetup.zoom_event_subscription_request.event_subscription_name, 'Timeline Zoom Events');

const webexRequests = buildWebexWebhookRequests({
  targetUrl: 'https://timeline.example.com/api/platform-events/webex',
  name: 'Timeline Webex Events',
  secret: 'webex-secret',
  ownedBy: 'org',
});
assert.equal(webexRequests.length, WEBEX_WEBHOOK_RESOURCES.reduce((sum, item) => sum + item.events.length, 0));
assert.equal(webexRequests.some((item) => item.resource === 'meetings' && item.event === 'started'), true);
assert.equal(webexRequests.some((item) => item.resource === 'meetingParticipants' && item.event === 'joined'), true);
assert.equal(webexRequests[0].targetUrl, 'https://timeline.example.com/api/platform-events/webex');
assert.equal(webexRequests[0].secret, 'webex-secret');

const webexSetup = buildPlatformSetup('webex', {
  baseUrl,
  webexSubscription: {
    targetUrl: 'https://timeline.example.com/api/platform-events/webex',
    name: 'Timeline Webex Events',
  },
});
assert.equal(webexSetup.webex_webhook_requests.length, webexRequests.length);
const webexManifest = platformSetupManifest('webex', { baseUrl });
assert.equal(webexManifest.required_scopes.includes('meeting:transcripts_read'), true);
assert.equal(webexManifest.capabilities.post_meeting_transcript.sdk_normalizer, 'normalizeWebexTranscript');
assert.equal(webexManifest.capabilities.realtime_axis.status, 'supported_best_effort');

const all = allPlatformSetupManifests({ baseUrl });
assert.equal(all.length, 6);
assert.deepEqual(all.map((item) => item.platform), ['local_detector', 'lark', 'google_meet', 'microsoft_teams', 'zoom', 'webex']);
const allCapabilities = allPlatformCapabilityContracts({ baseUrl });
assert.deepEqual(allCapabilities.map((item) => item.platform), ['local_detector', 'lark', 'google_meet', 'microsoft_teams', 'zoom', 'webex']);
assert.equal(allCapabilities.every((item) => item.endpoints.transcript_import === 'https://timeline.example.com/api/import/transcript'), true);

const googleReady = evaluatePlatformSetupReadiness('google-meet', {
  baseUrl,
  env: {
    GOOGLE_PUBSUB_OIDC_AUDIENCE: 'https://timeline.example.com/api/platform-events/google-meet',
    GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL: 'pubsub@demo.iam.gserviceaccount.com',
  },
});
assert.equal(googleReady.ready, true);
assert.equal(googleReady.checks.find((item) => item.id === 'google_pubsub_auth').detail, 'oidc_configured');

const googleBearerFallback = evaluatePlatformSetupReadiness('google-meet', {
  baseUrl,
  env: { GOOGLE_PUBSUB_BEARER_TOKEN: 'dev-token' },
});
assert.equal(googleBearerFallback.ready, true);
assert.equal(googleBearerFallback.warning_count, 1);

const zoomNotReady = evaluatePlatformSetupReadiness('zoom', {
  baseUrl: 'http://timeline.example.com',
  env: {},
});
assert.equal(zoomNotReady.ready, false);
assert.equal(zoomNotReady.blocking_count, 2);
assert.equal(zoomNotReady.checks.some((item) => item.id === 'endpoint' && item.detail === 'endpoint_must_be_https_or_localhost'), true);

const allReadiness = evaluateAllPlatformSetupReadiness({
  baseUrl,
  env: {
    GOOGLE_PUBSUB_OIDC_AUDIENCE: 'https://timeline.example.com/api/platform-events/google-meet',
    MICROSOFT_GRAPH_CLIENT_STATE: 'teams-state',
    ZOOM_WEBHOOK_SECRET_TOKEN: 'zoom-secret',
    WEBEX_WEBHOOK_SECRET: 'webex-secret',
  },
});
assert.deepEqual(allReadiness.map((item) => item.ready), [true, true, true, true, true, true]);

const graphRenewal = buildMicrosoftGraphSubscriptionRenewalRequest({
  subscriptionId: 'graph-sub-1',
  now: '2026-06-26T02:00:00.000Z',
  ttlSeconds: 172800,
});
assert.deepEqual(graphRenewal, {
  method: 'PATCH',
  path: '/subscriptions/graph-sub-1',
  body: { expirationDateTime: '2026-06-28T02:00:00.000Z' },
});

const googleRenewal = buildGoogleWorkspaceSubscriptionRenewalRequest({
  subscriptionName: 'subscriptions/google-sub-1',
  ttl: '86400s',
});
assert.deepEqual(googleRenewal, {
  method: 'PATCH',
  path: '/v1beta/subscriptions/google-sub-1',
  query: { updateMask: 'ttl' },
  body: { ttl: '86400s' },
});

const teamsMaintenance = evaluatePlatformSubscriptionMaintenance('teams', {
  id: 'graph-sub-1',
  expirationDateTime: '2026-06-26T06:00:00.000Z',
}, {
  now: '2026-06-26T02:00:00.000Z',
  renewalWindowMs: 12 * 60 * 60 * 1000,
});
assert.equal(teamsMaintenance.status, 'renewal_due');
assert.equal(teamsMaintenance.renewal_due, true);
assert.equal(teamsMaintenance.renewal_request.path, '/subscriptions/graph-sub-1');

const googleMaintenance = evaluatePlatformSubscriptionMaintenance('google-meet', {
  name: 'subscriptions/google-sub-1',
  expireTime: '2026-06-28T02:00:00.000Z',
}, {
  now: '2026-06-26T02:00:00.000Z',
});
assert.equal(googleMaintenance.status, 'active');
assert.equal(googleMaintenance.renewal_due, false);
assert.equal(googleMaintenance.renewal_request.body.ttl, '86400s');

const zoomMaintenance = evaluatePlatformSubscriptionMaintenance('zoom', {}, {
  now: '2026-06-26T02:00:00.000Z',
});
assert.equal(zoomMaintenance.renewal_supported, false);
assert.equal(zoomMaintenance.renewal_due, false);
const webexMaintenance = evaluatePlatformSubscriptionMaintenance('webex', {}, {
  now: '2026-06-26T02:00:00.000Z',
});
assert.equal(webexMaintenance.renewal_supported, false);
assert.equal(webexMaintenance.renewal_due, false);

const allMaintenance = evaluateAllPlatformSubscriptionMaintenance({
  google_meet: { name: 'subscriptions/google-sub-1', expireTime: '2026-06-26T03:00:00.000Z' },
  microsoft_teams: { id: 'graph-sub-1', expirationDateTime: '2026-06-29T02:00:00.000Z' },
}, {
  now: '2026-06-26T02:00:00.000Z',
});
assert.deepEqual(allMaintenance.map((item) => item.platform), ['local_detector', 'lark', 'google_meet', 'microsoft_teams', 'zoom', 'webex']);
assert.equal(allMaintenance[0].renewal_supported, false);
assert.equal(allMaintenance[1].renewal_supported, false);
assert.equal(allMaintenance[2].renewal_due, true);
assert.equal(allMaintenance[3].status, 'active');

console.log('ok platform setup builders');
