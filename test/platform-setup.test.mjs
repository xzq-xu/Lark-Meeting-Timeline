import assert from 'node:assert/strict';

import {
  GOOGLE_MEET_EVENT_TYPES,
  ZOOM_MEETING_EVENT_TYPES,
  allPlatformSetupManifests,
  buildGoogleMeetWorkspaceSubscriptionRequest,
  buildMicrosoftTeamsMeetingCallSubscriptionRequest,
  buildPlatformSetup,
  buildZoomEventSubscriptionRequest,
  platformEventEndpoint,
  platformSetupManifest,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';

const baseUrl = 'https://timeline.example.com';

assert.equal(platformEventEndpoint(baseUrl, 'google-meet'), 'https://timeline.example.com/api/platform-events/google-meet');
assert.equal(platformEventEndpoint(baseUrl, 'teams'), 'https://timeline.example.com/api/platform-events/teams');
assert.equal(platformEventEndpoint(baseUrl, 'zoom'), 'https://timeline.example.com/api/platform-events/zoom');

const googleManifest = platformSetupManifest('google_meet', { baseUrl });
assert.equal(googleManifest.endpoint, 'https://timeline.example.com/api/platform-events/google-meet');
assert.equal(googleManifest.default_event_types.includes('google.workspace.meet.conference.v2.started'), true);
assert.equal(googleManifest.required_security_env.includes('GOOGLE_PUBSUB_OIDC_AUDIENCE'), true);

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
assert.match(teamsRequest.resource, /^\/communications\/onlineMeetings\(joinWebUrl='/);
assert.equal(teamsRequest.clientState, 'client-state');
assert.equal(teamsRequest.expirationDateTime, '2026-06-28T02:00:00.000Z');

const zoomRequest = buildZoomEventSubscriptionRequest({
  webhookUrl: 'https://timeline.example.com/api/platform-events/zoom',
  accountId: 'zoom-account-1',
});
assert.equal(zoomRequest.event_webhook_url, 'https://timeline.example.com/api/platform-events/zoom');
assert.deepEqual(zoomRequest.events, ZOOM_MEETING_EVENT_TYPES);
assert.equal(zoomRequest.subscription_scope, 'account');
assert.equal(zoomRequest.account_id, 'zoom-account-1');

const zoomSetup = buildPlatformSetup('zoom', {
  baseUrl,
  zoomSubscription: {
    webhookUrl: 'https://timeline.example.com/api/platform-events/zoom',
    name: 'Timeline Zoom Events',
  },
});
assert.equal(zoomSetup.zoom_event_subscription_request.event_subscription_name, 'Timeline Zoom Events');

const all = allPlatformSetupManifests({ baseUrl });
assert.equal(all.length, 3);
assert.deepEqual(all.map((item) => item.platform), ['google_meet', 'microsoft_teams', 'zoom']);

console.log('ok platform setup builders');
