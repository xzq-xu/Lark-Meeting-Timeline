import assert from 'node:assert/strict';

import {
  MEETING_PLATFORM_PROVIDER_CONNECTION_MATRIX_SCHEMA,
  MEETING_PLATFORM_PROVIDER_CONNECTION_PACK_SCHEMA,
  buildMeetingPlatformProviderConnectionMatrix,
  buildMeetingPlatformProviderConnectionPack,
} from '../packages/meeting-timeline-sdk/adapters/platform-provider-connection.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';
const googleEnv = {
  GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
  GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL: 'pubsub@example.iam.gserviceaccount.com',
};

const googlePack = buildMeetingPlatformProviderConnectionPack('google-meet', {
  baseUrl,
  env: googleEnv,
  subscription: {
    targetResource: '//cloudidentity.googleapis.com/users/me',
    pubsubTopic: 'projects/demo/topics/meet-events',
  },
});

assert.equal(googlePack.schema, MEETING_PLATFORM_PROVIDER_CONNECTION_PACK_SCHEMA);
assert.equal(googlePack.platform, 'google_meet');
assert.equal(googlePack.provider_role, 'reconcile_and_backfill_after_local_axis');
assert.equal(googlePack.readiness.ready, true);
assert.equal(googlePack.subscription.request.notificationEndpoint.pubsubTopic, 'projects/demo/topics/meet-events');
assert.equal(googlePack.event_mapping.some((item) => item.provider_event === 'google.workspace.meet.conference.v2.started' && item.normalized_signal === 'meeting_started'), true);
assert.equal(googlePack.event_mapping.some((item) => item.timeline_role === 'post_meeting_transcript'), true);
assert.equal(googlePack.official_docs.some((doc) => doc.url.includes('developers.google.com/workspace/events')), true);
assert.equal(googlePack.security.verifier, 'verifyGooglePubSubOidcJwt');
assert.equal(googlePack.realtime_annotation_policy.provider_events_block_realtime, false);
assert.equal(googlePack.realtime_annotation_policy.annotation_timestamp_field, 'captured_at_ms');

const teamsPack = buildMeetingPlatformProviderConnectionPack('teams', {
  baseUrl,
  env: { MICROSOFT_GRAPH_CLIENT_STATE: 'client-state' },
  subscription: {
    joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/fixture',
    notificationUrl: `${baseUrl}/api/platform-events/teams`,
    clientState: 'client-state',
    now: '2026-06-26T02:00:00.000Z',
  },
});
assert.equal(teamsPack.platform, 'microsoft_teams');
assert.equal(teamsPack.readiness.ready, true);
assert.match(teamsPack.subscription.request.resource, /meetingCallEvents/);
assert.equal(teamsPack.permissions.required_permissions.includes('OnlineMeetings.Read.All or OnlineMeetings.ReadWrite.All'), true);
assert.equal(teamsPack.event_mapping.some((item) => item.provider_event === 'meetingCallEvents.updated:rosterUpdated'), true);
assert.equal(teamsPack.security.verifier, 'verifyMicrosoftGraphClientState');

const zoomPack = buildMeetingPlatformProviderConnectionPack('zoom', {
  baseUrl,
  env: {},
  subscription: {
    webhookUrl: `${baseUrl}/api/platform-events/zoom`,
    accountId: 'zoom-account-1',
  },
});
assert.equal(zoomPack.readiness.ready, false);
assert.equal(zoomPack.security.missing_env.includes('ZOOM_WEBHOOK_SECRET_TOKEN'), true);
assert.equal(zoomPack.next_actions.includes('configure_env:ZOOM_WEBHOOK_SECRET_TOKEN'), true);
assert.equal(zoomPack.subscription.request.event_webhook_url, `${baseUrl}/api/platform-events/zoom`);
assert.equal(zoomPack.event_mapping.some((item) => item.provider_event === 'recording.completed'), true);

const webexPack = buildMeetingPlatformProviderConnectionPack('webex', {
  baseUrl,
  env: { WEBEX_WEBHOOK_SECRET: 'secret' },
  subscription: {
    targetUrl: `${baseUrl}/api/platform-events/webex`,
    secret: 'secret',
    ownedBy: 'org',
  },
});
assert.equal(webexPack.readiness.ready, true);
assert.equal(Array.isArray(webexPack.subscription.request), true);
assert.equal(webexPack.subscription.request.some((item) => item.resource === 'meetings' && item.event === 'started'), true);
assert.equal(webexPack.permissions.admin_scopes.includes('meeting:admin_transcripts_read'), true);
assert.equal(webexPack.security.verifier, 'verifyWebexWebhookEvent');

const matrix = buildMeetingPlatformProviderConnectionMatrix({
  baseUrl,
  platforms: ['google-meet', 'zoom', 'webex'],
  env: {
    ...googleEnv,
    WEBEX_WEBHOOK_SECRET: 'secret',
  },
});
assert.equal(matrix.schema, MEETING_PLATFORM_PROVIDER_CONNECTION_MATRIX_SCHEMA);
assert.deepEqual(matrix.platforms, ['google_meet', 'zoom', 'webex']);
assert.equal(matrix.platform_count, 3);
assert.equal(matrix.ready_count, 2);
assert.equal(matrix.blocked_count, 1);
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').missing_env.includes('ZOOM_WEBHOOK_SECRET_TOKEN'), true);
assert.equal(matrix.next_actions.includes('capture_real_provider_events'), true);

const client = {
  async startMeeting(input) { return { ok: true, input }; },
  async endMeeting(input) { return { ok: true, input }; },
  async insertMark(input) { return { ok: true, input }; },
  async insertMarks(input) { return { ok: true, input }; },
};
const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl,
  env: googleEnv,
  verify: false,
});
assert.equal(kit.platformProviderConnectionPack('google-meet').security.verifier, 'verifyGooglePubSubOidcJwt');
assert.equal(kit.platformProviderConnectionMatrix({ platforms: ['google-meet'] }).ready_count, 1);
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_provider_connection_matrix.platform_count, 1);

console.log('ok meeting platform provider connection packs');
