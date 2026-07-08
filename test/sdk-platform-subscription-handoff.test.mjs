import assert from 'node:assert/strict';

import {
  assertMeetingPlatformSubscriptionHandoff,
  assertMeetingPlatformSubscriptionHandoffMatrix,
  buildMeetingPlatformSubscriptionHandoff,
  buildMeetingPlatformSubscriptionHandoffMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-subscription-handoff.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';
const env = {
  GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
  GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL: 'pubsub@example.iam.gserviceaccount.com',
  MICROSOFT_GRAPH_CLIENT_STATE: 'client-state',
  ZOOM_WEBHOOK_SECRET_TOKEN: 'zoom-secret',
  WEBEX_WEBHOOK_SECRET: 'webex-secret',
};
const subscriptions = {
  google_meet: {
    targetResource: '//cloudidentity.googleapis.com/users/me',
    pubsubTopic: 'projects/demo/topics/meet-events',
    ttl: '86400s',
    name: 'subscriptions/google-existing',
    expireTime: '2026-06-26T12:00:00.000Z',
  },
  microsoft_teams: {
    joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/fixture',
    notificationUrl: `${baseUrl}/api/platform-events/teams`,
    clientState: 'client-state',
    now: '2026-06-26T02:00:00.000Z',
    id: 'graph-subscription-1',
    expirationDateTime: '2026-06-26T08:00:00.000Z',
  },
  zoom: {
    webhookUrl: `${baseUrl}/api/platform-events/zoom`,
    accountId: 'zoom-account-1',
  },
  webex: {
    targetUrl: `${baseUrl}/api/platform-events/webex`,
    secret: 'webex-secret',
    ownedBy: 'org',
  },
};

function assertRuntimeHandoff(handoff) {
  assert.equal(handoff.adapter_selection.readiness.selection_ready, true);
  assert.equal(handoff.adapter_selection.selection.timestamp_field, 'captured_at_ms');
  assert.equal(handoff.runtime_binding_contract.annotation_timestamp_field, 'captured_at_ms');
  assert.equal(handoff.runtime_binding_contract.provider_events_block_realtime, false);
  assert.equal(handoff.runtime_binding_contract.transcript_blocks_realtime, false);
  assert.equal(handoff.handoff_contract.adapter_selection_snapshot_required, true);
  assert.equal(handoff.handoff_contract.provider_reconcile_non_blocking, true);
  assert.equal(handoff.handoff_contract.post_meeting_artifact_non_blocking, true);
}

const google = buildMeetingPlatformSubscriptionHandoff('google-meet', {
  baseUrl,
  env,
  subscriptions,
  now: '2026-06-26T02:00:00.000Z',
});
assert.equal(google.schema, 'meeting_platform_subscription_handoff');
assert.equal(google.platform, 'google_meet');
assert.equal(google.status, 'subscription_request_ready');
assert.equal(google.ready_to_create, true);
assert.equal(google.requests[0].provider_api, 'google_workspace_events');
assert.equal(google.requests[0].path, '/v1beta/subscriptions');
assert.equal(google.requests[0].body.notificationEndpoint.pubsubTopic, 'projects/demo/topics/meet-events');
assert.equal(google.maintenance.renewal_supported, true);
assert.equal(google.realtime_annotation_policy.provider_events_block_realtime, false);
assert.equal(google.handoff_contract.annotation_timestamp_field, 'captured_at_ms');
assert.equal(google.handoff_contract.realtime_axis_source, 'local_observer_axis');
assertRuntimeHandoff(google);
assert.equal(assertMeetingPlatformSubscriptionHandoff('google-meet', {
  baseUrl,
  env,
  subscriptions,
}).ready_to_create, true);

const missingParams = buildMeetingPlatformSubscriptionHandoff('teams', {
  baseUrl,
  env,
});
assert.equal(missingParams.status, 'needs_subscription_parameters');
assert.equal(missingParams.next_actions.includes('provide_subscription_creation_parameters'), true);
assert.throws(
  () => assertMeetingPlatformSubscriptionHandoff('teams', { baseUrl, env }),
  /Meeting platform subscription handoff is not ready/,
);

const matrix = buildMeetingPlatformSubscriptionHandoffMatrix({
  baseUrl,
  env,
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  subscriptions,
  now: '2026-06-26T02:00:00.000Z',
});
assert.equal(matrix.schema, 'meeting_platform_subscription_handoff_matrix');
assert.equal(matrix.platform_count, 5);
assert.equal(matrix.ready_to_create_count, 4);
assert.equal(matrix.request_count, 1 + 1 + 1 + 7);
assert.equal(matrix.manual_setup_count, 1);
assert.equal(matrix.security_blocked_count, 0);
assert.equal(matrix.parameter_missing_count, 0);
assert.equal(matrix.rows.find((row) => row.platform === 'webex').request_count, 7);
assert.equal(matrix.rows.find((row) => row.platform === 'lark').status, 'manual_setup');
assert.equal(matrix.rows.every((row) => row.adapter_selection_ready === true), true);
assert.equal(matrix.rows.every((row) => row.annotation_timestamp_field === 'captured_at_ms'), true);
assert.equal(matrix.rows.every((row) => row.provider_events_block_realtime === false), true);
for (const handoff of matrix.handoffs) {
  assertRuntimeHandoff(handoff);
}
assert.equal(assertMeetingPlatformSubscriptionHandoffMatrix({
  baseUrl,
  env,
  platforms: ['google-meet', 'zoom', 'webex'],
  subscriptions,
}).ready_to_create_count, 3);

const blocked = buildMeetingPlatformSubscriptionHandoffMatrix({
  baseUrl,
  env: {},
  platforms: ['zoom'],
  subscriptions,
});
assert.equal(blocked.security_blocked_count, 1);
assert.equal(blocked.rows[0].missing_env.includes('ZOOM_WEBHOOK_SECRET_TOKEN'), true);

const client = {
  async startMeeting(input) { return { ok: true, input }; },
  async endMeeting(input) { return { ok: true, input }; },
  async insertMark(input) { return { ok: true, input }; },
  async insertMarks(input) { return { ok: true, input }; },
};
const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl,
  env,
  verify: false,
});
assert.equal(kit.platformSubscriptionHandoff('zoom', {
  subscriptions,
}).ready_to_create, true);
assert.equal(kit.platformSubscriptionHandoffMatrix({
  platforms: ['google-meet'],
  subscriptions,
}).ready_to_create_count, 1);
assert.equal(kit.report({
  platforms: ['google-meet'],
  subscriptions,
}).platform_subscription_handoff_matrix.platform_count, 1);

console.log('ok meeting platform subscription handoff');
