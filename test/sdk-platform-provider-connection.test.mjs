import assert from 'node:assert/strict';

import {
  MEETING_PLATFORM_PROVIDER_CONNECTION_MATRIX_SCHEMA,
  MEETING_PLATFORM_PROVIDER_CONNECTION_PACK_SCHEMA,
  buildMeetingPlatformProviderConnectionMatrix,
  buildMeetingPlatformProviderConnectionPack,
} from '../packages/meeting-timeline-sdk/adapters/platform-provider-connection.mjs';
import { buildMeetingPlatformAdapterSelection } from '../packages/meeting-timeline-sdk/adapters/platform-adapter-selection.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';
const googleEnv = {
  GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
  GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL: 'pubsub@example.iam.gserviceaccount.com',
};

function assertNonBlockingLocalAxis(pack) {
  assert.equal(pack.adapter_selection.readiness.selection_ready, true);
  assert.equal(pack.adapter_selection.selection.timestamp_field, 'captured_at_ms');
  assert.equal(pack.runtime_binding_contract.annotation_timestamp_field, 'captured_at_ms');
  assert.equal(pack.runtime_binding_contract.provider_events_block_realtime, false);
  assert.equal(pack.runtime_binding_contract.transcript_blocks_realtime, false);
  assert.equal(pack.runtime_binding_contract.must_read_adapter_selection_before_session, true);
  assert.equal(pack.runtime_binding_contract.per_meeting_annotation_isolation_required, true);
  assert.equal(pack.realtime_annotation_policy.provider_events_block_realtime, false);
}

function assertMatchesCanonicalSelection(platform, pack) {
  const selection = buildMeetingPlatformAdapterSelection(platform, {}, { baseUrl });
  assert.equal(pack.adapter_selection.selection.axis_source, selection.selection.axis_source);
  assert.equal(pack.adapter_selection.selection.axis_surface, selection.selection.axis_surface);
  assert.equal(pack.adapter_selection.selection.speaker_track_source, selection.selection.speaker_track_source);
  assert.equal(pack.adapter_selection.selection.post_meeting_artifact_source, selection.selection.post_meeting_artifact_source);
  assert.equal(pack.runtime_binding_contract.realtime_axis_surface, selection.selection.axis_surface);
}

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
assert.equal(googlePack.adapter_selection.selection.axis_source, 'local_observer_axis');
assert.equal(googlePack.adapter_selection.selection.axis_surface, 'browser_extension');
assert.equal(googlePack.runtime_binding_contract.provider_reconcile_source, 'provider_reconcile');
assertNonBlockingLocalAxis(googlePack);
assertMatchesCanonicalSelection('google-meet', googlePack);

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
assert.equal(teamsPack.runtime_binding_contract.realtime_axis_source, 'local_observer_axis');
assert.equal(teamsPack.runtime_binding_contract.realtime_axis_surface, 'desktop_or_browser_observer');
assertNonBlockingLocalAxis(teamsPack);
assertMatchesCanonicalSelection('teams', teamsPack);

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
assert.equal(zoomPack.adapter_selection.selection.provider_reconcile_required_for_production, true);
assert.equal(zoomPack.runtime_binding_contract.realtime_axis_surface, 'native_detector');
assertNonBlockingLocalAxis(zoomPack);
assertMatchesCanonicalSelection('zoom', zoomPack);

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
assert.equal(webexPack.runtime_binding_contract.realtime_axis_surface, 'browser_extension_or_native_detector');
assertNonBlockingLocalAxis(webexPack);
assertMatchesCanonicalSelection('webex', webexPack);

const larkPack = buildMeetingPlatformProviderConnectionPack('lark', {
  baseUrl,
  env: {},
});
assert.equal(larkPack.platform, 'lark');
assert.equal(larkPack.readiness.ready, true);
assert.equal(larkPack.transport, 'Feishu/Lark long connection or HTTP event callback');
assert.equal(larkPack.event_mapping.some((item) => item.provider_event === 'vc.meeting.all_meeting_started_v1'), true);
assert.equal(larkPack.runtime_binding_contract.provider_role, 'reconcile_and_backfill_after_local_axis');
assert.equal(larkPack.runtime_binding_contract.realtime_axis_surface, 'browser_extension_or_desktop_observer');
assertNonBlockingLocalAxis(larkPack);
assertMatchesCanonicalSelection('lark', larkPack);

const matrix = buildMeetingPlatformProviderConnectionMatrix({
  baseUrl,
  platforms: ['google-meet', 'zoom', 'webex', 'lark'],
  env: {
    ...googleEnv,
    WEBEX_WEBHOOK_SECRET: 'secret',
  },
});
assert.equal(matrix.schema, MEETING_PLATFORM_PROVIDER_CONNECTION_MATRIX_SCHEMA);
assert.deepEqual(matrix.platforms, ['google_meet', 'zoom', 'webex', 'lark']);
assert.equal(matrix.platform_count, 4);
assert.equal(matrix.ready_count, 3);
assert.equal(matrix.blocked_count, 1);
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').missing_env.includes('ZOOM_WEBHOOK_SECRET_TOKEN'), true);
assert.equal(matrix.rows.every((row) => row.adapter_selection_ready === true), true);
assert.equal(matrix.rows.every((row) => row.provider_events_block_realtime === false), true);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').realtime_axis_surface, 'browser_extension');
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').realtime_axis_surface, 'native_detector');
assert.equal(matrix.rows.find((row) => row.platform === 'lark').realtime_axis_surface, 'browser_extension_or_desktop_observer');
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
