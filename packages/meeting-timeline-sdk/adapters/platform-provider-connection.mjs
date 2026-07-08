import { compactObject } from './internal-utils.mjs';
import {
  MEETING_PLATFORM_KEYS,
  buildPlatformIntegrationPlan,
  buildPlatformPermissionPlan,
  buildPlatformSetup,
  normalizeMeetingPlatform,
  platformSetupManifest,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_PROVIDER_CONNECTION_PACK_SCHEMA = 'meeting_platform_provider_connection_pack';
export const MEETING_PLATFORM_PROVIDER_CONNECTION_MATRIX_SCHEMA = 'meeting_platform_provider_connection_matrix';
export const MEETING_PLATFORM_PROVIDER_CONNECTION_SCHEMA_VERSION = 1;

const OFFICIAL_DOCS = Object.freeze({
  local_detector: Object.freeze([
    {
      label: 'Meeting Timeline SDK local detector',
      url: 'https://github.com/xzq-xu/Lark-Meeting-Timeline',
      note: 'Host-controlled detector path; no remote provider subscription is required.',
    },
  ]),
  lark: Object.freeze([
    {
      label: 'Feishu/Lark Open Platform event subscriptions',
      url: 'https://open.feishu.cn/document/server-docs/event-subscription-guide/overview',
      note: 'Use long-connection or HTTP callbacks depending on deployment mode.',
    },
  ]),
  google_meet: Object.freeze([
    {
      label: 'Subscribe to Google Meet events',
      url: 'https://developers.google.com/workspace/events/guides/events-meet',
      note: 'Conference lifecycle, participant, recording, transcript, and smart note events.',
    },
    {
      label: 'Choose Google Workspace Events API scopes',
      url: 'https://developers.google.com/workspace/events/guides/auth',
      note: 'Workspace Events uses product scopes such as Meet and Drive scopes.',
    },
    {
      label: 'Google Workspace Events API release notes',
      url: 'https://developers.google.com/workspace/events/release-notes',
      note: 'Use the generally available v1 endpoint for Meet subscriptions.',
    },
  ]),
  microsoft_teams: Object.freeze([
    {
      label: 'Get change notifications for Microsoft Teams meeting call updates',
      url: 'https://learn.microsoft.com/en-us/graph/changenotifications-for-onlinemeeting',
      note: 'Call started, call ended, and roster updated change notifications.',
    },
    {
      label: 'Set up Microsoft Graph change notifications with resource data',
      url: 'https://learn.microsoft.com/en-us/graph/change-notifications-with-resource-data',
      note: 'Rich notifications can reduce follow-up Graph reads.',
    },
    {
      label: 'Teams transcript and recording change notifications',
      url: 'https://learn.microsoft.com/en-us/graph/teams-changenotifications-callrecording-and-calltranscript',
      note: 'Post-meeting transcript and recording availability notifications.',
    },
  ]),
  zoom: Object.freeze([
    {
      label: 'Zoom Meeting webhooks',
      url: 'https://developers.zoom.us/docs/api/meetings/events/',
      note: 'Meeting started/ended, participant, and recording events.',
    },
    {
      label: 'Using Zoom webhooks',
      url: 'https://developers.zoom.us/docs/api/webhooks/',
      note: 'Endpoint validation, delivery, and troubleshooting guidance.',
    },
    {
      label: 'Zoom webhook validation announcement',
      url: 'https://developers.zoom.us/docs/platform/announcements/#webhook-validation',
      note: 'Zoom validates webhook endpoints before event delivery.',
    },
  ]),
  webex: Object.freeze([
    {
      label: 'Create a Webex webhook',
      url: 'https://developer.webex.com/messaging/docs/api/v1/webhooks/create-a-webhook',
      note: 'Meetings, participants, recordings, and transcript resource webhooks.',
    },
    {
      label: 'Webex webhook guide',
      url: 'https://developer.webex.com/messaging/docs/api/guides/webhooks',
      note: 'Meeting started/ended and participant events require individual webhooks.',
    },
  ]),
});

const PROVIDER_EVENT_MAPPING = Object.freeze({
  local_detector: Object.freeze([
    { provider_event: 'meeting_started', normalized_signal: 'meeting_started', timeline_role: 'primary_axis_start' },
    { provider_event: 'meeting_ended', normalized_signal: 'meeting_ended', timeline_role: 'primary_axis_end' },
    { provider_event: 'speaker_started', normalized_signal: 'speaker_started', timeline_role: 'speaker_marker' },
  ]),
  lark: Object.freeze([
    { provider_event: 'vc.meeting.all_meeting_started_v1', normalized_signal: 'meeting_started', timeline_role: 'axis_start_reconcile' },
    { provider_event: 'vc.meeting.all_meeting_ended_v1', normalized_signal: 'meeting_ended', timeline_role: 'axis_end_reconcile' },
    { provider_event: 'vc.meeting.join_meeting_v1', normalized_signal: 'participant_joined', timeline_role: 'participant_marker' },
    { provider_event: 'vc.meeting.leave_meeting_v1', normalized_signal: 'participant_left', timeline_role: 'participant_marker' },
  ]),
  google_meet: Object.freeze([
    { provider_event: 'google.workspace.meet.conference.v2.started', normalized_signal: 'meeting_started', timeline_role: 'axis_start_reconcile' },
    { provider_event: 'google.workspace.meet.conference.v2.ended', normalized_signal: 'meeting_ended', timeline_role: 'axis_end_reconcile' },
    { provider_event: 'google.workspace.meet.participant.v2.joined', normalized_signal: 'participant_joined', timeline_role: 'participant_marker' },
    { provider_event: 'google.workspace.meet.participant.v2.left', normalized_signal: 'participant_left', timeline_role: 'participant_marker' },
    { provider_event: 'google.workspace.meet.transcript.v2.fileGenerated', normalized_signal: 'artifact_ready', timeline_role: 'post_meeting_transcript' },
    { provider_event: 'google.workspace.events.subscription.v1.expirationReminder', normalized_signal: 'subscription_lifecycle', timeline_role: 'maintenance' },
  ]),
  microsoft_teams: Object.freeze([
    { provider_event: 'meetingCallEvents.created', normalized_signal: 'meeting_started', timeline_role: 'axis_start_reconcile' },
    { provider_event: 'meetingCallEvents.updated:callEnded', normalized_signal: 'meeting_ended', timeline_role: 'axis_end_reconcile' },
    { provider_event: 'meetingCallEvents.updated:rosterUpdated', normalized_signal: 'participant_joined_or_left', timeline_role: 'participant_marker' },
    { provider_event: 'callTranscript.created', normalized_signal: 'artifact_ready', timeline_role: 'post_meeting_transcript' },
    { provider_event: 'reauthorizationRequired', normalized_signal: 'subscription_lifecycle', timeline_role: 'maintenance' },
  ]),
  zoom: Object.freeze([
    { provider_event: 'meeting.started', normalized_signal: 'meeting_started', timeline_role: 'axis_start_reconcile' },
    { provider_event: 'meeting.ended', normalized_signal: 'meeting_ended', timeline_role: 'axis_end_reconcile' },
    { provider_event: 'meeting.participant_joined', normalized_signal: 'participant_joined', timeline_role: 'participant_marker' },
    { provider_event: 'meeting.participant_left', normalized_signal: 'participant_left', timeline_role: 'participant_marker' },
    { provider_event: 'recording.completed', normalized_signal: 'artifact_ready', timeline_role: 'post_meeting_transcript_or_recording' },
  ]),
  webex: Object.freeze([
    { provider_event: 'meetings.started', normalized_signal: 'meeting_started', timeline_role: 'axis_start_reconcile' },
    { provider_event: 'meetings.ended', normalized_signal: 'meeting_ended', timeline_role: 'axis_end_reconcile' },
    { provider_event: 'meetingParticipants.joined', normalized_signal: 'participant_joined', timeline_role: 'participant_marker' },
    { provider_event: 'meetingParticipants.left', normalized_signal: 'participant_left', timeline_role: 'participant_marker' },
    { provider_event: 'meetingTranscripts.created', normalized_signal: 'artifact_ready', timeline_role: 'post_meeting_transcript' },
    { provider_event: 'recordings.created', normalized_signal: 'artifact_ready', timeline_role: 'post_meeting_recording' },
  ]),
});

const SECURITY_STRATEGY = Object.freeze({
  local_detector: Object.freeze({
    verifier: 'trusted_host_context',
    module: '@ai-annotation/meeting-timeline-sdk/adapters/local-detector',
    notes: ['Only accept local detector calls from a trusted app, extension, or authenticated host channel.'],
  }),
  lark: Object.freeze({
    verifier: 'lark_event_callback_or_long_connection_auth',
    module: '@ai-annotation/meeting-timeline-sdk/adapters/lark',
    notes: ['Use Lark app credentials and optional verification token/encryption key for HTTP callbacks.'],
  }),
  google_meet: Object.freeze({
    verifier: 'verifyGooglePubSubOidcJwt',
    module: '@ai-annotation/meeting-timeline-sdk/adapters/webhook-security',
    notes: ['Prefer authenticated Pub/Sub push OIDC. Bearer fallback is only a development fallback.'],
  }),
  microsoft_teams: Object.freeze({
    verifier: 'verifyMicrosoftGraphClientState',
    module: '@ai-annotation/meeting-timeline-sdk/adapters/webhook-security',
    notes: ['Validate clientState and handle lifecycle notifications before subscription expiry.'],
  }),
  zoom: Object.freeze({
    verifier: 'verifyZoomWebhookEvent',
    module: '@ai-annotation/meeting-timeline-sdk/adapters/webhook-security',
    notes: ['Respond to endpoint.url_validation and verify x-zm-signature on every request.'],
  }),
  webex: Object.freeze({
    verifier: 'verifyWebexWebhookEvent',
    module: '@ai-annotation/meeting-timeline-sdk/adapters/webhook-security',
    notes: ['Verify X-Spark-Signature with the configured shared secret.'],
  }),
});

const ADAPTER_SELECTION_SNAPSHOT_DEFAULTS = Object.freeze({
  local_detector: Object.freeze({
    display_name: 'Local Meeting Detector',
    recommended_mode: 'trusted_host_detector_only',
    axis_surface: 'host_sdk',
    axis_source_role: 'authoritative_realtime_axis',
    speaker_track_source: 'provider_or_local_observer',
    post_meeting_artifact_source: 'post_meeting_artifact_import',
  }),
  lark: Object.freeze({
    display_name: 'Feishu / Lark',
    recommended_mode: 'local_observer_first_provider_reconcile',
    axis_surface: 'browser_extension_or_desktop_observer',
    axis_source_role: 'primary_low_latency_axis',
    speaker_track_source: 'local_observer_or_detector',
    post_meeting_artifact_source: 'post_meeting_artifact_import',
  }),
  google_meet: Object.freeze({
    display_name: 'Google Meet',
    recommended_mode: 'local_observer_first_provider_reconcile',
    axis_surface: 'browser_extension',
    axis_source_role: 'primary_low_latency_axis',
    speaker_track_source: 'local_observer_or_detector',
    post_meeting_artifact_source: 'post_meeting_artifact_import',
  }),
  microsoft_teams: Object.freeze({
    display_name: 'Microsoft Teams',
    recommended_mode: 'local_observer_first_provider_reconcile',
    axis_surface: 'desktop_or_browser_observer',
    axis_source_role: 'primary_low_latency_axis',
    speaker_track_source: 'local_observer_or_detector',
    post_meeting_artifact_source: 'post_meeting_artifact_import',
  }),
  zoom: Object.freeze({
    display_name: 'Zoom',
    recommended_mode: 'local_observer_first_provider_reconcile',
    axis_surface: 'native_detector',
    axis_source_role: 'primary_low_latency_axis',
    speaker_track_source: 'local_observer_or_detector',
    post_meeting_artifact_source: 'post_meeting_artifact_import',
  }),
  webex: Object.freeze({
    display_name: 'Webex',
    recommended_mode: 'local_observer_first_provider_reconcile',
    axis_surface: 'browser_extension_or_native_detector',
    axis_source_role: 'primary_low_latency_axis',
    speaker_track_source: 'local_observer_or_detector',
    post_meeting_artifact_source: 'post_meeting_artifact_import',
  }),
});

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

function unique(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function selectedPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_PLATFORM_KEYS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

function setupOptionsFor(platform, options = {}) {
  const setupInput = firstNonEmpty(
    options.subscription,
    options.subscriptions?.[platform],
    options.providerSubscription,
    options.provider_subscription,
  );
  if (!setupInput) return options;
  const optionKey = {
    google_meet: 'googleMeetSubscription',
    microsoft_teams: 'microsoftTeamsSubscription',
    zoom: 'zoomSubscription',
    webex: 'webexSubscription',
  }[platform];
  return optionKey ? { ...options, [optionKey]: setupInput } : options;
}

function providerSubscriptionRequest(platform, setup = {}) {
  return firstNonEmpty(
    setup.workspace_subscription_request,
    setup.graph_subscription_request,
    setup.zoom_event_subscription_request,
    setup.webex_webhook_requests,
  );
}

function adapterSelectionFor(platform, options = {}) {
  const selected = firstNonEmpty(
    options.adapterSelection,
    options.adapter_selection,
    options.adapterSelections?.[platform],
    options.adapter_selections?.[platform],
  );
  if (selected) return selected;
  const localDetector = platform === 'local_detector';
  const defaults = ADAPTER_SELECTION_SNAPSHOT_DEFAULTS[platform] ?? {};
  return {
    schema: 'meeting_platform_adapter_selection_snapshot',
    schema_version: 1,
    platform,
    display_name: defaults.display_name,
    recommended_mode: defaults.recommended_mode,
    selection: {
      axis_source: localDetector ? 'host_detector_axis' : 'local_observer_axis',
      axis_surface: defaults.axis_surface,
      axis_source_role: defaults.axis_source_role,
      annotation_source: 'annotation_insert',
      timestamp_field: 'captured_at_ms',
      provider_reconcile_source: localDetector ? undefined : 'provider_reconcile',
      provider_reconcile_required_for_production: !localDetector,
      speaker_track_source: defaults.speaker_track_source,
      post_meeting_artifact_source: defaults.post_meeting_artifact_source,
    },
    runtime_policy: {
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      annotations_use_absolute_captured_at_ms: true,
      per_meeting_annotation_isolation_required: true,
      startup_order: [
        'validate_raw_signal',
        localDetector ? 'host_detector_axis' : 'local_observer_axis',
        'annotation_insert',
        'speaker_position_markers',
        localDetector ? undefined : 'provider_reconcile_non_blocking',
        'post_meeting_artifact_import_non_blocking',
      ].filter(Boolean),
      runtime_actions: [
        'validate_raw_signal',
        localDetector ? 'receive_host_detector_signal' : 'observe_platform_candidates',
        localDetector ? 'open_host_detector_axis' : 'open_adapter_session',
        'insert_annotation',
        'append_speaker_position_marker_if_available',
        localDetector ? undefined : 'reconcile_provider_event_when_available',
        'import_post_meeting_artifact_when_available',
      ].filter(Boolean),
    },
    readiness: {
      selection_ready: true,
      route_ready: true,
      pilot_evidence_ready: false,
      production_evidence_ready: false,
      provider_reconcile_required_for_production: !localDetector,
      missing: [],
    },
  };
}

function adapterSelectionSummary(selection = {}) {
  return compactObject({
    schema: selection.schema,
    schema_version: selection.schema_version,
    platform: selection.platform,
    display_name: selection.display_name,
    recommended_mode: selection.recommended_mode,
    selection: {
      axis_source: selection.selection?.axis_source,
      axis_surface: selection.selection?.axis_surface,
      axis_source_role: selection.selection?.axis_source_role,
      annotation_source: selection.selection?.annotation_source,
      timestamp_field: selection.selection?.timestamp_field,
      provider_reconcile_source: selection.selection?.provider_reconcile_source,
      provider_reconcile_required_for_production: selection.selection?.provider_reconcile_required_for_production,
      speaker_track_source: selection.selection?.speaker_track_source,
      post_meeting_artifact_source: selection.selection?.post_meeting_artifact_source,
    },
    runtime_policy: {
      provider_events_block_realtime: selection.runtime_policy?.provider_events_block_realtime,
      transcript_blocks_realtime: selection.runtime_policy?.transcript_blocks_realtime,
      annotations_use_absolute_captured_at_ms: selection.runtime_policy?.annotations_use_absolute_captured_at_ms,
      per_meeting_annotation_isolation_required: selection.runtime_policy?.per_meeting_annotation_isolation_required,
      startup_order: selection.runtime_policy?.startup_order,
      runtime_actions: selection.runtime_policy?.runtime_actions,
    },
    readiness: {
      selection_ready: selection.readiness?.selection_ready,
      pilot_evidence_ready: selection.readiness?.pilot_evidence_ready,
      production_evidence_ready: selection.readiness?.production_evidence_ready,
      provider_reconcile_required_for_production: selection.readiness?.provider_reconcile_required_for_production,
      missing: selection.readiness?.missing,
    },
  });
}

function runtimeBindingContract(platform, selection = {}) {
  return {
    realtime_axis_source: selection.selection?.axis_source,
    realtime_axis_surface: selection.selection?.axis_surface,
    annotation_source: selection.selection?.annotation_source,
    annotation_timestamp_field: selection.selection?.timestamp_field ?? 'captured_at_ms',
    speaker_track_source: selection.selection?.speaker_track_source,
    provider_reconcile_source: selection.selection?.provider_reconcile_source,
    provider_role: platform === 'local_detector'
      ? 'primary_low_latency_axis'
      : 'reconcile_and_backfill_after_local_axis',
    provider_reconcile_required_for_production: platform !== 'local_detector',
    provider_events_block_realtime: false,
    transcript_blocks_realtime: false,
    must_read_adapter_selection_before_session: true,
    per_meeting_annotation_isolation_required: true,
    startup_order: selection.runtime_policy?.startup_order,
  };
}

function missingSecurityNextActions(permissionPlan = {}) {
  return (permissionPlan.missing_security_env ?? []).map((name) => `configure_env:${name}`);
}

function readinessNextActions(readiness = {}) {
  return (readiness.checks ?? [])
    .filter((check) => check.ok !== true && check.severity === 'error')
    .map((check) => `fix_provider_setup:${check.id}`);
}

export function buildMeetingPlatformProviderConnectionPack(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const merged = setupOptionsFor(key, options);
  const manifest = platformSetupManifest(key, merged);
  const permissionPlan = buildPlatformPermissionPlan(key, merged);
  const integrationPlan = buildPlatformIntegrationPlan(key, merged);
  const setup = buildPlatformSetup(key, merged);
  const request = providerSubscriptionRequest(key, setup);
  const adapterSelection = adapterSelectionFor(key, merged);
  const bindingContract = runtimeBindingContract(key, adapterSelection);
  return compactObject({
    type: 'meeting_platform_provider_connection_pack',
    schema: MEETING_PLATFORM_PROVIDER_CONNECTION_PACK_SCHEMA,
    schema_version: MEETING_PLATFORM_PROVIDER_CONNECTION_SCHEMA_VERSION,
    platform: key,
    display_name: manifest.display_name,
    endpoint: manifest.endpoint,
    status_endpoint: manifest.status_endpoint,
    transport: manifest.transport,
    provider_role: key === 'local_detector'
      ? 'primary_low_latency_axis'
      : 'reconcile_and_backfill_after_local_axis',
    adapter_selection: adapterSelectionSummary(adapterSelection),
    runtime_binding_contract: bindingContract,
    official_docs: OFFICIAL_DOCS[key] ?? [],
    event_mapping: PROVIDER_EVENT_MAPPING[key] ?? [],
    subscription: compactObject({
      builders: manifest.builders ?? [],
      request,
      lifecycle_event_types: manifest.lifecycle_event_types ?? manifest.lifecycle_events,
      maintenance: integrationPlan.subscription_maintenance,
    }),
    permissions: {
      selected_features: permissionPlan.selected_features,
      required_permissions: permissionPlan.required_permissions,
      required_scopes: permissionPlan.required_scopes,
      admin_scopes: permissionPlan.admin_scopes,
    },
    security: compactObject({
      ...SECURITY_STRATEGY[key],
      required_env: permissionPlan.required_security_env,
      optional_env: permissionPlan.optional_security_env,
      present_env: permissionPlan.present_security_env,
      missing_env: permissionPlan.missing_security_env,
    }),
    realtime_annotation_policy: {
      annotation_timestamp_field: bindingContract.annotation_timestamp_field,
      provider_events_block_realtime: bindingContract.provider_events_block_realtime,
      transcript_blocks_realtime: bindingContract.transcript_blocks_realtime,
      local_observer_can_create_axis_before_provider: key !== 'local_detector',
      realtime_axis_source: bindingContract.realtime_axis_source,
      realtime_axis_surface: bindingContract.realtime_axis_surface,
      provider_reconcile_required_for_production: bindingContract.provider_reconcile_required_for_production,
    },
    setup_steps: manifest.required_setup ?? [],
    readiness: permissionPlan.readiness,
    integration: {
      recommended_mode: integrationPlan.recommended_mode,
      source_priority: integrationPlan.source_priority,
      provider_events: integrationPlan.provider_events,
      post_meeting_transcript: integrationPlan.post_meeting_transcript,
    },
    commands: {
      validate_provider_setup: 'buildMeetingPlatformProviderConnectionPack(platform, { baseUrl, env, subscription })',
      read_adapter_selection: `npm run meeting-platform:adapter-selection -- --platforms=${key} --json=true`,
      validate_provider_events: 'buildPlatformAcceptanceReport(platform, { samples, env, baseUrl })',
      validate_rollout: 'npm run meeting-platform:rollout-matrix',
      validate_live_readiness: 'npm run meeting-platform:live-readiness',
    },
    next_actions: unique([
      ...missingSecurityNextActions(permissionPlan),
      ...readinessNextActions(permissionPlan.readiness),
      key === 'local_detector' ? 'connect_trusted_local_detector' : 'capture_real_provider_events',
      key === 'local_detector' ? 'verify_timestamp_quality' : 'verify_signature_or_callback_auth',
      'read_adapter_selection_before_provider_reconcile',
      'keep_realtime_annotations_on_local_axis_until_provider_reconciles',
    ]),
  });
}

export function buildMeetingPlatformProviderConnectionMatrix(options = {}) {
  const packs = selectedPlatforms(options).map((platform) => buildMeetingPlatformProviderConnectionPack(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_provider_connection_matrix',
    schema: MEETING_PLATFORM_PROVIDER_CONNECTION_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_PROVIDER_CONNECTION_SCHEMA_VERSION,
    platform_count: packs.length,
    ready_count: packs.filter((pack) => pack.readiness?.ready === true).length,
    blocked_count: packs.filter((pack) => pack.readiness?.ready !== true).length,
    platforms: packs.map((pack) => pack.platform),
    rows: packs.map((pack) => ({
      platform: pack.platform,
      display_name: pack.display_name,
      transport: pack.transport,
      provider_role: pack.provider_role,
      adapter_selection_ready: pack.adapter_selection?.readiness?.selection_ready === true,
      realtime_axis_source: pack.runtime_binding_contract?.realtime_axis_source,
      realtime_axis_surface: pack.runtime_binding_contract?.realtime_axis_surface,
      annotation_timestamp_field: pack.runtime_binding_contract?.annotation_timestamp_field,
      provider_events_block_realtime: pack.runtime_binding_contract?.provider_events_block_realtime,
      ready: pack.readiness?.ready === true,
      missing_env: pack.security?.missing_env ?? [],
      event_count: pack.event_mapping?.length ?? 0,
      docs: pack.official_docs?.map((doc) => doc.url) ?? [],
      next_actions: pack.next_actions,
    })),
    packs,
    next_actions: unique(packs.flatMap((pack) => pack.next_actions ?? [])),
  };
}
