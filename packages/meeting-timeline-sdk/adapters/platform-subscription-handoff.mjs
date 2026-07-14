import { MeetingTimelineSdkError, compactObject } from './internal-utils.mjs';
import { buildMeetingPlatformProviderConnectionPack } from './platform-provider-connection.mjs';
import {
  evaluatePlatformSubscriptionMaintenance,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_SUBSCRIPTION_HANDOFF_SCHEMA = 'meeting_platform_subscription_handoff';
export const MEETING_PLATFORM_SUBSCRIPTION_HANDOFF_MATRIX_SCHEMA = 'meeting_platform_subscription_handoff_matrix';
export const MEETING_PLATFORM_SUBSCRIPTION_HANDOFF_SCHEMA_VERSION = 1;

const DEFAULT_SUBSCRIPTION_HANDOFF_PLATFORMS = Object.freeze([
  'google_meet',
  'microsoft_teams',
  'zoom',
  'webex',
  'lark',
]);

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

function hasSubscriptionInput(value) {
  if (value == null || value === '') return false;
  if (Array.isArray(value)) return value.length > 0;
  if (typeof value === 'object') return Object.keys(value).length > 0;
  return true;
}

function selectedPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(
    options.platforms,
    options.platform_keys,
    DEFAULT_SUBSCRIPTION_HANDOFF_PLATFORMS,
  )).map((platform) => normalizeMeetingPlatform(platform)));
}

function subscriptionFor(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  return [
    options.subscription,
    options.subscriptions?.[key],
    options.providerSubscription,
    options.provider_subscription,
    options.providerSubscriptions?.[key],
    options.provider_subscriptions?.[key],
    options[key],
    options.platforms?.[key],
  ].find((value) => hasSubscriptionInput(value));
}

function requestList(platform, request) {
  if (request == null) return [];
  const key = normalizeMeetingPlatform(platform);
  if (Array.isArray(request)) {
    return request.map((body, index) => normalizedRequest(key, body, index));
  }
  return [normalizedRequest(key, request, 0)];
}

function normalizedRequest(platform, body = {}, index = 0) {
  const metadata = {
    google_meet: {
      provider_api: 'google_workspace_events',
      method: 'POST',
      path: '/v1beta/subscriptions',
      content_type: 'application/json',
    },
    microsoft_teams: {
      provider_api: 'microsoft_graph_change_notifications',
      method: 'POST',
      path: '/subscriptions',
      content_type: 'application/json',
    },
    zoom: {
      provider_api: 'zoom_event_subscription',
      method: 'PUT_OR_MANUAL_CONFIG',
      path: '/marketplace/app/event-subscriptions',
      content_type: 'application/json',
    },
    webex: {
      provider_api: 'webex_webhooks',
      method: 'POST',
      path: '/v1/webhooks',
      content_type: 'application/json',
    },
  }[platform] ?? {
    provider_api: `${platform}_manual_setup`,
    method: 'MANUAL',
    path: undefined,
    content_type: 'application/json',
  };
  return compactObject({
    id: `${platform}_subscription_request_${index + 1}`,
    platform,
    ...metadata,
    body,
  });
}

function statusFor(platform, pack = {}, requests = []) {
  const missingEnv = pack.security?.missing_env ?? [];
  const hasBuilder = (pack.subscription?.builders ?? []).length > 0;
  if (platform === 'local_detector') return 'not_applicable';
  if (platform === 'lark') return missingEnv.length > 0 ? 'manual_setup_missing_env' : 'manual_setup';
  if (missingEnv.length > 0) return 'security_env_missing';
  if (hasBuilder && requests.length === 0) return 'needs_subscription_parameters';
  if (pack.readiness?.ready !== true) return 'provider_connection_not_ready';
  return 'subscription_request_ready';
}

function createInstructions(platform, requests = []) {
  const key = normalizeMeetingPlatform(platform);
  const defaultInstructions = {
    google_meet: [
      'Create or confirm the Pub/Sub topic and authenticated push subscription first.',
      'POST the body to Google Workspace Events subscriptions API with a token that has the requested Meet/Drive scopes.',
      'Store the returned subscription name and expiration for maintenance checks.',
    ],
    microsoft_teams: [
      'POST the body to Microsoft Graph /subscriptions with application permissions granted by the tenant admin.',
      'Persist subscription id, expirationDateTime, resource, and clientState.',
      'Renew before expiration; Graph meeting call subscriptions are short lived.',
    ],
    zoom: [
      'Configure the event subscription on the Zoom app or app API using the generated event list and webhook URL.',
      'Complete endpoint URL validation before enabling events.',
      'Persist the configured event subscription id or app id for audit.',
    ],
    webex: [
      'POST each generated body to Webex /v1/webhooks.',
      'Persist every webhook id because Webex uses one webhook per resource/event.',
      'Verify X-Spark-Signature with the same shared secret configured in each webhook.',
    ],
    lark: [
      'Configure Feishu/Lark event subscriptions or long-connection delivery in the Open Platform console.',
      'Use meeting start/end events for provider reconcile and current-user scan/local observer for low-latency fallback.',
      'Persist app id and event delivery mode for support diagnostics.',
    ],
  };
  return [
    ...(defaultInstructions[key] ?? ['Use provider console or API to configure the event delivery target.']),
    requests.length > 0 ? 'After creation, run meeting-platform:provider-connection and meeting-platform:subscription-handoff again with returned subscription metadata.' : undefined,
  ].filter(Boolean);
}

function nextActions(platform, status, pack = {}, maintenance = {}) {
  const mapped = {
    not_applicable: 'connect_local_detector_instead_of_provider_subscription',
    manual_setup: 'record_manual_provider_subscription_metadata',
    manual_setup_missing_env: 'complete_manual_provider_security_configuration',
    security_env_missing: 'configure_provider_security_env',
    needs_subscription_parameters: 'provide_subscription_creation_parameters',
    provider_connection_not_ready: 'fix_provider_connection_readiness',
    subscription_request_ready: 'create_provider_subscription_and_capture_returned_metadata',
  }[status];
  return unique([
    mapped,
    ...(pack.security?.missing_env ?? []).map((name) => `configure_env:${name}`),
    ...(maintenance.renewal_due ? ['renew_provider_subscription'] : []),
    ...(pack.next_actions ?? []),
  ]);
}

export function buildMeetingPlatformSubscriptionHandoff(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const subscription = subscriptionFor(key, options);
  const providerOptions = {
    ...options,
    subscription,
    subscriptions: undefined,
    providerSubscription: undefined,
    provider_subscription: undefined,
    providerSubscriptions: undefined,
    provider_subscriptions: undefined,
  };
  const pack = buildMeetingPlatformProviderConnectionPack(key, {
    ...providerOptions,
  });
  const requests = requestList(key, pack.subscription?.request);
  const maintenance = evaluatePlatformSubscriptionMaintenance(key, subscription ?? {}, options);
  const status = statusFor(key, pack, requests);
  return compactObject({
    type: 'meeting_platform_subscription_handoff',
    schema: MEETING_PLATFORM_SUBSCRIPTION_HANDOFF_SCHEMA,
    schema_version: MEETING_PLATFORM_SUBSCRIPTION_HANDOFF_SCHEMA_VERSION,
    platform: key,
    display_name: pack.display_name,
    status,
    ready_to_create: status === 'subscription_request_ready',
    provider_ready: pack.readiness?.ready === true,
    request_count: requests.length,
    endpoint: pack.endpoint,
    status_endpoint: pack.status_endpoint,
    transport: pack.transport,
    provider_role: pack.provider_role,
    adapter_selection: pack.adapter_selection,
    runtime_binding_contract: pack.runtime_binding_contract,
    subscription_builders: pack.subscription?.builders ?? [],
    requests,
    permissions: pack.permissions,
    security: pack.security,
    maintenance,
    realtime_annotation_policy: pack.realtime_annotation_policy,
    create_instructions: createInstructions(key, requests),
    handoff_contract: {
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      annotation_timestamp_field: 'captured_at_ms',
      realtime_axis_source: pack.runtime_binding_contract?.realtime_axis_source,
      realtime_axis_surface: pack.runtime_binding_contract?.realtime_axis_surface,
      adapter_selection_snapshot_required: true,
      provider_reconcile_non_blocking: true,
      post_meeting_artifact_non_blocking: true,
      returned_metadata_required: ['subscription_id_or_name', 'expiration_when_provider_returns_it', 'provider_resource_or_event_list'],
    },
    commands: {
      export_subscription_handoff: `npm run meeting-platform:subscription-handoff -- --platforms=${key} --json=true`,
      validate_provider_connection: 'npm run meeting-platform:provider-connection',
      validate_real_intake: 'npm run meeting-platform:real-intake',
      validate_handoff_readiness: 'npm run meeting-platform:handoff-readiness',
    },
    official_docs: pack.official_docs,
    event_mapping: pack.event_mapping,
    next_actions: nextActions(key, status, pack, maintenance),
    provider_connection: pack,
  });
}

export function buildMeetingPlatformSubscriptionHandoffMatrix(options = {}) {
  const platforms = selectedPlatforms(options);
  const handoffs = platforms.map((platform) => buildMeetingPlatformSubscriptionHandoff(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_subscription_handoff_matrix',
    schema: MEETING_PLATFORM_SUBSCRIPTION_HANDOFF_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_SUBSCRIPTION_HANDOFF_SCHEMA_VERSION,
    platform_count: handoffs.length,
    ready_to_create_count: handoffs.filter((handoff) => handoff.ready_to_create).length,
    request_count: handoffs.reduce((total, handoff) => total + handoff.request_count, 0),
    security_blocked_count: handoffs.filter((handoff) => handoff.status === 'security_env_missing' || handoff.status === 'manual_setup_missing_env').length,
    parameter_missing_count: handoffs.filter((handoff) => handoff.status === 'needs_subscription_parameters').length,
    manual_setup_count: handoffs.filter((handoff) => String(handoff.status).startsWith('manual_setup')).length,
    renewal_due_count: handoffs.filter((handoff) => handoff.maintenance?.renewal_due === true).length,
    platforms,
    rows: handoffs.map((handoff) => ({
      platform: handoff.platform,
      display_name: handoff.display_name,
      status: handoff.status,
      ready_to_create: handoff.ready_to_create,
      provider_ready: handoff.provider_ready,
      adapter_selection_ready: handoff.adapter_selection?.readiness?.selection_ready === true,
      realtime_axis_source: handoff.runtime_binding_contract?.realtime_axis_source,
      realtime_axis_surface: handoff.runtime_binding_contract?.realtime_axis_surface,
      annotation_timestamp_field: handoff.runtime_binding_contract?.annotation_timestamp_field,
      provider_events_block_realtime: handoff.runtime_binding_contract?.provider_events_block_realtime,
      request_count: handoff.request_count,
      missing_env: handoff.security?.missing_env ?? [],
      renewal_due: handoff.maintenance?.renewal_due === true,
      renewal_status: handoff.maintenance?.status,
      subscription_builders: handoff.subscription_builders,
      first_next_action: handoff.next_actions?.[0],
      next_actions: handoff.next_actions,
    })),
    handoffs,
    next_actions: unique(handoffs.flatMap((handoff) => handoff.next_actions ?? [])),
  };
}

export function assertMeetingPlatformSubscriptionHandoff(platform, options = {}) {
  const handoff = buildMeetingPlatformSubscriptionHandoff(platform, options);
  if (handoff.ready_to_create !== true && handoff.status !== 'manual_setup') {
    throw new MeetingTimelineSdkError(`Meeting platform subscription handoff is not ready for ${handoff.platform}`, {
      platform: handoff.platform,
      status: handoff.status,
      next_actions: handoff.next_actions,
      handoff,
    });
  }
  return handoff;
}

export function assertMeetingPlatformSubscriptionHandoffMatrix(options = {}) {
  const matrix = buildMeetingPlatformSubscriptionHandoffMatrix(options);
  const failed = matrix.handoffs.filter((handoff) => handoff.ready_to_create !== true && handoff.status !== 'manual_setup');
  if (failed.length > 0) {
    throw new MeetingTimelineSdkError('Meeting platform subscription handoff matrix is not ready', {
      platform_count: matrix.platform_count,
      failed_platforms: failed.map((handoff) => handoff.platform),
      next_actions: matrix.next_actions,
      matrix,
    });
  }
  return matrix;
}
