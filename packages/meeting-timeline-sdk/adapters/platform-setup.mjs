import { MeetingTimelineSdkError, compactObject } from '../index.mjs';

export const GOOGLE_MEET_EVENT_TYPES = Object.freeze([
  'google.workspace.meet.conference.v2.started',
  'google.workspace.meet.conference.v2.ended',
  'google.workspace.meet.participant.v2.joined',
  'google.workspace.meet.participant.v2.left',
  'google.workspace.meet.recording.v2.fileGenerated',
  'google.workspace.meet.transcript.v2.fileGenerated',
  'google.workspace.meet.smartNote.v2.fileGenerated',
]);

export const MICROSOFT_TEAMS_CHANGE_TYPES = Object.freeze(['created', 'updated']);

export const ZOOM_MEETING_EVENT_TYPES = Object.freeze([
  'meeting.started',
  'meeting.ended',
  'meeting.participant_joined',
  'meeting.participant_left',
  'recording.completed',
]);

const platformAliases = new Map([
  ['google-meet', 'google_meet'],
  ['google_meet', 'google_meet'],
  ['meet', 'google_meet'],
  ['microsoft-teams', 'microsoft_teams'],
  ['microsoft_teams', 'microsoft_teams'],
  ['teams', 'microsoft_teams'],
  ['zoom', 'zoom'],
]);

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function ensureBaseUrl(baseUrl) {
  if (!baseUrl) throw new MeetingTimelineSdkError('baseUrl is required to build platform setup');
  return String(baseUrl).replace(/\/+$/, '');
}

function absoluteEndpoint(baseUrl, path) {
  return `${ensureBaseUrl(baseUrl)}${path.startsWith('/') ? path : `/${path}`}`;
}

function addSeconds(date, seconds) {
  return new Date(date.getTime() + seconds * 1000).toISOString();
}

function normalizePlatform(platform) {
  const key = platformAliases.get(String(platform ?? '').toLowerCase().replace(/\s+/g, '-'));
  if (!key) {
    throw new MeetingTimelineSdkError(`Unsupported meeting platform: ${platform}`, {
      supportedPlatforms: Array.from(new Set(platformAliases.values())),
    });
  }
  return key;
}

function urlEncode(value) {
  return encodeURIComponent(String(value));
}

function envValue(env = {}, name) {
  return env[name] ?? env[name.toLowerCase()] ?? env[name.replace(/_/g, '-')];
}

function hasEnv(env = {}, name) {
  const value = envValue(env, name);
  return value != null && value !== '';
}

function endpointReadiness(endpoint) {
  if (!endpoint) return { ok: false, reason: 'endpoint_missing' };
  let parsed;
  try {
    parsed = new URL(endpoint);
  } catch {
    return { ok: false, reason: 'endpoint_invalid' };
  }
  const host = parsed.hostname.toLowerCase();
  const local = ['localhost', '127.0.0.1', '::1'].includes(host);
  if (parsed.protocol !== 'https:' && !local) {
    return { ok: false, reason: 'endpoint_must_be_https_or_localhost', protocol: parsed.protocol };
  }
  return { ok: true, reason: local ? 'local_development_endpoint' : 'public_https_endpoint' };
}

function missingEnv(env = {}, names = []) {
  return names.filter((name) => !hasEnv(env, name));
}

function platformReadinessChecks(platform, manifest = {}, env = {}) {
  const endpointCheck = endpointReadiness(manifest.endpoint);
  const checks = [{
    id: 'endpoint',
    ok: endpointCheck.ok,
    severity: endpointCheck.ok ? 'info' : 'error',
    detail: endpointCheck.reason,
  }];
  if (platform === 'google_meet') {
    const oidcConfigured = hasEnv(env, 'GOOGLE_PUBSUB_OIDC_AUDIENCE');
    const bearerConfigured = hasEnv(env, 'GOOGLE_PUBSUB_BEARER_TOKEN');
    checks.push({
      id: 'google_pubsub_auth',
      ok: oidcConfigured || bearerConfigured,
      severity: oidcConfigured ? 'info' : bearerConfigured ? 'warn' : 'error',
      detail: oidcConfigured
        ? 'oidc_configured'
        : bearerConfigured
          ? 'bearer_fallback_configured'
          : 'missing_google_pubsub_oidc_or_bearer',
      required_any_env: ['GOOGLE_PUBSUB_OIDC_AUDIENCE', 'GOOGLE_PUBSUB_BEARER_TOKEN'],
    });
    checks.push({
      id: 'google_pubsub_service_account_email',
      ok: hasEnv(env, 'GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL'),
      severity: hasEnv(env, 'GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL') ? 'info' : 'warn',
      detail: hasEnv(env, 'GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL')
        ? 'service_account_email_bound'
        : 'service_account_email_not_bound',
    });
  } else {
    const missing = missingEnv(env, manifest.required_security_env ?? []);
    checks.push({
      id: 'required_security_env',
      ok: missing.length === 0,
      severity: missing.length === 0 ? 'info' : 'error',
      detail: missing.length === 0 ? 'required_security_env_configured' : 'missing_required_security_env',
      missing_env: missing,
    });
  }
  return checks;
}

function readinessFromChecks(checks = []) {
  const blocking = checks.filter((item) => item.severity === 'error' && item.ok !== true);
  const warnings = checks.filter((item) => item.severity === 'warn' && item.ok !== true);
  return {
    ready: blocking.length === 0,
    warning_count: warnings.length,
    blocking_count: blocking.length,
    checks,
  };
}

export function buildGoogleMeetWorkspaceSubscriptionRequest(input = {}) {
  const targetResource = firstNonEmpty(input.targetResource, input.target_resource);
  const pubsubTopic = firstNonEmpty(input.pubsubTopic, input.pubsub_topic, input.topic);
  if (!targetResource) {
    throw new MeetingTimelineSdkError('targetResource is required for Google Meet subscription setup');
  }
  if (!pubsubTopic) {
    throw new MeetingTimelineSdkError('pubsubTopic is required for Google Meet subscription setup');
  }
  return compactObject({
    targetResource,
    eventTypes: input.eventTypes ?? input.event_types ?? GOOGLE_MEET_EVENT_TYPES,
    notificationEndpoint: {
      pubsubTopic,
    },
    payloadOptions: input.includeResource === false || input.include_resource === false
      ? undefined
      : { includeResource: true },
    ttl: input.ttl,
  });
}

export function buildMicrosoftTeamsMeetingCallSubscriptionRequest(input = {}) {
  const joinWebUrl = firstNonEmpty(input.joinWebUrl, input.join_web_url, input.meetingUrl, input.meeting_url);
  const notificationUrl = firstNonEmpty(input.notificationUrl, input.notification_url, input.webhookUrl, input.webhook_url);
  const clientState = firstNonEmpty(input.clientState, input.client_state);
  if (!joinWebUrl) {
    throw new MeetingTimelineSdkError('joinWebUrl is required for Microsoft Teams meetingCallEvents subscription setup');
  }
  if (!notificationUrl) {
    throw new MeetingTimelineSdkError('notificationUrl is required for Microsoft Teams subscription setup');
  }
  const now = input.now instanceof Date ? input.now : new Date(input.now ?? Date.now());
  return compactObject({
    changeType: input.changeType ?? input.change_type ?? MICROSOFT_TEAMS_CHANGE_TYPES.join(','),
    notificationUrl,
    resource: `/communications/onlineMeetings(joinWebUrl='${urlEncode(joinWebUrl)}')/meetingCallEvents`,
    expirationDateTime: input.expirationDateTime
      ?? input.expiration_date_time
      ?? addSeconds(now, Number(input.ttlSeconds ?? input.ttl_seconds ?? 2 * 24 * 60 * 60)),
    clientState,
    includeResourceData: input.includeResourceData ?? input.include_resource_data ?? false,
    encryptionCertificate: input.encryptionCertificate ?? input.encryption_certificate,
    encryptionCertificateId: input.encryptionCertificateId ?? input.encryption_certificate_id,
    lifecycleNotificationUrl: input.lifecycleNotificationUrl ?? input.lifecycle_notification_url,
  });
}

export function buildZoomEventSubscriptionRequest(input = {}) {
  const webhookUrl = firstNonEmpty(input.eventWebhookUrl, input.event_webhook_url, input.webhookUrl, input.webhook_url);
  if (!webhookUrl) {
    throw new MeetingTimelineSdkError('webhookUrl is required for Zoom event subscription setup');
  }
  return compactObject({
    event_subscription_name: input.name ?? input.eventSubscriptionName ?? input.event_subscription_name ?? 'Meeting Timeline Events',
    event_webhook_url: webhookUrl,
    events: input.events ?? ZOOM_MEETING_EVENT_TYPES,
    subscription_scope: input.subscriptionScope ?? input.subscription_scope ?? 'account',
    account_id: input.accountId ?? input.account_id,
    user_ids: input.userIds ?? input.user_ids,
  });
}

export function platformEventEndpoint(baseUrl, platform) {
  const key = normalizePlatform(platform);
  const path = {
    google_meet: '/api/platform-events/google-meet',
    microsoft_teams: '/api/platform-events/teams',
    zoom: '/api/platform-events/zoom',
  }[key];
  return absoluteEndpoint(baseUrl, path);
}

export function platformSetupManifest(platform, options = {}) {
  const key = normalizePlatform(platform);
  const endpoint = options.baseUrl ? platformEventEndpoint(options.baseUrl, key) : null;
  const statusEndpoint = options.baseUrl
    ? `${endpoint}/status`
    : null;
  const manifests = {
    google_meet: {
      platform: 'google_meet',
      display_name: 'Google Meet',
      endpoint,
      status_endpoint: statusEndpoint,
      transport: 'Google Workspace Events API -> Google Cloud Pub/Sub push',
      default_event_types: GOOGLE_MEET_EVENT_TYPES,
      required_security_env: ['GOOGLE_PUBSUB_OIDC_AUDIENCE'],
      optional_security_env: ['GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL', 'GOOGLE_PUBSUB_BEARER_TOKEN'],
      required_setup: [
        'Enable Google Meet REST API, Google Workspace Events API, and Pub/Sub.',
        'Create a Pub/Sub topic and push subscription targeting the endpoint.',
        'Configure Pub/Sub authenticated push OIDC audience to match GOOGLE_PUBSUB_OIDC_AUDIENCE.',
        'Create a Workspace Events subscription for a Meet space or user target resource.',
      ],
      builders: ['buildGoogleMeetWorkspaceSubscriptionRequest'],
    },
    microsoft_teams: {
      platform: 'microsoft_teams',
      display_name: 'Microsoft Teams',
      endpoint,
      status_endpoint: statusEndpoint,
      transport: 'Microsoft Graph change notifications',
      default_change_types: MICROSOFT_TEAMS_CHANGE_TYPES,
      resource_template: "/communications/onlineMeetings(joinWebUrl='{encodedJoinWebUrl}')/meetingCallEvents",
      required_permissions: ['OnlineMeetings.Read.All or OnlineMeetings.ReadWrite.All'],
      required_security_env: ['MICROSOFT_GRAPH_CLIENT_STATE'],
      required_setup: [
        'Create a Graph application permission grant for OnlineMeetings.Read.All or OnlineMeetings.ReadWrite.All.',
        'Create a change notification subscription for the meeting joinWebUrl.',
        'Renew the subscription before its maximum 3 day expiration.',
        'Use rich notifications with includeResourceData for active meeting call changes when available.',
      ],
      builders: ['buildMicrosoftTeamsMeetingCallSubscriptionRequest'],
    },
    zoom: {
      platform: 'zoom',
      display_name: 'Zoom',
      endpoint,
      status_endpoint: statusEndpoint,
      transport: 'Zoom Meeting webhooks',
      default_event_types: ZOOM_MEETING_EVENT_TYPES,
      required_security_env: ['ZOOM_WEBHOOK_SECRET_TOKEN'],
      required_scopes: [
        'meeting:read:meeting or meeting:read:meeting:admin',
        'meeting:read:participant or meeting:read:participant:admin',
        'cloud_recording:read:recording or cloud_recording:read:recording:admin',
      ],
      required_setup: [
        'Enable Event Subscriptions in the Zoom app or create an event subscription by API.',
        'Set the endpoint as the event notification URL and complete URL validation.',
        'Add scopes required by selected events before creating API-managed subscriptions.',
        'Return 2xx within 3 seconds and verify x-zm-signature on every request.',
      ],
      builders: ['buildZoomEventSubscriptionRequest'],
    },
  };
  return compactObject(manifests[key]);
}

export function allPlatformSetupManifests(options = {}) {
  return ['google_meet', 'microsoft_teams', 'zoom'].map((platform) => platformSetupManifest(platform, options));
}

export function buildPlatformSetup(platform, options = {}) {
  const key = normalizePlatform(platform);
  const manifest = platformSetupManifest(key, options);
  if (key === 'google_meet') {
    return compactObject({
      ...manifest,
      workspace_subscription_request: options.googleMeetSubscription
        ? buildGoogleMeetWorkspaceSubscriptionRequest(options.googleMeetSubscription)
        : undefined,
    });
  }
  if (key === 'microsoft_teams') {
    return compactObject({
      ...manifest,
      graph_subscription_request: options.microsoftTeamsSubscription
        ? buildMicrosoftTeamsMeetingCallSubscriptionRequest(options.microsoftTeamsSubscription)
        : undefined,
    });
  }
  if (key === 'zoom') {
    return compactObject({
      ...manifest,
      zoom_event_subscription_request: options.zoomSubscription
        ? buildZoomEventSubscriptionRequest(options.zoomSubscription)
        : undefined,
    });
  }
  return manifest;
}

export function evaluatePlatformSetupReadiness(platform, options = {}) {
  const key = normalizePlatform(platform);
  const manifest = platformSetupManifest(key, options);
  const env = options.env ?? {};
  return {
    platform: key,
    ...readinessFromChecks(platformReadinessChecks(key, manifest, env)),
  };
}

export function evaluateAllPlatformSetupReadiness(options = {}) {
  return ['google_meet', 'microsoft_teams', 'zoom'].map((platform) => (
    evaluatePlatformSetupReadiness(platform, options)
  ));
}

export const MEETING_PLATFORM_SETUP_BUILDERS = Object.freeze({
  buildGoogleMeetWorkspaceSubscriptionRequest,
  buildMicrosoftTeamsMeetingCallSubscriptionRequest,
  buildZoomEventSubscriptionRequest,
  platformSetupManifest,
  allPlatformSetupManifests,
  buildPlatformSetup,
  evaluatePlatformSetupReadiness,
  evaluateAllPlatformSetupReadiness,
});
