export const LARK_MEETING_EVENT_TYPES: readonly string[];
export const LOCAL_DETECTOR_EVENT_TYPES: readonly string[];
export const GOOGLE_MEET_EVENT_TYPES: readonly string[];
export const GOOGLE_WORKSPACE_SUBSCRIPTION_LIFECYCLE_EVENT_TYPES: readonly string[];
export const MICROSOFT_TEAMS_CHANGE_TYPES: readonly string[];
export const MICROSOFT_GRAPH_LIFECYCLE_EVENTS: readonly string[];
export const ZOOM_MEETING_EVENT_TYPES: readonly string[];
export const WEBEX_WEBHOOK_RESOURCES: readonly { resource: string; events: readonly string[]; filter?: string }[];
export const MEETING_PLATFORM_KEYS: readonly string[];
export const MEETING_PLATFORM_ALIASES: Readonly<Record<string, string>>;

export interface PlatformSetupOptions {
  baseUrl?: string;
  env?: Record<string, unknown>;
  googleMeetSubscription?: Record<string, unknown>;
  microsoftTeamsSubscription?: Record<string, unknown>;
  zoomSubscription?: Record<string, unknown>;
  webexSubscription?: Record<string, unknown>;
  [key: string]: unknown;
}

export function buildGoogleMeetWorkspaceSubscriptionRequest(input?: {
  targetResource?: string;
  target_resource?: string;
  pubsubTopic?: string;
  pubsub_topic?: string;
  topic?: string;
  eventTypes?: string[];
  event_types?: string[];
  includeResource?: boolean;
  include_resource?: boolean;
  ttl?: string;
  [key: string]: unknown;
}): Record<string, unknown>;

export function buildMicrosoftTeamsMeetingCallSubscriptionRequest(input?: {
  joinWebUrl?: string;
  join_web_url?: string;
  meetingUrl?: string;
  meeting_url?: string;
  notificationUrl?: string;
  notification_url?: string;
  webhookUrl?: string;
  webhook_url?: string;
  clientState?: string;
  client_state?: string;
  changeType?: string;
  change_type?: string;
  expirationDateTime?: string;
  expiration_date_time?: string;
  ttlSeconds?: number;
  ttl_seconds?: number;
  includeResourceData?: boolean;
  include_resource_data?: boolean;
  encryptionCertificate?: string;
  encryption_certificate?: string;
  encryptionCertificateId?: string;
  encryption_certificate_id?: string;
  lifecycleNotificationUrl?: string;
  lifecycle_notification_url?: string;
  now?: number | string | Date;
  [key: string]: unknown;
}): Record<string, unknown>;

export function buildZoomEventSubscriptionRequest(input?: {
  eventWebhookUrl?: string;
  event_webhook_url?: string;
  webhookUrl?: string;
  webhook_url?: string;
  name?: string;
  eventSubscriptionName?: string;
  event_subscription_name?: string;
  events?: string[];
  subscriptionScope?: string;
  subscription_scope?: string;
  accountId?: string;
  account_id?: string;
  userIds?: string[];
  user_ids?: string[];
  [key: string]: unknown;
}): Record<string, unknown>;

export function buildWebexWebhookRequests(input?: {
  targetUrl?: string;
  target_url?: string;
  webhookUrl?: string;
  webhook_url?: string;
  name?: string;
  secret?: string;
  webhookSecret?: string;
  webhook_secret?: string;
  ownedBy?: string;
  owned_by?: string;
  status?: string;
  filter?: string;
  resources?: { resource: string; events: string[]; filter?: string }[];
  [key: string]: unknown;
}): Record<string, unknown>[];

export function buildMicrosoftGraphSubscriptionRenewalRequest(input?: {
  subscriptionId?: string;
  subscription_id?: string;
  id?: string;
  expirationDateTime?: string;
  expiration_date_time?: string;
  expiresAt?: string;
  expires_at?: string;
  ttlSeconds?: number;
  ttl_seconds?: number;
  now?: number | string | Date;
  [key: string]: unknown;
}): Record<string, unknown>;

export function buildGoogleWorkspaceSubscriptionRenewalRequest(input?: {
  subscriptionName?: string;
  subscription_name?: string;
  name?: string;
  ttl?: string;
  ttlSeconds?: number;
  ttl_seconds?: number;
  expireTime?: string;
  expire_time?: string;
  expiresAt?: string;
  expires_at?: string;
  [key: string]: unknown;
}): Record<string, unknown>;

export function normalizeMeetingPlatform(platform: string): string;
export function platformEventEndpoint(baseUrl: string, platform: string): string;
export function platformCapabilityContract(platform: string, options?: PlatformSetupOptions): Record<string, unknown>;
export function allPlatformCapabilityContracts(options?: PlatformSetupOptions): Record<string, unknown>[];
export function platformSetupManifest(platform: string, options?: PlatformSetupOptions): Record<string, unknown>;
export function allPlatformSetupManifests(options?: PlatformSetupOptions): Record<string, unknown>[];
export function buildPlatformPermissionPlan(platform: string, options?: PlatformSetupOptions & {
  features?: string[] | string;
  featureSet?: string[] | string;
  feature_set?: string[] | string;
}): Record<string, unknown>;
export function allPlatformPermissionPlans(options?: PlatformSetupOptions & {
  features?: string[] | string;
  featureSet?: string[] | string;
  feature_set?: string[] | string;
}): Record<string, unknown>[];
export function buildPlatformIntegrationPlan(platform: string, options?: PlatformSetupOptions & {
  axisMode?: string;
  axis_mode?: string;
  subscription?: Record<string, unknown>;
  subscriptions?: Record<string, Record<string, unknown>>;
  now?: number | string | Date;
  renewalWindowMs?: number;
  renewal_window_ms?: number;
  renewalTtl?: string;
  renewal_ttl?: string;
  renewalTtlSeconds?: number;
  renewal_ttl_seconds?: number;
}): Record<string, unknown>;
export function allPlatformIntegrationPlans(options?: PlatformSetupOptions & {
  axisMode?: string;
  axis_mode?: string;
  subscriptions?: Record<string, Record<string, unknown>>;
  now?: number | string | Date;
}): Record<string, unknown>[];
export function buildPlatformSetup(platform: string, options?: PlatformSetupOptions): Record<string, unknown>;
export function evaluatePlatformSetupReadiness(platform: string, options?: PlatformSetupOptions): Record<string, unknown>;
export function evaluateAllPlatformSetupReadiness(options?: PlatformSetupOptions): Record<string, unknown>[];
export function evaluatePlatformSubscriptionMaintenance(platform: string, subscription?: Record<string, unknown>, options?: Record<string, unknown>): Record<string, unknown>;
export function evaluateAllPlatformSubscriptionMaintenance(subscriptions?: Record<string, Record<string, unknown>>, options?: Record<string, unknown>): Record<string, unknown>[];

export const MEETING_PLATFORM_SETUP_BUILDERS: Readonly<Record<string, unknown>>;
