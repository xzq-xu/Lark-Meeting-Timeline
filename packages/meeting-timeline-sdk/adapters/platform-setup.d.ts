export const GOOGLE_MEET_EVENT_TYPES: readonly string[];
export const MICROSOFT_TEAMS_CHANGE_TYPES: readonly string[];
export const ZOOM_MEETING_EVENT_TYPES: readonly string[];

export interface PlatformSetupOptions {
  baseUrl?: string;
  googleMeetSubscription?: Record<string, unknown>;
  microsoftTeamsSubscription?: Record<string, unknown>;
  zoomSubscription?: Record<string, unknown>;
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

export function platformEventEndpoint(baseUrl: string, platform: string): string;
export function platformSetupManifest(platform: string, options?: PlatformSetupOptions): Record<string, unknown>;
export function allPlatformSetupManifests(options?: PlatformSetupOptions): Record<string, unknown>[];
export function buildPlatformSetup(platform: string, options?: PlatformSetupOptions): Record<string, unknown>;

export const MEETING_PLATFORM_SETUP_BUILDERS: Readonly<Record<string, unknown>>;
