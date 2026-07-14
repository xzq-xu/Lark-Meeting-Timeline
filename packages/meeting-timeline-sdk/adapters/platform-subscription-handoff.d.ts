import type {
  MeetingPlatformProviderConnectionPack,
  MeetingPlatformProviderRuntimeBindingContract,
} from './platform-provider-connection.mjs';
import type { PlatformSetupOptions } from './platform-setup.mjs';

export const MEETING_PLATFORM_SUBSCRIPTION_HANDOFF_SCHEMA: 'meeting_platform_subscription_handoff';
export const MEETING_PLATFORM_SUBSCRIPTION_HANDOFF_MATRIX_SCHEMA: 'meeting_platform_subscription_handoff_matrix';
export const MEETING_PLATFORM_SUBSCRIPTION_HANDOFF_SCHEMA_VERSION: number;

export interface MeetingPlatformSubscriptionHandoffOptions extends PlatformSetupOptions {
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  subscription?: Record<string, unknown>;
  subscriptions?: Record<string, Record<string, unknown>>;
  providerSubscription?: Record<string, unknown>;
  provider_subscription?: Record<string, unknown>;
  providerSubscriptions?: Record<string, Record<string, unknown>>;
  provider_subscriptions?: Record<string, Record<string, unknown>>;
  now?: number | string | Date;
  renewalWindowMs?: number;
  renewal_window_ms?: number;
  renewalTtl?: string;
  renewal_ttl?: string;
  renewalTtlSeconds?: number;
  renewal_ttl_seconds?: number;
}

export interface MeetingPlatformSubscriptionRequest {
  id: string;
  platform: string;
  provider_api: string;
  method: string;
  path?: string;
  content_type: string;
  body: Record<string, unknown>;
}

export interface MeetingPlatformSubscriptionHandoff {
  type: 'meeting_platform_subscription_handoff';
  schema: 'meeting_platform_subscription_handoff';
  schema_version: number;
  platform: string;
  display_name?: string;
  status: string;
  ready_to_create: boolean;
  provider_ready: boolean;
  request_count: number;
  endpoint?: string;
  status_endpoint?: string;
  transport?: string;
  provider_role?: string;
  adapter_selection?: Record<string, unknown>;
  runtime_binding_contract: MeetingPlatformProviderRuntimeBindingContract;
  subscription_builders: string[];
  requests: MeetingPlatformSubscriptionRequest[];
  permissions: Record<string, unknown>;
  security: Record<string, unknown>;
  maintenance: Record<string, unknown>;
  realtime_annotation_policy: Record<string, unknown>;
  create_instructions: string[];
  handoff_contract: Record<string, unknown>;
  commands: Record<string, string>;
  official_docs: Array<Record<string, unknown>>;
  event_mapping: Array<Record<string, unknown>>;
  next_actions: string[];
  provider_connection: MeetingPlatformProviderConnectionPack;
}

export interface MeetingPlatformSubscriptionHandoffMatrix {
  type: 'meeting_platform_subscription_handoff_matrix';
  schema: 'meeting_platform_subscription_handoff_matrix';
  schema_version: number;
  platform_count: number;
  ready_to_create_count: number;
  request_count: number;
  security_blocked_count: number;
  parameter_missing_count: number;
  manual_setup_count: number;
  renewal_due_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  handoffs: MeetingPlatformSubscriptionHandoff[];
  next_actions: string[];
}

export function buildMeetingPlatformSubscriptionHandoff(
  platform: string,
  options?: MeetingPlatformSubscriptionHandoffOptions,
): MeetingPlatformSubscriptionHandoff;

export function buildMeetingPlatformSubscriptionHandoffMatrix(
  options?: MeetingPlatformSubscriptionHandoffOptions,
): MeetingPlatformSubscriptionHandoffMatrix;

export function assertMeetingPlatformSubscriptionHandoff(
  platform: string,
  options?: MeetingPlatformSubscriptionHandoffOptions,
): MeetingPlatformSubscriptionHandoff;

export function assertMeetingPlatformSubscriptionHandoffMatrix(
  options?: MeetingPlatformSubscriptionHandoffOptions,
): MeetingPlatformSubscriptionHandoffMatrix;
