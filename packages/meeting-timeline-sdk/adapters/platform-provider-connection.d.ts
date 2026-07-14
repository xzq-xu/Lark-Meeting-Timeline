import type { PlatformSetupOptions } from './platform-setup.mjs';
import type { MeetingPlatformAdapterSelection } from './platform-adapter-selection.mjs';

export const MEETING_PLATFORM_PROVIDER_CONNECTION_PACK_SCHEMA: string;
export const MEETING_PLATFORM_PROVIDER_CONNECTION_MATRIX_SCHEMA: string;
export const MEETING_PLATFORM_PROVIDER_CONNECTION_SCHEMA_VERSION: number;

export interface MeetingPlatformProviderConnectionOptions extends PlatformSetupOptions {
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  subscription?: Record<string, unknown>;
  subscriptions?: Record<string, Record<string, unknown>>;
  providerSubscription?: Record<string, unknown>;
  provider_subscription?: Record<string, unknown>;
  adapterSelection?: MeetingPlatformAdapterSelection | Record<string, unknown>;
  adapter_selection?: MeetingPlatformAdapterSelection | Record<string, unknown>;
  adapterSelections?: Record<string, MeetingPlatformAdapterSelection | Record<string, unknown>>;
  adapter_selections?: Record<string, MeetingPlatformAdapterSelection | Record<string, unknown>>;
}

export interface MeetingPlatformProviderConnectionDoc {
  label: string;
  url: string;
  note?: string;
}

export interface MeetingPlatformProviderEventMapping {
  provider_event: string;
  normalized_signal: string;
  timeline_role: string;
}

export interface MeetingPlatformProviderRuntimeBindingContract {
  realtime_axis_source?: string;
  realtime_axis_surface?: string;
  annotation_source?: string;
  annotation_timestamp_field: string;
  speaker_track_source?: string;
  provider_reconcile_source?: string;
  provider_role: string;
  provider_reconcile_required_for_production: boolean;
  provider_events_block_realtime: boolean;
  transcript_blocks_realtime: boolean;
  must_read_adapter_selection_before_session: boolean;
  per_meeting_annotation_isolation_required: boolean;
  startup_order?: string[];
}

export interface MeetingPlatformProviderConnectionPack {
  type: 'meeting_platform_provider_connection_pack';
  schema: string;
  schema_version: number;
  platform: string;
  display_name?: string;
  endpoint?: string;
  status_endpoint?: string;
  transport?: string;
  provider_role: string;
  adapter_selection?: Record<string, unknown>;
  runtime_binding_contract: MeetingPlatformProviderRuntimeBindingContract;
  official_docs: MeetingPlatformProviderConnectionDoc[];
  event_mapping: MeetingPlatformProviderEventMapping[];
  subscription?: Record<string, unknown>;
  permissions: Record<string, unknown>;
  security: Record<string, unknown>;
  realtime_annotation_policy: Record<string, unknown>;
  setup_steps: string[];
  readiness: Record<string, unknown>;
  integration: Record<string, unknown>;
  commands: Record<string, string>;
  next_actions: string[];
}

export interface MeetingPlatformProviderConnectionMatrix {
  type: 'meeting_platform_provider_connection_matrix';
  schema: string;
  schema_version: number;
  platform_count: number;
  ready_count: number;
  blocked_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  packs: MeetingPlatformProviderConnectionPack[];
  next_actions: string[];
}

export function buildMeetingPlatformProviderConnectionPack(
  platform: string,
  options?: MeetingPlatformProviderConnectionOptions,
): MeetingPlatformProviderConnectionPack;

export function buildMeetingPlatformProviderConnectionMatrix(
  options?: MeetingPlatformProviderConnectionOptions,
): MeetingPlatformProviderConnectionMatrix;
