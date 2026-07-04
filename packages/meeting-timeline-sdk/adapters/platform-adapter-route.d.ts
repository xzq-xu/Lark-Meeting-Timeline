import type { MeetingPlatformRuntimeProfileOptions } from './platform-runtime-profile.mjs';

export const MEETING_PLATFORM_ADAPTER_ROUTE_SCHEMA: 'meeting_platform_adapter_route';
export const MEETING_PLATFORM_ADAPTER_ROUTE_MATRIX_SCHEMA: 'meeting_platform_adapter_route_matrix';
export const MEETING_PLATFORM_ADAPTER_ROUTE_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdapterRouteOptions extends MeetingPlatformRuntimeProfileOptions {
  platforms?: string[];
  platform_keys?: string[];
  baseUrl?: string;
  base_url?: string;
}

export interface MeetingPlatformAdapterRouteStep {
  route: string;
  priority: number;
  role: string;
  required_for_realtime: boolean;
  blocks_realtime_if_missing: boolean;
  [key: string]: unknown;
}

export interface MeetingPlatformAdapterRoute {
  type: 'meeting_platform_adapter_route';
  schema: 'meeting_platform_adapter_route';
  schema_version: 1;
  platform: string;
  display_name?: string;
  objective: string;
  recommended_mode: string;
  route_order: string[];
  route_count: number;
  routes: MeetingPlatformAdapterRouteStep[];
  entrypoints: Record<string, unknown>;
  realtime_invariants: {
    primary_axis_must_be_created_before_provider_reconcile?: boolean;
    annotations_use_absolute_captured_at_ms: boolean;
    provider_events_block_realtime: boolean;
    transcript_blocks_realtime: boolean;
    per_meeting_annotation_isolation_required: boolean;
  };
  gates: Record<string, unknown>;
  implementation_order: string[];
  next_actions: string[];
}

export interface MeetingPlatformAdapterRouteMatrix {
  type: 'meeting_platform_adapter_route_matrix';
  schema: 'meeting_platform_adapter_route_matrix';
  schema_version: 1;
  platform_count: number;
  local_observer_first_count: number;
  provider_non_blocking_count: number;
  transcript_non_blocking_count: number;
  platforms: string[];
  rows: Record<string, unknown>[];
  routes: MeetingPlatformAdapterRoute[];
}

export function buildMeetingPlatformAdapterRoute(
  platform: string,
  options?: MeetingPlatformAdapterRouteOptions,
): MeetingPlatformAdapterRoute;

export function buildMeetingPlatformAdapterRouteMatrix(
  options?: MeetingPlatformAdapterRouteOptions,
): MeetingPlatformAdapterRouteMatrix;
