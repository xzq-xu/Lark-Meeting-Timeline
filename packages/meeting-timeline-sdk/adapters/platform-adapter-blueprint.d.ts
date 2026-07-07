import type { MeetingPlatformRuntimeProfileOptions } from './platform-runtime-profile.mjs';

export const MEETING_PLATFORM_ADAPTER_BLUEPRINT_SCHEMA: 'meeting_platform_adapter_blueprint';
export const MEETING_PLATFORM_ADAPTER_BLUEPRINT_MATRIX_SCHEMA: 'meeting_platform_adapter_blueprint_matrix';
export const MEETING_PLATFORM_ADAPTER_BLUEPRINT_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdapterBlueprintOptions extends MeetingPlatformRuntimeProfileOptions {
  platforms?: string[];
  platform_keys?: string[];
  baseUrl?: string;
  base_url?: string;
}

export interface MeetingPlatformAdapterBlueprintReadinessIssue {
  code: string;
  message: string;
  [key: string]: unknown;
}

export interface MeetingPlatformAdapterBlueprintReadiness {
  type: 'meeting_platform_adapter_blueprint_readiness';
  platform?: string;
  ready: boolean;
  missing: string[];
  issues: MeetingPlatformAdapterBlueprintReadinessIssue[];
}

export interface MeetingPlatformAdapterSurfaceBlueprint {
  surface: string;
  recommended?: boolean;
  priority?: number;
  role?: string;
  evidence?: Record<string, unknown>;
  required_for_realtime?: boolean;
  blocks_realtime?: boolean;
  [key: string]: unknown;
}

export interface MeetingPlatformAdapterBlueprint {
  type: 'meeting_platform_adapter_blueprint';
  schema: 'meeting_platform_adapter_blueprint';
  schema_version: 1;
  platform: string;
  display_name?: string;
  objective: string;
  recommended_mode: string;
  primary_surface?: string;
  surface_order: string[];
  surfaces: {
    host_detector?: MeetingPlatformAdapterSurfaceBlueprint;
    browser_extension?: MeetingPlatformAdapterSurfaceBlueprint;
    native_detector?: MeetingPlatformAdapterSurfaceBlueprint;
    provider_reconcile?: MeetingPlatformAdapterSurfaceBlueprint;
    post_meeting_artifact?: MeetingPlatformAdapterSurfaceBlueprint;
  };
  realtime_axis_contract: Record<string, unknown>;
  annotation_contract: Record<string, unknown>;
  speaker_marker_contract?: Record<string, unknown>;
  runtime_contract: {
    annotations_use_absolute_captured_at_ms: boolean;
    provider_events_block_realtime?: boolean;
    transcript_blocks_realtime?: boolean;
    local_observer_first?: boolean;
    provider_reconcile_after_local_axis?: boolean;
  };
  adapter_route_readiness: Record<string, unknown>;
  platform_notes?: Record<string, unknown>;
  implementation_sequence: string[];
  acceptance_gates: Record<string, string[]>;
  risks: string[];
  next_actions: string[];
  readiness: MeetingPlatformAdapterBlueprintReadiness;
}

export interface MeetingPlatformAdapterBlueprintMatrix {
  type: 'meeting_platform_adapter_blueprint_matrix';
  schema: 'meeting_platform_adapter_blueprint_matrix';
  schema_version: 1;
  platform_count: number;
  ready_count: number;
  browser_primary_count: number;
  native_primary_count: number;
  provider_non_blocking_count: number;
  transcript_non_blocking_count: number;
  platforms: string[];
  rows: Record<string, unknown>[];
  blueprints: MeetingPlatformAdapterBlueprint[];
}

export function buildMeetingPlatformAdapterBlueprint(
  platformOrOptions?: string | Record<string, unknown>,
  options?: MeetingPlatformAdapterBlueprintOptions,
): MeetingPlatformAdapterBlueprint;

export function verifyMeetingPlatformAdapterBlueprint(
  blueprintOrPlatform?: string | MeetingPlatformAdapterBlueprint | Record<string, unknown>,
  options?: MeetingPlatformAdapterBlueprintOptions,
): MeetingPlatformAdapterBlueprintReadiness;

export function assertMeetingPlatformAdapterBlueprint(
  blueprintOrPlatform?: string | MeetingPlatformAdapterBlueprint | Record<string, unknown>,
  options?: MeetingPlatformAdapterBlueprintOptions,
): MeetingPlatformAdapterBlueprint;

export function buildMeetingPlatformAdapterBlueprintMatrix(
  options?: MeetingPlatformAdapterBlueprintOptions,
): MeetingPlatformAdapterBlueprintMatrix;

export function assertMeetingPlatformAdapterBlueprintMatrix(
  options?: MeetingPlatformAdapterBlueprintOptions | MeetingPlatformAdapterBlueprintMatrix,
): MeetingPlatformAdapterBlueprintMatrix;
