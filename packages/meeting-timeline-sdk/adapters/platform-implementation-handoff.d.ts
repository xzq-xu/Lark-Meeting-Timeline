import type { MeetingPlatformRuntimeBundleOptions } from './platform-runtime-bundle.mjs';

export const MEETING_PLATFORM_IMPLEMENTATION_HANDOFF_SCHEMA: 'meeting_platform_implementation_handoff';
export const MEETING_PLATFORM_IMPLEMENTATION_HANDOFF_MATRIX_SCHEMA: 'meeting_platform_implementation_handoff_matrix';
export const MEETING_PLATFORM_IMPLEMENTATION_HANDOFF_SCHEMA_VERSION: 1;

export interface MeetingPlatformImplementationHandoffOptions extends MeetingPlatformRuntimeBundleOptions {
  priorityPlatformOrder?: Iterable<string> | string[];
  priority_platform_order?: Iterable<string> | string[];
}

export interface MeetingPlatformImplementationHandoff {
  type: 'meeting_platform_implementation_handoff';
  schema: 'meeting_platform_implementation_handoff';
  schema_version: 1;
  platform: string;
  display_name?: string;
  objective: string;
  priority_tier?: string;
  rank_hint?: number;
  implementation_ready: boolean;
  pilot_ready: boolean;
  production_ready: boolean;
  recommended_first_surface: string;
  next_phase?: string;
  next_action?: string;
  package_entrypoints: Record<string, string>;
  install_surface: Record<string, unknown>;
  runtime_events: Record<string, unknown>;
  provider_reconcile: Record<string, unknown>;
  adapter_selection: Record<string, unknown>;
  adapter_preflight: Record<string, unknown>;
  contracts: Record<string, unknown>;
  implementation_flow: Array<Record<string, unknown>>;
  acceptance: Record<string, unknown>;
  readiness: Record<string, unknown>;
  adaptation_strategy?: Record<string, unknown>;
  production_gaps: string[];
  next_actions: string[];
}

export interface MeetingPlatformImplementationHandoffMatrix {
  type: 'meeting_platform_implementation_handoff_matrix';
  schema: 'meeting_platform_implementation_handoff_matrix';
  schema_version: 1;
  platform_count: number;
  implementation_ready_count: number;
  pilot_ready_count: number;
  production_ready_count: number;
  adapter_preflight_startup_ready_count: number;
  adapter_preflight_realtime_ready_count: number;
  adapter_selection_ready_count: number;
  recommended_first_platform?: string;
  recommended_first_surface?: string;
  priority_order: string[];
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  handoffs: MeetingPlatformImplementationHandoff[];
  next_actions: string[];
}

export function buildMeetingPlatformImplementationHandoff(
  platform: string,
  options?: MeetingPlatformImplementationHandoffOptions,
): MeetingPlatformImplementationHandoff;

export function buildMeetingPlatformImplementationHandoffMatrix(
  options?: MeetingPlatformImplementationHandoffOptions,
): MeetingPlatformImplementationHandoffMatrix;
