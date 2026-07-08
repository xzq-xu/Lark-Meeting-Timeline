import type { MeetingPlatformAdapterInstallManifest } from './platform-adapter-install-manifest.mjs';

export const MEETING_PLATFORM_ADAPTER_LAUNCH_PLAN_SCHEMA: 'meeting_platform_adapter_launch_plan';
export const MEETING_PLATFORM_ADAPTER_CANDIDATE_LAUNCH_PLAN_SCHEMA: 'meeting_platform_adapter_candidate_launch_plan';
export const MEETING_PLATFORM_ADAPTER_LAUNCH_PLAN_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdapterLaunchPlanOptions {
  installManifest?: MeetingPlatformAdapterInstallManifest;
  install_manifest?: MeetingPlatformAdapterInstallManifest;
  platform?: string;
  platform_key?: string;
  surface?: string;
  preferredSurface?: string;
  preferred_surface?: string;
  url?: string;
  href?: string;
  meetingUrl?: string;
  meeting_url?: string;
  capturedAtMs?: number;
  captured_at_ms?: number;
}

export interface MeetingPlatformAdapterLaunchPlan {
  type: 'meeting_platform_adapter_launch_plan';
  schema: 'meeting_platform_adapter_launch_plan';
  schema_version: 1;
  accepted: boolean;
  platform?: string;
  detection_reason: string;
  selected_surface?: string;
  install_manifest_schema?: string;
  install_manifest_accepted: boolean;
  detected_meeting?: Record<string, unknown>;
  current_url?: string;
  platform_row?: Record<string, unknown>;
  adapter_blueprint?: Record<string, unknown>;
  raw_signal_validation?: Record<string, unknown>;
  surface_entrypoint?: Record<string, unknown>;
  axis_contract: Record<string, unknown>;
  runtime_actions: Array<Record<string, unknown>>;
  mark_template?: Record<string, unknown>;
  readiness: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingPlatformAdapterCandidateLaunchPlan {
  type: 'meeting_platform_adapter_candidate_launch_plan';
  schema: 'meeting_platform_adapter_candidate_launch_plan';
  schema_version: 1;
  accepted: boolean;
  status: string;
  platform?: string;
  selected_platform?: string;
  selected_surface?: string;
  selection_strategy?: string;
  selected_candidate_index?: number;
  selected_candidate_score?: number;
  selected_candidate_rank?: number;
  selected_candidate_reason?: string;
  selected_candidate?: Record<string, unknown>;
  install_manifest_schema?: string;
  install_manifest_accepted: boolean;
  candidate_preflight: import('./platform-adapter-preflight.mjs').MeetingPlatformAdapterCandidatePreflight;
  launch_input: Record<string, unknown>;
  launch_plan: MeetingPlatformAdapterLaunchPlan;
  axis_contract: Record<string, unknown>;
  adapter_blueprint?: Record<string, unknown>;
  raw_signal_validation?: Record<string, unknown>;
  runtime_actions: Array<Record<string, unknown>>;
  mark_template?: Record<string, unknown>;
  readiness: Record<string, unknown>;
  next_actions: string[];
}

export function buildMeetingPlatformAdapterLaunchPlan(
  manifestOrInput?: MeetingPlatformAdapterInstallManifest | { installManifest?: MeetingPlatformAdapterInstallManifest; install_manifest?: MeetingPlatformAdapterInstallManifest; manifest?: MeetingPlatformAdapterInstallManifest } | Record<string, unknown>,
  input?: MeetingPlatformAdapterLaunchPlanOptions | string,
  options?: MeetingPlatformAdapterLaunchPlanOptions,
): MeetingPlatformAdapterLaunchPlan;

export function buildMeetingPlatformAdapterCandidateLaunchPlan(
  manifestOrInput?: MeetingPlatformAdapterInstallManifest | { installManifest?: MeetingPlatformAdapterInstallManifest; install_manifest?: MeetingPlatformAdapterInstallManifest; manifest?: MeetingPlatformAdapterInstallManifest } | Record<string, unknown>,
  input?: Record<string, unknown>,
  options?: MeetingPlatformAdapterLaunchPlanOptions & import('./platform-adapter-preflight.mjs').MeetingPlatformAdapterPreflightOptions,
): MeetingPlatformAdapterCandidateLaunchPlan;

export function assertMeetingPlatformAdapterLaunchPlan(
  planOrInput?: MeetingPlatformAdapterLaunchPlan | MeetingPlatformAdapterInstallManifest | Record<string, unknown>,
  input?: MeetingPlatformAdapterLaunchPlanOptions | string,
  options?: MeetingPlatformAdapterLaunchPlanOptions,
): MeetingPlatformAdapterLaunchPlan;

export function assertMeetingPlatformAdapterCandidateLaunchPlan(
  planOrInput?: MeetingPlatformAdapterCandidateLaunchPlan | MeetingPlatformAdapterInstallManifest | Record<string, unknown>,
  input?: Record<string, unknown>,
  options?: MeetingPlatformAdapterLaunchPlanOptions & import('./platform-adapter-preflight.mjs').MeetingPlatformAdapterPreflightOptions,
): MeetingPlatformAdapterCandidateLaunchPlan;
