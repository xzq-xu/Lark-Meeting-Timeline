import type { MeetingPlatformImplementationHandoffOptions } from './platform-implementation-handoff.mjs';

export const MEETING_PLATFORM_ADAPTER_AUTHORING_PLAN_SCHEMA: 'meeting_platform_adapter_authoring_plan';
export const MEETING_PLATFORM_ADAPTER_AUTHORING_MATRIX_SCHEMA: 'meeting_platform_adapter_authoring_matrix';
export const MEETING_PLATFORM_ADAPTER_AUTHORING_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdapterAuthoringOptions extends MeetingPlatformImplementationHandoffOptions {
  displayName?: string;
  display_name?: string;
  name?: string;
  browserMatches?: string[] | Iterable<string>;
  browser_matches?: string[] | Iterable<string>;
  matches?: string[] | Iterable<string>;
  providerPath?: string;
  provider_path?: string;
  providerTransport?: string;
  provider_transport?: string;
  transport?: string;
}

export interface MeetingPlatformAdapterAuthoringPlan {
  type: 'meeting_platform_adapter_authoring_plan';
  schema: 'meeting_platform_adapter_authoring_plan';
  schema_version: 1;
  platform: string;
  display_name: string;
  built_in: boolean;
  objective: string;
  current_sdk_status: string;
  recommended_first_surface: string;
  provider_reconcile: Record<string, unknown>;
  required_contracts: Record<string, unknown>;
  browser_surface: Record<string, unknown>;
  files: Array<Record<string, unknown>>;
  authoring_steps: Array<Record<string, unknown>>;
  normalizer_template?: string;
  commands: Record<string, string>;
  built_in_artifacts?: Record<string, unknown>;
  production_evidence: Record<string, unknown>;
  risks: string[];
  next_actions: string[];
}

export interface MeetingPlatformAdapterAuthoringMatrix {
  type: 'meeting_platform_adapter_authoring_matrix';
  schema: 'meeting_platform_adapter_authoring_matrix';
  schema_version: 1;
  platform_count: number;
  built_in_count: number;
  external_authoring_count: number;
  browser_surface_ready_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  plans: MeetingPlatformAdapterAuthoringPlan[];
  next_actions: string[];
}

export function buildMeetingPlatformAdapterAuthoringPlan(
  platform: string,
  options?: MeetingPlatformAdapterAuthoringOptions,
): MeetingPlatformAdapterAuthoringPlan;

export function buildMeetingPlatformAdapterAuthoringMatrix(
  options?: MeetingPlatformAdapterAuthoringOptions,
): MeetingPlatformAdapterAuthoringMatrix;
