import type {
  MeetingAppAdapterCapabilityReport,
  MeetingAppAdapterExecutionPlan,
} from './meeting-app-adapter-capability.mjs';
import type {
  MeetingPlatformAdapterRoute,
} from './platform-adapter-route.mjs';
import type {
  MeetingPlatformAdapterBlueprint,
} from './platform-adapter-blueprint.mjs';

export const MEETING_PLATFORM_ADAPTER_DECISION_SCHEMA: 'meeting_platform_adapter_decision';
export const MEETING_PLATFORM_ADAPTER_DECISION_MATRIX_SCHEMA: 'meeting_platform_adapter_decision_matrix';
export const MEETING_PLATFORM_ADAPTER_DECISION_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdapterDecisionOptions {
  platform?: string;
  provider?: string;
  adapter_key?: string;
  adapterKey?: string;
  defaultPlatform?: string;
  default_platform?: string;
  fallbackPlatform?: string;
  fallback_platform?: string;
  surface?: string;
  adapterSurface?: string;
  adapter_surface?: string;
  providerOnly?: boolean;
  provider_only?: boolean;
  platforms?: string[];
  platform_keys?: string[];
  input?: Record<string, unknown>;
  snapshot?: Record<string, unknown>;
  sample?: Record<string, unknown>;
  inputs?: Record<string, Record<string, unknown>>;
  inputByPlatform?: Record<string, Record<string, unknown>>;
  input_by_platform?: Record<string, Record<string, unknown>>;
  snapshots?: Record<string, Record<string, unknown>>;
  snapshotByPlatform?: Record<string, Record<string, unknown>>;
  snapshot_by_platform?: Record<string, Record<string, unknown>>;
  includeReports?: boolean;
  include_reports?: boolean;
  [key: string]: unknown;
}

export interface MeetingPlatformAdapterDecisionInput {
  platform?: string;
  provider?: string;
  adapter_key?: string;
  adapterKey?: string;
  url?: string;
  title?: string;
  meeting_url?: string;
  meetingUrl?: string;
  join_url?: string;
  joinUrl?: string;
  surface?: string;
  adapterSurface?: string;
  adapter_surface?: string;
  tabs?: Array<Record<string, unknown>>;
  window?: Record<string, unknown>;
  browser?: Record<string, unknown>;
  tab?: Record<string, unknown>;
  app?: Record<string, unknown>;
  process?: Record<string, unknown>;
  native?: Record<string, unknown>;
  native_detector?: Record<string, unknown>;
  provider_event?: Record<string, unknown>;
  providerEvent?: Record<string, unknown>;
  provider_events?: Record<string, unknown>[];
  providerEvents?: Record<string, unknown>[];
  platforms?: string[];
  platform_keys?: string[];
  input?: Record<string, unknown>;
  snapshot?: Record<string, unknown>;
  sample?: Record<string, unknown>;
  inputs?: Record<string, Record<string, unknown>>;
  inputByPlatform?: Record<string, Record<string, unknown>>;
  input_by_platform?: Record<string, Record<string, unknown>>;
  snapshots?: Record<string, Record<string, unknown>>;
  snapshotByPlatform?: Record<string, Record<string, unknown>>;
  snapshot_by_platform?: Record<string, Record<string, unknown>>;
  [key: string]: unknown;
}

export interface MeetingPlatformAdapterDecision {
  type: 'meeting_platform_adapter_decision';
  schema: 'meeting_platform_adapter_decision';
  schema_version: 1;
  accepted: boolean;
  realtime_ready: boolean;
  status: string;
  platform?: string;
  display_name?: string;
  detected?: Record<string, unknown>;
  platform_source?: string;
  selected_surface?: string;
  surface_source?: string;
  selected_route?: string;
  recommended_mode?: string;
  adapter_blueprint?: Record<string, unknown>;
  first_blocked_step?: string;
  contracts?: Record<string, unknown>;
  timeline_capabilities?: MeetingAppAdapterCapabilityReport['timeline_capabilities'];
  route_summary?: Record<string, unknown>;
  runtime_actions?: Array<Record<string, unknown>>;
  evidence_requirements?: Record<string, unknown>;
  reports?: {
    route?: MeetingPlatformAdapterRoute;
    route_readiness?: Record<string, unknown>;
    adapter_blueprint?: MeetingPlatformAdapterBlueprint;
    capability?: MeetingAppAdapterCapabilityReport;
    execution_plan?: MeetingAppAdapterExecutionPlan;
  };
  issues?: Array<Record<string, unknown>>;
  next_actions: string[];
}

export interface MeetingPlatformAdapterDecisionMatrix {
  type: 'meeting_platform_adapter_decision_matrix';
  schema: 'meeting_platform_adapter_decision_matrix';
  schema_version: 1;
  platform_count: number;
  accepted_count: number;
  realtime_ready_count: number;
  adapter_blueprint_ready_count: number;
  browser_surface_count: number;
  native_surface_count: number;
  provider_reconcile_surface_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  decisions: MeetingPlatformAdapterDecision[];
  next_actions: string[];
}

export function buildMeetingPlatformAdapterDecision(
  input?: string | URL | MeetingPlatformAdapterDecisionInput,
  options?: MeetingPlatformAdapterDecisionOptions,
): MeetingPlatformAdapterDecision;

export function buildMeetingPlatformAdapterDecisionMatrix(
  input?: MeetingPlatformAdapterDecisionInput,
  options?: MeetingPlatformAdapterDecisionOptions,
): MeetingPlatformAdapterDecisionMatrix;

export function assertMeetingPlatformAdapterDecision(
  input?: string | URL | MeetingPlatformAdapterDecisionInput,
  options?: MeetingPlatformAdapterDecisionOptions,
): MeetingPlatformAdapterDecision;

export function assertMeetingPlatformAdapterDecisionMatrix(
  input?: MeetingPlatformAdapterDecisionInput,
  options?: MeetingPlatformAdapterDecisionOptions,
): MeetingPlatformAdapterDecisionMatrix;
