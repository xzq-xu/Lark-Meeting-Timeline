import type {
  MeetingPlatformAdapterDecision,
  MeetingPlatformAdapterDecisionInput,
  MeetingPlatformAdapterDecisionOptions,
} from './platform-adapter-decision.mjs';
import type {
  MeetingPlatformRuntimeBundle,
  MeetingPlatformRuntimeBundleOptions,
} from './platform-runtime-bundle.mjs';
import type {
  MeetingPlatformRuntimeProfile,
} from './platform-runtime-profile.mjs';

export const MEETING_PLATFORM_ADAPTER_STARTUP_PLAN_SCHEMA: 'meeting_platform_adapter_startup_plan';
export const MEETING_PLATFORM_ADAPTER_STARTUP_PLAN_MATRIX_SCHEMA: 'meeting_platform_adapter_startup_plan_matrix';
export const MEETING_PLATFORM_ADAPTER_STARTUP_PLAN_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdapterStartupPlanOptions extends MeetingPlatformAdapterDecisionOptions, MeetingPlatformRuntimeBundleOptions {
  startupId?: string;
  startup_id?: string;
  includeReports?: boolean;
  include_reports?: boolean;
  [key: string]: unknown;
}

export interface MeetingPlatformAdapterStartupPlan {
  type: 'meeting_platform_adapter_startup_plan';
  schema: 'meeting_platform_adapter_startup_plan';
  schema_version: 1;
  id?: string;
  accepted: boolean;
  realtime_startup_ready: boolean;
  status: string;
  platform?: string;
  display_name?: string;
  input?: Record<string, unknown>;
  selected_surface?: string;
  install_target?: string;
  decision: MeetingPlatformAdapterDecision;
  runtime_contract?: Record<string, unknown>;
  browser?: {
    matches?: string[];
    host_permissions?: string[];
    content_script?: Record<string, unknown>;
    manifest?: Record<string, unknown>;
  };
  bridge?: Record<string, unknown>;
  runtime?: Record<string, unknown>;
  axis?: Record<string, unknown>;
  actions?: Array<Record<string, unknown>>;
  message_contract?: Record<string, unknown>;
  code_refs?: Record<string, unknown>;
  provider_reconcile?: Record<string, unknown>;
  reports?: {
    bundle?: MeetingPlatformRuntimeBundle;
    profile?: MeetingPlatformRuntimeProfile;
  };
  issues?: Array<Record<string, unknown>>;
  next_actions: string[];
}

export interface MeetingPlatformAdapterStartupPlanMatrix {
  type: 'meeting_platform_adapter_startup_plan_matrix';
  schema: 'meeting_platform_adapter_startup_plan_matrix';
  schema_version: 1;
  platform_count: number;
  accepted_count: number;
  realtime_startup_ready_count: number;
  browser_surface_count: number;
  native_surface_count: number;
  provider_reconcile_surface_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  plans: MeetingPlatformAdapterStartupPlan[];
  next_actions: string[];
}

export function buildMeetingPlatformAdapterStartupPlan(
  input?: string | URL | MeetingPlatformAdapterDecisionInput,
  options?: MeetingPlatformAdapterStartupPlanOptions,
): MeetingPlatformAdapterStartupPlan;

export function buildMeetingPlatformAdapterStartupPlanMatrix(
  input?: MeetingPlatformAdapterDecisionInput,
  options?: MeetingPlatformAdapterStartupPlanOptions,
): MeetingPlatformAdapterStartupPlanMatrix;

export function assertMeetingPlatformAdapterStartupPlan(
  input?: string | URL | MeetingPlatformAdapterDecisionInput,
  options?: MeetingPlatformAdapterStartupPlanOptions,
): MeetingPlatformAdapterStartupPlan;

export function assertMeetingPlatformAdapterStartupPlanMatrix(
  input?: MeetingPlatformAdapterDecisionInput,
  options?: MeetingPlatformAdapterStartupPlanOptions,
): MeetingPlatformAdapterStartupPlanMatrix;
