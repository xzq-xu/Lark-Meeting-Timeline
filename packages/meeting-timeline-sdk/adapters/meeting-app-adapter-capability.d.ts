import type {
  MeetingAppAdapterFitMatrix,
  MeetingAppAdapterFitReport,
  MeetingAppSnapshot,
} from './meeting-apps.mjs';
import type {
  MeetingAppAdapterHandoffPackage,
  MeetingAppAdapterVerificationReport,
} from './meeting-app-adapter-handoff-package.mjs';

export const MEETING_APP_ADAPTER_CAPABILITY_REPORT_SCHEMA: 'meeting_app_adapter_capability_report';
export const MEETING_APP_ADAPTER_CAPABILITY_MATRIX_SCHEMA: 'meeting_app_adapter_capability_matrix';
export const MEETING_APP_ADAPTER_EXECUTION_PLAN_SCHEMA: 'meeting_app_adapter_execution_plan';
export const MEETING_APP_ADAPTER_EXECUTION_PLAN_MATRIX_SCHEMA: 'meeting_app_adapter_execution_plan_matrix';
export const MEETING_APP_ADAPTER_CAPABILITY_SCHEMA_VERSION: 1;

export interface MeetingAppAdapterCapabilityOptions {
  platform?: string;
  provider?: string;
  adapter_key?: string;
  adapterKey?: string;
  platforms?: string[];
  platform_keys?: string[];
  platformKeys?: string[];
  input?: MeetingAppSnapshot | MeetingAppSnapshot[] | Record<string, unknown>;
  snapshot?: MeetingAppSnapshot | MeetingAppSnapshot[] | Record<string, unknown>;
  sample?: MeetingAppSnapshot | MeetingAppSnapshot[] | Record<string, unknown>;
  inputs?: Record<string, MeetingAppSnapshot | MeetingAppSnapshot[] | Record<string, unknown>>;
  inputByPlatform?: Record<string, MeetingAppSnapshot | MeetingAppSnapshot[] | Record<string, unknown>>;
  input_by_platform?: Record<string, MeetingAppSnapshot | MeetingAppSnapshot[] | Record<string, unknown>>;
  snapshots?: Record<string, MeetingAppSnapshot | MeetingAppSnapshot[] | Record<string, unknown>>;
  evidence?: Record<string, unknown>;
  evidence_package?: Record<string, unknown>;
  evidencePackage?: Record<string, unknown>;
  evidenceByAdapter?: Record<string, Record<string, unknown>>;
  evidence_by_adapter?: Record<string, Record<string, unknown>>;
  evidenceByPlatform?: Record<string, Record<string, unknown>>;
  evidence_by_platform?: Record<string, Record<string, unknown>>;
  includeVerification?: boolean;
  include_verification?: boolean;
  includeHandoffPackage?: boolean;
  include_handoff_package?: boolean;
  verify?: boolean;
  [key: string]: unknown;
}

export interface MeetingAppAdapterTimelineCapability {
  name: string;
  status?: string;
  provider_status?: string;
  provider_declared?: boolean;
  local_ready?: boolean;
  post_meeting_backfill?: boolean;
  source?: string;
  fallback?: string;
  signal_types?: string[];
}

export interface MeetingAppAdapterCapabilityReport {
  type: 'meeting_app_adapter_capability_report';
  schema: 'meeting_app_adapter_capability_report';
  schema_version: 1;
  platform: string;
  display_name?: string;
  accepted: boolean;
  static_ready: boolean;
  pilot_ready: boolean;
  production_ready: boolean;
  recommended_mode: string;
  risk_level: string;
  timeline_capabilities: {
    realtime_axis: MeetingAppAdapterTimelineCapability;
    speaker_track: MeetingAppAdapterTimelineCapability;
    participant_track: MeetingAppAdapterTimelineCapability;
    annotation_timeline: MeetingAppAdapterTimelineCapability;
    post_meeting_transcript: MeetingAppAdapterTimelineCapability;
    recording: MeetingAppAdapterTimelineCapability;
    subscription_lifecycle: MeetingAppAdapterTimelineCapability;
  };
  evidence_state: {
    live_snapshot_supplied: boolean;
    live_snapshot_accepted?: boolean;
    live_evidence_ready?: boolean;
    verification_target?: string;
    missing_evidence_count?: number;
  };
  fit_report?: MeetingAppAdapterFitReport;
  verification_report?: MeetingAppAdapterVerificationReport;
  handoff_package?: MeetingAppAdapterHandoffPackage;
  next_actions: string[];
}

export interface MeetingAppAdapterCapabilityMatrix {
  type: 'meeting_app_adapter_capability_matrix';
  schema: 'meeting_app_adapter_capability_matrix';
  schema_version: 1;
  platform_count: number;
  accepted_count: number;
  static_ready_count: number;
  pilot_ready_count: number;
  production_ready_count: number;
  local_axis_ready_count: number;
  local_speaker_ready_count: number;
  provider_axis_declared_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  fit_matrix: MeetingAppAdapterFitMatrix;
  reports: MeetingAppAdapterCapabilityReport[];
  next_actions: string[];
}

export interface MeetingAppAdapterExecutionStep {
  id: string;
  role?: string;
  source?: string;
  status?: string;
  required_for_realtime?: boolean;
  required_for_pilot?: boolean;
  required_for_production?: boolean;
  blocks_realtime_if_missing?: boolean;
  sdk_modules?: string[];
  evidence?: string[];
  fallback?: string;
}

export interface MeetingAppAdapterExecutionGate {
  id: string;
  required_for: string[];
  status: string;
  missing_evidence_count?: number;
}

export interface MeetingAppAdapterExecutionPlan {
  type: 'meeting_app_adapter_execution_plan';
  schema: 'meeting_app_adapter_execution_plan';
  schema_version: 1;
  platform: string;
  display_name?: string;
  accepted: boolean;
  production_ready: boolean;
  recommended_mode: string;
  risk_level: string;
  first_blocked_step?: string;
  realtime_ready: boolean;
  steps: MeetingAppAdapterExecutionStep[];
  gates: MeetingAppAdapterExecutionGate[];
  commands: Record<string, string>;
  handoff_requirements: string[];
  next_actions: string[];
  capability_report: MeetingAppAdapterCapabilityReport;
}

export interface MeetingAppAdapterExecutionPlanMatrix {
  type: 'meeting_app_adapter_execution_plan_matrix';
  schema: 'meeting_app_adapter_execution_plan_matrix';
  schema_version: 1;
  platform_count: number;
  accepted_count: number;
  production_ready_count: number;
  realtime_ready_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  plans: MeetingAppAdapterExecutionPlan[];
  next_actions: string[];
}

export function buildMeetingAppAdapterCapabilityReport(
  platformOrOptions?: string | MeetingAppAdapterCapabilityOptions,
  options?: MeetingAppAdapterCapabilityOptions,
): MeetingAppAdapterCapabilityReport;

export function buildMeetingAppAdapterCapabilityMatrix(
  options?: MeetingAppAdapterCapabilityOptions,
): MeetingAppAdapterCapabilityMatrix;

export function buildMeetingAppAdapterExecutionPlan(
  capabilityOrPlatform?: string | MeetingAppAdapterCapabilityReport | MeetingAppAdapterCapabilityOptions,
  options?: MeetingAppAdapterCapabilityOptions,
): MeetingAppAdapterExecutionPlan;

export function buildMeetingAppAdapterExecutionPlanMatrix(
  options?: MeetingAppAdapterCapabilityOptions & { capabilityMatrix?: MeetingAppAdapterCapabilityMatrix },
): MeetingAppAdapterExecutionPlanMatrix;
