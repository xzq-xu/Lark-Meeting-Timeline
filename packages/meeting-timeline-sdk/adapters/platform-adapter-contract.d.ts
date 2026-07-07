import type { MeetingPlatformFieldCaptureOptions } from './platform-field-capture.mjs';

export const MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA: 'meeting_platform_adapter_contract';
export const MEETING_PLATFORM_ADAPTER_CONTRACT_MATRIX_SCHEMA: 'meeting_platform_adapter_contract_matrix';
export const MEETING_PLATFORM_ADAPTER_CONTRACT_ACCEPTANCE_SCHEMA: 'meeting_platform_adapter_contract_acceptance';
export const MEETING_PLATFORM_ADAPTER_CONTRACT_ACCEPTANCE_MATRIX_SCHEMA: 'meeting_platform_adapter_contract_acceptance_matrix';
export const MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA_VERSION: 1;
export const MEETING_PLATFORM_CANDIDATE_OBSERVATION_ENDPOINT: '/api/meeting-platform/observe-candidates';
export const MEETING_PLATFORM_CANDIDATE_OBSERVATION_MESSAGE_TYPE: 'meeting_timeline.observe_candidates';

export interface MeetingPlatformAdapterContractOptions extends MeetingPlatformFieldCaptureOptions {
  platforms?: string[];
  platform_keys?: string[];
}

export interface MeetingPlatformAdapterContract {
  type: 'meeting_platform_adapter_contract';
  schema: 'meeting_platform_adapter_contract';
  schema_version: 1;
  platform: string;
  display_name?: string;
  objective: string;
  mode: string;
  supported_surfaces: {
    local_observer_or_host_detector: boolean;
    browser_observer: boolean;
    candidate_observation: boolean;
    provider_webhook_or_event_subscription: boolean;
    post_meeting_transcript_import: string;
    realtime_transcript_required: boolean;
    primary?: string;
    recommended_order?: string[];
    provider_reconcile_surface?: string | null;
  };
  adapter_surfaces: Record<string, unknown>;
  launch_requirements: Record<string, unknown>;
  evidence_thresholds: Record<string, unknown>;
  fallback_policy: Record<string, unknown>;
  timebase: Record<string, unknown>;
  realtime_axis: Record<string, unknown>;
  annotations: Record<string, unknown>;
  candidate_observation?: Record<string, unknown>;
  local_observer?: Record<string, unknown>;
  provider_observer: Record<string, unknown>;
  transcript: Record<string, unknown>;
  evidence: Record<string, unknown>;
  implementation: Record<string, unknown>;
  readiness: {
    status?: string;
    production_ready: boolean;
    ready_for_realtime_annotations: boolean;
    provider_ready: boolean;
    provider_missing_env: string[];
    missing_items: string[];
    risks: string[];
  };
  next_actions: string[];
}

export interface MeetingPlatformAdapterContractMatrix {
  type: 'meeting_platform_adapter_contract_matrix';
  schema: 'meeting_platform_adapter_contract_matrix';
  schema_version: 1;
  platform_count: number;
  browser_observer_count: number;
  provider_observer_count: number;
  candidate_observer_count: number;
  production_ready_count: number;
  realtime_ready_count: number;
  platforms: string[];
  rows: Record<string, unknown>[];
  contracts: MeetingPlatformAdapterContract[];
  next_actions: string[];
}

export interface MeetingPlatformAdapterContractAcceptanceReport {
  type: 'meeting_platform_adapter_contract_acceptance';
  schema: 'meeting_platform_adapter_contract_acceptance';
  schema_version: 1;
  platform: string;
  display_name?: string;
  target: string;
  accepted: boolean;
  issue_count: number;
  error_count: number;
  warning_count: number;
  issues: Record<string, unknown>[];
  summary: Record<string, unknown>;
  contract: MeetingPlatformAdapterContract;
}

export interface MeetingPlatformAdapterContractAcceptanceMatrix {
  type: 'meeting_platform_adapter_contract_acceptance_matrix';
  schema: 'meeting_platform_adapter_contract_acceptance_matrix';
  schema_version: 1;
  target: string;
  platform_count: number;
  accepted_count: number;
  rejected_count: number;
  issue_count: number;
  error_count: number;
  warning_count: number;
  platforms: string[];
  rows: Record<string, unknown>[];
  reports: MeetingPlatformAdapterContractAcceptanceReport[];
}

export function buildMeetingPlatformAdapterContract(
  platform: string,
  options?: MeetingPlatformAdapterContractOptions,
): MeetingPlatformAdapterContract;

export function buildMeetingPlatformAdapterContractMatrix(
  options?: MeetingPlatformAdapterContractOptions,
): MeetingPlatformAdapterContractMatrix;

export function buildMeetingPlatformAdapterContractAcceptanceReport(
  contractOrPlatform: string | MeetingPlatformAdapterContract,
  options?: MeetingPlatformAdapterContractOptions,
): MeetingPlatformAdapterContractAcceptanceReport;

export function buildMeetingPlatformAdapterContractAcceptanceMatrix(
  options?: MeetingPlatformAdapterContractOptions,
): MeetingPlatformAdapterContractAcceptanceMatrix;

export function assertMeetingPlatformAdapterContract(
  contractOrPlatform: string | MeetingPlatformAdapterContract,
  options?: MeetingPlatformAdapterContractOptions,
): MeetingPlatformAdapterContractAcceptanceReport;
