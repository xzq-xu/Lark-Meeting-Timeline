import type { MeetingPlatformFieldCaptureOptions } from './platform-field-capture.mjs';

export const MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA: 'meeting_platform_adapter_contract';
export const MEETING_PLATFORM_ADAPTER_CONTRACT_MATRIX_SCHEMA: 'meeting_platform_adapter_contract_matrix';
export const MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA_VERSION: 1;

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
    provider_webhook_or_event_subscription: boolean;
    post_meeting_transcript_import: string;
    realtime_transcript_required: boolean;
  };
  timebase: Record<string, unknown>;
  realtime_axis: Record<string, unknown>;
  annotations: Record<string, unknown>;
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
  production_ready_count: number;
  realtime_ready_count: number;
  platforms: string[];
  rows: Record<string, unknown>[];
  contracts: MeetingPlatformAdapterContract[];
  next_actions: string[];
}

export function buildMeetingPlatformAdapterContract(
  platform: string,
  options?: MeetingPlatformAdapterContractOptions,
): MeetingPlatformAdapterContract;

export function buildMeetingPlatformAdapterContractMatrix(
  options?: MeetingPlatformAdapterContractOptions,
): MeetingPlatformAdapterContractMatrix;
