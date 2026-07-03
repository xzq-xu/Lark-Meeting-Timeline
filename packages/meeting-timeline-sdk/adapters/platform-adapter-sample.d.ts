export const MEETING_PLATFORM_ADAPTER_SAMPLE_SCHEMA: 'meeting_platform_adapter_sample';
export const MEETING_PLATFORM_ADAPTER_SAMPLE_MATRIX_SCHEMA: 'meeting_platform_adapter_sample_matrix';
export const MEETING_PLATFORM_ADAPTER_SAMPLE_PLAN_SCHEMA: 'meeting_platform_adapter_sample_plan';
export const MEETING_PLATFORM_ADAPTER_SAMPLE_SCHEMA_VERSION: number;
export const MEETING_PLATFORM_ADAPTER_SAMPLE_PLATFORMS: readonly string[];

export interface MeetingPlatformAdapterSampleOptions {
  baseUrl?: string;
  base_url?: string;
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  startMs?: number | string | Date;
  start_ms?: number | string | Date;
  durationMs?: number;
  duration_ms?: number;
  annotationOffsetMs?: number;
  annotation_offset_ms?: number;
  providerSignalTypes?: Iterable<string> | string[];
  provider_signal_types?: Iterable<string> | string[];
  client?: Record<string, unknown>;
  adapter?: Record<string, unknown>;
  liveAdapter?: Record<string, unknown>;
  live_adapter?: Record<string, unknown>;
  env?: Record<string, unknown>;
  annotation?: Record<string, unknown>;
  includeEvidencePackage?: boolean;
  include_evidence_package?: boolean;
  includeTimelineCalls?: boolean;
  include_timeline_calls?: boolean;
  [key: string]: unknown;
}

export interface MeetingPlatformAdapterSamplePlan {
  type: 'meeting_platform_adapter_sample_plan';
  schema: string;
  schema_version: number;
  platform: string;
  base_url: string;
  endpoint: string;
  clock: Record<string, unknown>;
  provider_signal_types: string[];
  local_observer_snapshots: Array<Record<string, unknown>>;
  annotation: Record<string, unknown>;
  flow: string[];
  contract: Record<string, unknown>;
  contract_acceptance: Record<string, unknown>;
}

export interface MeetingPlatformAdapterSample {
  type: 'meeting_platform_adapter_sample';
  schema: string;
  schema_version: number;
  platform: string;
  accepted: boolean;
  base_url: string;
  endpoint: string;
  operation_count: number;
  provider_event_count: number;
  timeline_call_count: number;
  timeline_method_counts: Record<string, number>;
  contract_acceptance: Record<string, unknown>;
  readiness: Record<string, unknown>;
  verification: Record<string, unknown>;
  summary: Record<string, unknown>;
  plan: MeetingPlatformAdapterSamplePlan;
  operations: Record<string, unknown>;
  evidence_package?: Record<string, unknown>;
  timeline_calls?: Array<Record<string, unknown>>;
}

export interface MeetingPlatformAdapterSampleMatrix {
  type: 'meeting_platform_adapter_sample_matrix';
  schema: string;
  schema_version: number;
  platform_count: number;
  accepted_count: number;
  rejected_count: number;
  production_ready_count: number;
  realtime_ready_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  samples: MeetingPlatformAdapterSample[];
}

export function buildMeetingPlatformAdapterSamplePlan(
  platform: string,
  options?: MeetingPlatformAdapterSampleOptions,
): MeetingPlatformAdapterSamplePlan;

export function runMeetingPlatformAdapterSample(
  platform: string,
  options?: MeetingPlatformAdapterSampleOptions,
): Promise<MeetingPlatformAdapterSample>;

export function runMeetingPlatformAdapterSampleMatrix(
  options?: MeetingPlatformAdapterSampleOptions,
): Promise<MeetingPlatformAdapterSampleMatrix>;

export function assertMeetingPlatformAdapterSample(
  sampleOrPlatform: string | MeetingPlatformAdapterSample,
  options?: MeetingPlatformAdapterSampleOptions,
): Promise<MeetingPlatformAdapterSample>;

export function assertMeetingPlatformAdapterSampleMatrix(
  options?: MeetingPlatformAdapterSampleOptions,
): Promise<MeetingPlatformAdapterSampleMatrix>;
