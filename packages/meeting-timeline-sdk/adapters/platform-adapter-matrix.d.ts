export const MEETING_PLATFORM_ADAPTER_MATRIX_ROW_SCHEMA: 'meeting_platform_adapter_matrix_row';
export const MEETING_PLATFORM_ADAPTER_MATRIX_SCHEMA: 'meeting_platform_adapter_matrix';
export const MEETING_PLATFORM_ADAPTER_MATRIX_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdapterMatrixOptions {
  platforms?: string[];
  platform_keys?: string[];
  platformKeys?: string[];
  includeLocalDetector?: boolean;
  include_local_detector?: boolean;
  includeDecision?: boolean;
  include_decision?: boolean;
  includeRuntimeProfile?: boolean;
  include_runtime_profile?: boolean;
  inputs?: Record<string, Record<string, unknown>>;
  inputByPlatform?: Record<string, Record<string, unknown>>;
  input_by_platform?: Record<string, Record<string, unknown>>;
  snapshots?: Record<string, Record<string, unknown>>;
  snapshotByPlatform?: Record<string, Record<string, unknown>>;
  snapshot_by_platform?: Record<string, Record<string, unknown>>;
  [key: string]: unknown;
}

export interface MeetingPlatformAdapterMatrixRow {
  type: 'meeting_platform_adapter_matrix_row';
  schema: 'meeting_platform_adapter_matrix_row';
  schema_version: 1;
  platform: string;
  display_name?: string;
  adapter_kind?: string;
  status?: string;
  realtime_ready?: boolean;
  accepted?: boolean;
  primary_surface?: string;
  recommended_order?: string[];
  fallback_surfaces?: string[];
  runtime?: Record<string, unknown>;
  realtime_axis?: Record<string, unknown>;
  browser_integration?: Record<string, unknown>;
  provider_reconcile?: Record<string, unknown>;
  speaker_positions?: Record<string, unknown>;
  post_meeting?: Record<string, unknown>;
  readiness?: Record<string, unknown>;
  evidence_first?: string[];
  risks?: string[];
  next_actions?: string[];
  decision?: Record<string, unknown>;
  runtime_profile?: Record<string, unknown>;
}

export interface MeetingPlatformAdapterMatrix {
  type: 'meeting_platform_adapter_matrix';
  schema: 'meeting_platform_adapter_matrix';
  schema_version: 1;
  objective: string;
  platform_count: number;
  realtime_ready_count: number;
  accepted_count: number;
  browser_first_count: number;
  native_first_count: number;
  provider_reconcile_required_count: number;
  speaker_position_enabled_count: number;
  post_meeting_transcript_count: number;
  platforms: string[];
  rows: Record<string, unknown>[];
  adapters: MeetingPlatformAdapterMatrixRow[];
  next_actions: string[];
}

export function buildMeetingPlatformAdapterMatrixRow(
  platform: string,
  input?: Record<string, unknown>,
  options?: MeetingPlatformAdapterMatrixOptions,
): MeetingPlatformAdapterMatrixRow;

export function buildMeetingPlatformAdapterMatrix(
  input?: Record<string, unknown>,
  options?: MeetingPlatformAdapterMatrixOptions,
): MeetingPlatformAdapterMatrix;
