export const MEETING_PLATFORM_HANDOFF_READINESS_SCHEMA: 'meeting_platform_handoff_readiness';
export const MEETING_PLATFORM_HANDOFF_READINESS_MATRIX_SCHEMA: 'meeting_platform_handoff_readiness_matrix';
export const MEETING_PLATFORM_HANDOFF_READINESS_SCHEMA_VERSION: number;

export interface MeetingPlatformHandoffReadinessOptions {
  platform?: string;
  provider?: string;
  key?: string;
  name?: string;
  platforms?: string[] | string;
  platform_keys?: string[] | string;
  baseUrl?: string;
  base_url?: string;
  env?: Record<string, unknown>;
  target?: 'pilot' | 'production' | string;
  readinessTarget?: 'pilot' | 'production' | string;
  readiness_target?: 'pilot' | 'production' | string;
  contractTarget?: 'contract' | 'pilot' | 'production' | string;
  contract_target?: 'contract' | 'pilot' | 'production' | string;
  requireProductionReady?: boolean;
  require_production_ready?: boolean;
  evidencePackage?: Record<string, unknown> | Record<string, unknown>[] | Record<string, Record<string, unknown>>;
  evidence_package?: Record<string, unknown> | Record<string, unknown>[] | Record<string, Record<string, unknown>>;
  package?: Record<string, unknown> | Record<string, unknown>[] | Record<string, Record<string, unknown>>;
  handoffPackage?: Record<string, unknown> | Record<string, unknown>[] | Record<string, Record<string, unknown>>;
  handoff_package?: Record<string, unknown> | Record<string, unknown>[] | Record<string, Record<string, unknown>>;
  providerRecords?: unknown[];
  provider_records?: unknown[];
  providerCaptureRecords?: unknown[];
  provider_capture_records?: unknown[];
  meetingAppRecords?: unknown[];
  meeting_app_records?: unknown[];
  meetingAppSnapshotRecords?: unknown[];
  meeting_app_snapshot_records?: unknown[];
  meetingAppRecordSet?: Record<string, unknown>;
  meeting_app_record_set?: Record<string, unknown>;
  recordSet?: Record<string, unknown>;
  record_set?: Record<string, unknown>;
  snapshots?: unknown[] | Record<string, unknown>;
  domSnapshots?: unknown[] | Record<string, unknown>;
  dom_snapshots?: unknown[] | Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingPlatformHandoffReadiness {
  type: 'meeting_platform_handoff_readiness';
  schema: 'meeting_platform_handoff_readiness';
  schema_version: number;
  platform: string;
  display_name: string;
  status: string;
  handoff_ready: boolean;
  pilot_ready: boolean;
  production_ready: boolean;
  provider_reconcile_ready: boolean;
  local_observer_ready: boolean;
  adapter_contract_accepted: boolean;
  real_intake_accepted: boolean;
  evidence_counts: Record<string, number>;
  missing: Record<string, unknown>;
  commands: Record<string, string | undefined>;
  required_host_contract: Record<string, unknown>;
  sdk_methods: string[];
  reports: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingPlatformHandoffReadinessMatrix {
  type: 'meeting_platform_handoff_readiness_matrix';
  schema: 'meeting_platform_handoff_readiness_matrix';
  schema_version: number;
  platform_count: number;
  handoff_ready_count: number;
  pilot_ready_count: number;
  production_ready_count: number;
  local_observer_ready_count: number;
  provider_reconcile_ready_count: number;
  provider_setup_needed_count: number;
  local_evidence_needed_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  reports: MeetingPlatformHandoffReadiness[];
  next_actions: string[];
}

export function buildMeetingPlatformHandoffReadiness(
  platformOrInput?: string | MeetingPlatformHandoffReadinessOptions,
  input?: MeetingPlatformHandoffReadinessOptions,
  options?: MeetingPlatformHandoffReadinessOptions,
): MeetingPlatformHandoffReadiness;

export function buildMeetingPlatformHandoffReadinessMatrix(
  input?: MeetingPlatformHandoffReadinessOptions,
  options?: MeetingPlatformHandoffReadinessOptions,
): MeetingPlatformHandoffReadinessMatrix;

export function assertMeetingPlatformHandoffReadiness(
  platformOrInput?: string | MeetingPlatformHandoffReadinessOptions,
  input?: MeetingPlatformHandoffReadinessOptions,
  options?: MeetingPlatformHandoffReadinessOptions,
): MeetingPlatformHandoffReadiness;

export function assertMeetingPlatformHandoffReadinessMatrix(
  input?: MeetingPlatformHandoffReadinessOptions,
  options?: MeetingPlatformHandoffReadinessOptions,
): MeetingPlatformHandoffReadinessMatrix;
