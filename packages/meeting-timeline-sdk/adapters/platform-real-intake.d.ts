export const MEETING_PLATFORM_REAL_INTAKE_SCHEMA: 'meeting_platform_real_evidence_intake';
export const MEETING_PLATFORM_REAL_INTAKE_MATRIX_SCHEMA: 'meeting_platform_real_evidence_intake_matrix';
export const MEETING_PLATFORM_REAL_INTAKE_PLAN_SCHEMA: 'meeting_platform_real_evidence_intake_plan';
export const MEETING_PLATFORM_REAL_INTAKE_SCHEMA_VERSION: number;

export interface MeetingPlatformRealEvidenceIntakeOptions {
  baseUrl?: string;
  base_url?: string;
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  providerRecords?: unknown[];
  provider_records?: unknown[];
  providerCaptureRecords?: unknown[];
  provider_capture_records?: unknown[];
  meetingAppRecords?: unknown[];
  meeting_app_records?: unknown[];
  meetingAppSnapshotRecords?: unknown[];
  meeting_app_snapshot_records?: unknown[];
  meetingAppSnapshots?: unknown[];
  meeting_app_snapshots?: unknown[];
  domSnapshots?: unknown[];
  dom_snapshots?: unknown[];
  evidencePackage?: Record<string, unknown>;
  evidence_package?: Record<string, unknown>;
  allowFixtureEvidence?: boolean;
  allow_fixture_evidence?: boolean;
  requireProductionReady?: boolean;
  require_production_ready?: boolean;
  includeEvidencePackage?: boolean;
  include_evidence_package?: boolean;
  [key: string]: unknown;
}

export interface MeetingPlatformRealEvidenceIntakePlan {
  type: 'meeting_platform_real_evidence_intake_plan';
  schema: string;
  schema_version: number;
  platform: string;
  base_url: string;
  provider_endpoint: string;
  required_inputs: string[];
  forbidden_inputs: string[];
  output_contract: Record<string, unknown>;
}

export interface MeetingPlatformRealEvidenceIntakeReport {
  type: 'meeting_platform_real_evidence_intake';
  schema: string;
  schema_version: number;
  platform: string;
  accepted: boolean;
  base_url: string;
  provider_record_count: number;
  meeting_app_record_count: number;
  fixture_evidence_count: number;
  blocking_count: number;
  checks: Array<Record<string, unknown>>;
  blocking_checks: Array<Record<string, unknown>>;
  summary: Record<string, unknown>;
  verification: Record<string, unknown>;
  readiness: Record<string, unknown>;
  runtime_host_replay?: Record<string, unknown>;
  runtime_host_replay_accepted?: boolean;
  plan: MeetingPlatformRealEvidenceIntakePlan;
  evidence_package?: Record<string, unknown>;
  issues: Array<Record<string, unknown>>;
  next_actions: string[];
}

export interface MeetingPlatformRealEvidenceIntakeMatrix {
  type: 'meeting_platform_real_evidence_intake_matrix';
  schema: string;
  schema_version: number;
  platform_count: number;
  accepted_count: number;
  rejected_count: number;
  production_ready_count: number;
  realtime_ready_count: number;
  runtime_host_replay_ready_count?: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  reports: MeetingPlatformRealEvidenceIntakeReport[];
}

export function buildMeetingPlatformRealEvidenceIntakePlan(
  platform: string,
  options?: MeetingPlatformRealEvidenceIntakeOptions,
): MeetingPlatformRealEvidenceIntakePlan;

export function buildMeetingPlatformRealEvidenceIntakeReport(
  platform: string,
  input?: Record<string, unknown>,
  options?: MeetingPlatformRealEvidenceIntakeOptions,
): MeetingPlatformRealEvidenceIntakeReport;

export function buildMeetingPlatformRealEvidenceIntakeMatrix(
  input?: Record<string, unknown>,
  options?: MeetingPlatformRealEvidenceIntakeOptions,
): MeetingPlatformRealEvidenceIntakeMatrix;

export function runMeetingPlatformRealEvidenceIntakeReport(
  platform: string,
  input?: Record<string, unknown>,
  options?: MeetingPlatformRealEvidenceIntakeOptions,
): Promise<MeetingPlatformRealEvidenceIntakeReport>;

export function runMeetingPlatformRealEvidenceIntakeMatrix(
  input?: Record<string, unknown>,
  options?: MeetingPlatformRealEvidenceIntakeOptions,
): Promise<MeetingPlatformRealEvidenceIntakeMatrix>;

export function assertMeetingPlatformRealEvidenceIntake(
  platform: string,
  input?: Record<string, unknown>,
  options?: MeetingPlatformRealEvidenceIntakeOptions,
): MeetingPlatformRealEvidenceIntakeReport;

export function assertMeetingPlatformRealEvidenceIntakeMatrix(
  input?: Record<string, unknown>,
  options?: MeetingPlatformRealEvidenceIntakeOptions,
): MeetingPlatformRealEvidenceIntakeMatrix;
