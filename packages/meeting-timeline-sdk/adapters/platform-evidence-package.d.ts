import type { MeetingAppSnapshotRecord, MeetingAppSnapshotRecordOptions, MeetingAppSnapshotRecordSet } from './meeting-app-snapshot-recorder.mjs';
import type { PlatformCaptureOptions, PlatformCaptureRecord } from './platform-capture.mjs';
import type { MeetingPlatformEvidenceCorrelation } from './platform-evidence-correlation.mjs';
import type { MeetingPlatformAdaptationRunbook, MeetingPlatformRolloutOptions, MeetingPlatformRolloutPlan } from './platform-rollout.mjs';

export const MEETING_PLATFORM_EVIDENCE_PACKAGE_SCHEMA: string;
export const MEETING_PLATFORM_EVIDENCE_PACKAGE_SCHEMA_VERSION: number;

export interface MeetingPlatformEvidencePackageOptions extends MeetingPlatformRolloutOptions, MeetingAppSnapshotRecordOptions {
  id?: string;
  label?: string;
  source?: string;
  createdAtMs?: number | string | Date;
  created_at_ms?: number | string | Date;
  baseUrl?: string;
  base_url?: string;
  env?: Record<string, unknown>;
  providerRecords?: PlatformCaptureRecord[] | Record<string, unknown>[];
  provider_records?: PlatformCaptureRecord[] | Record<string, unknown>[];
  providerCaptureRecords?: PlatformCaptureRecord[] | Record<string, unknown>[];
  provider_capture_records?: PlatformCaptureRecord[] | Record<string, unknown>[];
  providerSamples?: unknown[] | Record<string, unknown[]>;
  provider_samples?: unknown[] | Record<string, unknown[]>;
  meetingAppRecordSet?: MeetingAppSnapshotRecordSet | Record<string, unknown>;
  meeting_app_record_set?: MeetingAppSnapshotRecordSet | Record<string, unknown>;
  meetingAppRecords?: MeetingAppSnapshotRecord[] | Record<string, unknown>[];
  meeting_app_records?: MeetingAppSnapshotRecord[] | Record<string, unknown>[];
  meetingAppSnapshots?: unknown[] | Record<string, unknown[]>;
  meeting_app_snapshots?: unknown[] | Record<string, unknown[]>;
  domSnapshots?: unknown[] | Record<string, unknown[]>;
  dom_snapshots?: unknown[] | Record<string, unknown[]>;
  includeRunbook?: boolean;
  include_runbook?: boolean;
  includeEnvValues?: boolean;
  include_env_values?: boolean;
  requireCorrelation?: boolean;
  require_correlation?: boolean;
  maxClockSkewMs?: number;
  max_clock_skew_ms?: number;
  notes?: string;
  metadata?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingPlatformEvidencePackage {
  schema: string;
  schema_version: number;
  id: string;
  platform: string;
  display_name?: string;
  createdAtMs: number;
  created_at_ms: number;
  source: string;
  label?: string;
  base_url?: string;
  evidence_summary: Record<string, unknown>;
  provider_records: PlatformCaptureRecord[] | Record<string, unknown>[];
  provider_samples?: Record<string, unknown[]>;
  meeting_app_record_set?: MeetingAppSnapshotRecordSet | Record<string, unknown>;
  evidence_correlation?: MeetingPlatformEvidenceCorrelation;
  rollout_plan: MeetingPlatformRolloutPlan;
  runbook?: MeetingPlatformAdaptationRunbook;
  handoff: Record<string, unknown>;
  env_summary?: {
    configured_keys: string[];
    configured: Record<string, boolean>;
    values?: Record<string, unknown>;
  };
  notes?: string;
  metadata?: Record<string, unknown>;
}

export interface MeetingPlatformEvidencePackageSummary {
  type: 'meeting_platform_evidence_package_summary';
  package_id: string;
  platform: string;
  display_name?: string;
  status: string;
  recommended_mode?: string;
  production_ready: boolean;
  ready_for_realtime_annotations: boolean;
  provider_record_count: number;
  provider_sample_count: number;
  meeting_app_record_count: number;
  correlation_status?: string;
  correlation_passed?: boolean;
  correlation_confidence?: string;
  provider_missing_required_coverage: string[];
  local_dom_missing_required_coverage: string[];
  next_actions: string[];
}

export interface MeetingPlatformEvidencePackageVerification {
  type: 'meeting_platform_evidence_package_verification';
  package_id: string;
  platform: string;
  passed: boolean;
  requirement: 'production_ready' | 'ready_for_realtime_annotations';
  status: string;
  production_ready: boolean;
  ready_for_realtime_annotations: boolean;
  embedded_plan_matches?: boolean;
  embedded_status?: Record<string, unknown>;
  verified_status: Record<string, unknown>;
  provider_record_count: number;
  provider_sample_count: number;
  meeting_app_record_count: number;
  correlation_required: boolean;
  correlation_passed: boolean;
  correlation_status?: string;
  correlation_confidence?: string;
  evidence_correlation?: MeetingPlatformEvidenceCorrelation;
  provider_missing_required_coverage: string[];
  local_dom_missing_required_coverage: string[];
  next_actions: string[];
  verified_package?: MeetingPlatformEvidencePackage;
}

export interface MeetingPlatformEvidencePackageBuilder {
  platform: string;
  addProviderRecord(record?: PlatformCaptureRecord | Record<string, unknown>): PlatformCaptureRecord | Record<string, unknown>;
  captureProviderWebhook(input?: Record<string, unknown> | unknown, payload?: unknown, captureOptions?: PlatformCaptureOptions): PlatformCaptureRecord;
  addProviderSample(sample?: Record<string, unknown>): Record<string, unknown>;
  addMeetingAppRecord(input?: MeetingAppSnapshotRecord | Record<string, unknown>, recordOptions?: MeetingAppSnapshotRecordOptions): MeetingAppSnapshotRecord;
  captureMeetingAppSnapshot(input?: MeetingAppSnapshotRecord | Record<string, unknown>, recordOptions?: MeetingAppSnapshotRecordOptions): MeetingAppSnapshotRecord;
  exportPackage(packageOptions?: MeetingPlatformEvidencePackageOptions): MeetingPlatformEvidencePackage;
  summary(summaryOptions?: MeetingPlatformEvidencePackageOptions): MeetingPlatformEvidencePackageSummary;
  reset(): { removed: Record<string, number> };
  getState(): Record<string, unknown>;
  providerRecords(): (PlatformCaptureRecord | Record<string, unknown>)[];
  providerSamples(): Record<string, unknown>[];
  meetingAppRecords(): MeetingAppSnapshotRecord[];
}

export function buildMeetingPlatformEvidencePackage(
  platformOrInput: string | MeetingPlatformEvidencePackageOptions,
  inputOrOptions?: MeetingPlatformEvidencePackageOptions,
  maybeOptions?: MeetingPlatformEvidencePackageOptions,
): MeetingPlatformEvidencePackage;

export function buildMeetingPlatformEvidencePackageSummary(
  packageOrInput: MeetingPlatformEvidencePackage | string | MeetingPlatformEvidencePackageOptions,
  options?: MeetingPlatformEvidencePackageOptions,
): MeetingPlatformEvidencePackageSummary;

export function verifyMeetingPlatformEvidencePackage(
  packageOrInput: MeetingPlatformEvidencePackage | string | MeetingPlatformEvidencePackageOptions,
  options?: MeetingPlatformEvidencePackageOptions,
): MeetingPlatformEvidencePackageVerification;

export function createMeetingPlatformEvidencePackageBuilder(
  platform: string,
  options?: MeetingPlatformEvidencePackageOptions,
): MeetingPlatformEvidencePackageBuilder;
