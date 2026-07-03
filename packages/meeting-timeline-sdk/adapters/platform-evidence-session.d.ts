import type { MeetingAppSnapshotRecord, MeetingAppSnapshotRecordOptions, MeetingAppSnapshotRecordSet } from './meeting-app-snapshot-recorder.mjs';
import type { PlatformCaptureOptions, PlatformCaptureRecord } from './platform-capture.mjs';
import type {
  MeetingPlatformEvidencePackage,
  MeetingPlatformEvidencePackageOptions,
  MeetingPlatformEvidencePackageVerification,
} from './platform-evidence-package.mjs';
import type {
  MeetingPlatformEvidenceCorrelation,
  MeetingPlatformEvidenceCorrelationOptions,
} from './platform-evidence-correlation.mjs';
import type { MeetingPlatformAdaptationStrategy } from './platform-strategy.mjs';

export const MEETING_PLATFORM_EVIDENCE_SESSION_SCHEMA: string;
export const MEETING_PLATFORM_EVIDENCE_SESSION_SCHEMA_VERSION: number;

export interface MeetingPlatformEvidenceSessionOptions extends MeetingPlatformEvidencePackageOptions {
  id?: string;
  sessionId?: string;
  session_id?: string;
  packageId?: string;
  package_id?: string;
  meetingAppRecordSetId?: string;
  meeting_app_record_set_id?: string;
  recordSetId?: string;
  record_set_id?: string;
  includePackageSummary?: boolean;
  include_package_summary?: boolean;
  includeVerification?: boolean;
  include_verification?: boolean;
}

export interface MeetingPlatformEvidenceSessionSummary {
  type: 'meeting_platform_evidence_session_summary';
  schema: string;
  schema_version: number;
  session_id: string;
  package_id: string;
  platform: string;
  display_name?: string;
  createdAtMs: number;
  created_at_ms: number;
  status: string;
  recommended_mode?: string;
  production_ready: boolean;
  ready_for_realtime_annotations: boolean;
  can_insert_realtime_marks: boolean;
  local_observer_ready: boolean;
  provider_reconcile_ready: boolean;
  package_ready_for_handoff: boolean;
  verification_passed: boolean;
  verification_requirement: 'production_ready' | 'ready_for_realtime_annotations';
  provider_record_count: number;
  provider_sample_count: number;
  meeting_app_record_count: number;
  correlation_status?: string;
  correlation_passed?: boolean;
  correlation_confidence?: string;
  provider_missing_required_coverage: string[];
  local_dom_missing_required_coverage: string[];
  next_actions: string[];
  handoff: Record<string, unknown>;
  evidence_package_summary?: Record<string, unknown>;
  verification?: MeetingPlatformEvidencePackageVerification;
}

export interface MeetingPlatformEvidenceSessionState {
  type: 'meeting_platform_evidence_session_state';
  schema: string;
  schema_version: number;
  session_id: string;
  platform: string;
  createdAtMs: number;
  created_at_ms: number;
  provider_record_count: number;
  provider_sample_count: number;
  meeting_app_record_count: number;
  latest_provider_record_id?: string;
  latest_meeting_app_record_id?: string;
}

export interface MeetingPlatformEvidenceSession {
  schema: string;
  schema_version: number;
  platform: string;
  id: string;
  session_id: string;
  addProviderRecord(record?: PlatformCaptureRecord | Record<string, unknown>): PlatformCaptureRecord | Record<string, unknown>;
  captureProviderWebhook(input?: Record<string, unknown> | unknown, payload?: unknown, captureOptions?: PlatformCaptureOptions): PlatformCaptureRecord;
  addProviderSample(sample?: Record<string, unknown>): Record<string, unknown>;
  addMeetingAppRecord(input?: MeetingAppSnapshotRecord | Record<string, unknown>, recordOptions?: MeetingAppSnapshotRecordOptions): MeetingAppSnapshotRecord;
  captureMeetingAppSnapshot(input?: MeetingAppSnapshotRecord | Record<string, unknown>, recordOptions?: MeetingAppSnapshotRecordOptions): MeetingAppSnapshotRecord;
  meetingAppRecordSet(recordSetOptions?: MeetingPlatformEvidenceSessionOptions): MeetingAppSnapshotRecordSet | undefined;
  correlation(correlationOptions?: MeetingPlatformEvidenceCorrelationOptions): MeetingPlatformEvidenceCorrelation;
  strategy(strategyOptions?: MeetingPlatformEvidenceSessionOptions): MeetingPlatformAdaptationStrategy;
  exportPackage(packageOptions?: MeetingPlatformEvidenceSessionOptions): MeetingPlatformEvidencePackage;
  verify(verifyOptions?: MeetingPlatformEvidenceSessionOptions): MeetingPlatformEvidencePackageVerification;
  summary(summaryOptions?: MeetingPlatformEvidenceSessionOptions): MeetingPlatformEvidenceSessionSummary;
  getState(): MeetingPlatformEvidenceSessionState;
  reset(): { removed: Record<string, number> };
  providerRecords(): (PlatformCaptureRecord | Record<string, unknown>)[];
  providerSamples(): Record<string, unknown>[];
  meetingAppRecords(): MeetingAppSnapshotRecord[];
}

export function createMeetingPlatformEvidenceSession(
  platform: string,
  options?: MeetingPlatformEvidenceSessionOptions,
): MeetingPlatformEvidenceSession;
