import type { MeetingAppSnapshotRecord, MeetingAppSnapshotRecordSet } from './meeting-app-snapshot-recorder.mjs';
import type { PlatformCaptureRecord } from './platform-capture.mjs';

export const MEETING_PLATFORM_EVIDENCE_CORRELATION_SCHEMA: 'meeting_platform_evidence_correlation';
export const MEETING_PLATFORM_EVIDENCE_CORRELATION_SCHEMA_VERSION: 1;

export interface MeetingPlatformEvidenceCorrelationOptions {
  maxClockSkewMs?: number;
  max_clock_skew_ms?: number;
  [key: string]: unknown;
}

export interface MeetingPlatformEvidenceCorrelationInput extends MeetingPlatformEvidenceCorrelationOptions {
  platform?: string;
  provider?: string;
  adapter?: string;
  providerRecords?: PlatformCaptureRecord[] | Record<string, unknown>[];
  provider_records?: PlatformCaptureRecord[] | Record<string, unknown>[];
  providerCaptureRecords?: PlatformCaptureRecord[] | Record<string, unknown>[];
  provider_capture_records?: PlatformCaptureRecord[] | Record<string, unknown>[];
  captureRecords?: PlatformCaptureRecord[] | Record<string, unknown>[];
  capture_records?: PlatformCaptureRecord[] | Record<string, unknown>[];
  meetingAppRecordSet?: MeetingAppSnapshotRecordSet | Record<string, unknown>;
  meeting_app_record_set?: MeetingAppSnapshotRecordSet | Record<string, unknown>;
  recordSet?: MeetingAppSnapshotRecordSet | Record<string, unknown>;
  record_set?: MeetingAppSnapshotRecordSet | Record<string, unknown>;
  meetingAppRecords?: MeetingAppSnapshotRecord[] | Record<string, unknown>[];
  meeting_app_records?: MeetingAppSnapshotRecord[] | Record<string, unknown>[];
}

export interface MeetingPlatformEvidenceCorrelation {
  schema: 'meeting_platform_evidence_correlation';
  schema_version: 1;
  type: 'meeting_platform_evidence_correlation';
  platform: string;
  status: 'matched' | 'time_matched' | 'time_matched_identity_unconfirmed' | 'single_source' | 'failed';
  passed: boolean;
  confidence: 'high' | 'medium' | 'none';
  provider_record_count: number;
  provider_signal_count: number;
  meeting_app_record_count: number;
  identity_match: Record<string, unknown>;
  time_alignment: Record<string, unknown>;
  coverage: Record<string, boolean>;
  issues: Record<string, unknown>[];
}

export function buildMeetingPlatformEvidenceCorrelation(
  platform: string,
  input?: MeetingPlatformEvidenceCorrelationInput,
  options?: MeetingPlatformEvidenceCorrelationOptions,
): MeetingPlatformEvidenceCorrelation;
