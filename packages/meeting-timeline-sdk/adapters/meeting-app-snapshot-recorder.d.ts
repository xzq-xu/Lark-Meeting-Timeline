import type { MeetingAppDomCaptureInput, MeetingAppDomCaptureOptions, MeetingAppDomCaptureSnapshot } from './meeting-app-capture.mjs';
import type { MeetingAppFixturePlatform } from './meeting-app-fixtures.mjs';

export const MEETING_APP_SNAPSHOT_RECORD_SCHEMA: string;
export const MEETING_APP_SNAPSHOT_RECORD_SCHEMA_VERSION: number;
export const MEETING_APP_SNAPSHOT_RECORD_SET_SCHEMA: string;
export const MEETING_APP_SNAPSHOT_RECORD_SET_SCHEMA_VERSION: number;

export interface MeetingAppSnapshotRecordOptions extends MeetingAppDomCaptureOptions {
  id?: string;
  recordId?: string;
  record_id?: string;
  label?: string;
  phase?: string;
  state?: string;
  capturedAtMs?: number | string | Date;
  captured_at_ms?: number | string | Date;
  meetingId?: string;
  meeting_id?: string;
  externalMeetingId?: string;
  external_meeting_id?: string;
  notes?: string;
  snapshot?: MeetingAppDomCaptureSnapshot | Record<string, unknown>;
  captureOptions?: MeetingAppDomCaptureOptions;
  capture_options?: MeetingAppDomCaptureOptions;
  forceRebuild?: boolean;
  force_rebuild?: boolean;
}

export interface MeetingAppSnapshotRecord {
  schema: string;
  schema_version: number;
  id: string;
  label?: string;
  phase?: string;
  platform?: MeetingAppFixturePlatform | string;
  provider?: MeetingAppFixturePlatform | string;
  capturedAtMs: number;
  captured_at_ms: number;
  url?: string;
  title?: string;
  meeting_id?: string;
  external_meeting_id?: string;
  source?: string;
  notes?: string;
  snapshot: MeetingAppDomCaptureSnapshot | Record<string, unknown>;
}

export interface MeetingAppSnapshotRecordSet {
  schema: string;
  schema_version: number;
  id: string;
  createdAtMs: number;
  created_at_ms: number;
  source?: string;
  label?: string;
  platform_count: number;
  platforms: string[];
  record_count: number;
  records: MeetingAppSnapshotRecord[];
}

export interface MeetingAppSnapshotRecorder {
  add(input?: MeetingAppDomCaptureInput | MeetingAppDomCaptureSnapshot | MeetingAppSnapshotRecord | Record<string, unknown>, options?: MeetingAppSnapshotRecordOptions): MeetingAppSnapshotRecord;
  capture(input?: MeetingAppDomCaptureInput | Document | Record<string, unknown>, options?: MeetingAppSnapshotRecordOptions): MeetingAppSnapshotRecord;
  exportRecords(options?: MeetingAppSnapshotRecordOptions): MeetingAppSnapshotRecordSet;
  gateInput(options?: Record<string, unknown>): {
    snapshots: Record<string, Record<string, unknown>[]>;
    snapshot_records: MeetingAppSnapshotRecord[];
  };
  reset(): { removed: number };
  getState(): Record<string, unknown>;
  records(): MeetingAppSnapshotRecord[];
  findByLabel(label: string): MeetingAppSnapshotRecord[];
}

export function buildMeetingAppSnapshotRecord(
  input?: MeetingAppDomCaptureInput | MeetingAppDomCaptureSnapshot | MeetingAppSnapshotRecord | Record<string, unknown>,
  options?: MeetingAppSnapshotRecordOptions,
): MeetingAppSnapshotRecord;

export function buildMeetingAppSnapshotRecordSet(
  records?: (MeetingAppSnapshotRecord | MeetingAppDomCaptureSnapshot | Record<string, unknown>)[],
  options?: MeetingAppSnapshotRecordOptions,
): MeetingAppSnapshotRecordSet;

export function meetingAppSnapshotRecords(input?: MeetingAppSnapshotRecordSet | MeetingAppSnapshotRecord[] | Record<string, unknown>): MeetingAppSnapshotRecord[];

export function meetingAppSnapshotsFromRecords(
  input?: MeetingAppSnapshotRecordSet | MeetingAppSnapshotRecord[] | Record<string, unknown>,
  options?: MeetingAppSnapshotRecordOptions,
): Record<string, unknown>[];

export function buildMeetingAppLaunchGateInputFromRecords(
  input?: MeetingAppSnapshotRecordSet | MeetingAppSnapshotRecord[] | Record<string, unknown>,
  options?: MeetingAppSnapshotRecordOptions,
): {
  snapshots: Record<string, Record<string, unknown>[]>;
  snapshot_records: MeetingAppSnapshotRecord[];
};

export function createMeetingAppSnapshotRecorder(options?: MeetingAppSnapshotRecordOptions): MeetingAppSnapshotRecorder;
