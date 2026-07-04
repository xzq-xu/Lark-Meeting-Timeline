export const MEETING_PLATFORM_RUNTIME_HOST_VERIFICATION_SCHEMA: 'meeting_platform_runtime_host_verification';
export const MEETING_PLATFORM_RUNTIME_HOST_VERIFICATION_MATRIX_SCHEMA: 'meeting_platform_runtime_host_verification_matrix';
export const MEETING_PLATFORM_RUNTIME_HOST_REPLAY_SCHEMA: 'meeting_platform_runtime_host_replay';
export const MEETING_PLATFORM_RUNTIME_HOST_REPLAY_MATRIX_SCHEMA: 'meeting_platform_runtime_host_replay_matrix';
export const MEETING_PLATFORM_RUNTIME_HOST_VERIFICATION_SCHEMA_VERSION: 1;

export interface MeetingPlatformRuntimeHostVerificationClient {
  calls: Array<{ action: string; input: Record<string, unknown> }>;
  startMeeting(input: Record<string, unknown>): Promise<Record<string, unknown>>;
  endMeeting(input: Record<string, unknown>): Promise<Record<string, unknown>>;
  insertMark(input: Record<string, unknown>): Promise<Record<string, unknown>>;
  insertMarks(input: unknown): Promise<Record<string, unknown>>;
  getState(): Record<string, unknown>;
}

export interface MeetingPlatformRuntimeHostFixtureEnvironment {
  platform: string;
  window: Record<string, unknown>;
  readonly document: Record<string, unknown>;
  readonly location: Record<string, unknown>;
  readonly snapshot: Record<string, unknown>;
  now(): number;
  setState(state?: string, options?: Record<string, unknown>): Record<string, unknown>;
  setActive(options?: Record<string, unknown>): Record<string, unknown>;
  setEnded(options?: Record<string, unknown>): Record<string, unknown>;
  getState(): Record<string, unknown>;
}

export interface MeetingPlatformRuntimeHostVerificationReport {
  type: 'meeting_platform_runtime_host_verification';
  schema: typeof MEETING_PLATFORM_RUNTIME_HOST_VERIFICATION_SCHEMA;
  version: typeof MEETING_PLATFORM_RUNTIME_HOST_VERIFICATION_SCHEMA_VERSION;
  platform: string;
  accepted: boolean;
  start_at_ms: number;
  end_at_ms: number;
  call_count: number;
  actions: string[];
  signal_types: string[];
  result_actions: string[];
  coverage: Record<string, boolean>;
  missing: string[];
  capture: {
    start: Record<string, unknown>;
    end: Record<string, unknown>;
  };
  calls: Array<{ action: string; input: Record<string, unknown> }>;
  host_state?: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingPlatformRuntimeHostVerificationMatrix {
  type: 'meeting_platform_runtime_host_verification_matrix';
  schema: typeof MEETING_PLATFORM_RUNTIME_HOST_VERIFICATION_MATRIX_SCHEMA;
  version: typeof MEETING_PLATFORM_RUNTIME_HOST_VERIFICATION_SCHEMA_VERSION;
  platform_count: number;
  accepted_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  reports: MeetingPlatformRuntimeHostVerificationReport[];
  next_actions: string[];
}

export interface MeetingPlatformRuntimeHostReplayReport {
  type: 'meeting_platform_runtime_host_replay';
  schema: typeof MEETING_PLATFORM_RUNTIME_HOST_REPLAY_SCHEMA;
  version: typeof MEETING_PLATFORM_RUNTIME_HOST_VERIFICATION_SCHEMA_VERSION;
  platform: string;
  accepted: boolean;
  input_record_count: number;
  replay_row_count: number;
  active_at_ms?: number;
  end_at_ms?: number;
  active_row?: Record<string, unknown>;
  end_row?: Record<string, unknown>;
  rows: Array<Record<string, unknown>>;
  call_count: number;
  actions: string[];
  signal_types: string[];
  result_actions: string[];
  coverage: Record<string, boolean>;
  missing: string[];
  calls: Array<{ action: string; input: Record<string, unknown> }>;
  host_state?: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingPlatformRuntimeHostReplayMatrix {
  type: 'meeting_platform_runtime_host_replay_matrix';
  schema: typeof MEETING_PLATFORM_RUNTIME_HOST_REPLAY_MATRIX_SCHEMA;
  version: typeof MEETING_PLATFORM_RUNTIME_HOST_VERIFICATION_SCHEMA_VERSION;
  platform_count: number;
  accepted_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  reports: MeetingPlatformRuntimeHostReplayReport[];
  next_actions: string[];
}

export function createMeetingPlatformRuntimeHostVerificationClient(options?: Record<string, unknown>): MeetingPlatformRuntimeHostVerificationClient;
export function createMeetingPlatformRuntimeHostFixtureEnvironment(platform: string, options?: Record<string, unknown>): MeetingPlatformRuntimeHostFixtureEnvironment;
export function runMeetingPlatformRuntimeHostVerification(platform: string, options?: Record<string, unknown>): Promise<MeetingPlatformRuntimeHostVerificationReport>;
export function runMeetingPlatformRuntimeHostVerificationMatrix(options?: Record<string, unknown>): Promise<MeetingPlatformRuntimeHostVerificationMatrix>;
export function runMeetingPlatformRuntimeHostReplay(platform: string, input?: Record<string, unknown> | unknown[], options?: Record<string, unknown>): Promise<MeetingPlatformRuntimeHostReplayReport>;
export function runMeetingPlatformRuntimeHostReplay(input: Record<string, unknown> | unknown[], options?: Record<string, unknown>): Promise<MeetingPlatformRuntimeHostReplayReport>;
export function runMeetingPlatformRuntimeHostReplayMatrix(options?: Record<string, unknown>): Promise<MeetingPlatformRuntimeHostReplayMatrix>;
export function assertMeetingPlatformRuntimeHostVerification(report?: MeetingPlatformRuntimeHostVerificationReport): MeetingPlatformRuntimeHostVerificationReport;
export function assertMeetingPlatformRuntimeHostVerificationMatrix(matrix?: MeetingPlatformRuntimeHostVerificationMatrix): MeetingPlatformRuntimeHostVerificationMatrix;
export function assertMeetingPlatformRuntimeHostReplay(report?: MeetingPlatformRuntimeHostReplayReport): MeetingPlatformRuntimeHostReplayReport;
export function assertMeetingPlatformRuntimeHostReplayMatrix(matrix?: MeetingPlatformRuntimeHostReplayMatrix): MeetingPlatformRuntimeHostReplayMatrix;
