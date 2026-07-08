import type { MeetingPlatformAdapterBlueprint } from './platform-adapter-blueprint.mjs';
import type {
  MeetingPlatformAdapterPreflight,
  MeetingPlatformAdapterPreflightOptions,
} from './platform-adapter-preflight.mjs';
import type { MeetingPlatformAdapterSelection } from './platform-adapter-selection.mjs';
import type { MeetingPlatformProviderReplayReport } from './platform-ingest.mjs';

export const MEETING_PLATFORM_REALTIME_AXIS_READINESS_SCHEMA: 'meeting_platform_realtime_axis_readiness';
export const MEETING_PLATFORM_REALTIME_AXIS_READINESS_MATRIX_SCHEMA: 'meeting_platform_realtime_axis_readiness_matrix';
export const MEETING_PLATFORM_REALTIME_AXIS_READINESS_SCHEMA_VERSION: 1;

export interface MeetingPlatformRealtimeAxisReadinessOptions extends MeetingPlatformAdapterPreflightOptions {
  providerRecords?: unknown[];
  provider_records?: unknown[];
  providerEvents?: unknown[];
  provider_events?: unknown[];
  providerRecordsByPlatform?: Record<string, unknown[]>;
  provider_records_by_platform?: Record<string, unknown[]>;
  providerEventsByPlatform?: Record<string, unknown[]>;
  provider_events_by_platform?: Record<string, unknown[]>;
  [key: string]: unknown;
}

export interface MeetingPlatformRealtimeAxisReadinessIssue {
  severity: 'info' | 'warning' | 'error';
  code: string;
  message: string;
  [key: string]: unknown;
}

export interface MeetingPlatformRealtimeAxisReadiness {
  type: 'meeting_platform_realtime_axis_readiness';
  schema: 'meeting_platform_realtime_axis_readiness';
  schema_version: 1;
  accepted: boolean;
  status: string;
  platform: string;
  display_name?: string;
  realtime_axis_ready: boolean;
  production_handoff_ready: boolean;
  axis_source?: string;
  axis_surface?: string;
  selected_surface?: string;
  timestamp_field?: string;
  local_observer: Record<string, unknown>;
  provider_replay: Partial<MeetingPlatformProviderReplayReport> & Record<string, unknown>;
  adapter_selection: Partial<MeetingPlatformAdapterSelection> & Record<string, unknown>;
  adapter_blueprint: Partial<MeetingPlatformAdapterBlueprint> & Record<string, unknown>;
  runtime_policy: {
    provider_events_block_realtime: boolean;
    transcript_blocks_realtime: boolean;
    provider_replay_required_for_production_handoff: boolean;
    provider_replay_required_for_realtime_axis: boolean;
    local_observer_required_for_realtime_axis: boolean;
  };
  issue_count: number;
  issues: MeetingPlatformRealtimeAxisReadinessIssue[];
  next_actions: string[];
}

export interface MeetingPlatformRealtimeAxisReadinessMatrix {
  type: 'meeting_platform_realtime_axis_readiness_matrix';
  schema: 'meeting_platform_realtime_axis_readiness_matrix';
  schema_version: 1;
  platform_count: number;
  accepted_count: number;
  realtime_axis_ready_count: number;
  production_handoff_ready_count: number;
  local_observer_ready_count: number;
  live_evidence_ready_count: number;
  meeting_start_ready_count: number;
  speaker_track_ready_count: number;
  provider_replay_accepted_count: number;
  provider_blocking_count: number;
  transcript_blocking_count: number;
  waiting_for_live_evidence_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  reports: MeetingPlatformRealtimeAxisReadiness[];
  next_actions: string[];
}

export function buildMeetingPlatformRealtimeAxisReadiness(
  platformOrInput?: string | URL | Record<string, unknown> | MeetingPlatformAdapterPreflight,
  inputOrOptions?: Record<string, unknown> | MeetingPlatformRealtimeAxisReadinessOptions,
  maybeOptions?: MeetingPlatformRealtimeAxisReadinessOptions,
): MeetingPlatformRealtimeAxisReadiness;

export function buildMeetingPlatformRealtimeAxisReadinessMatrix(
  input?: Record<string, unknown>,
  options?: MeetingPlatformRealtimeAxisReadinessOptions,
): MeetingPlatformRealtimeAxisReadinessMatrix;

export function assertMeetingPlatformRealtimeAxisReadiness(
  platformOrInput?: string | URL | Record<string, unknown> | MeetingPlatformRealtimeAxisReadiness,
  inputOrOptions?: Record<string, unknown> | MeetingPlatformRealtimeAxisReadinessOptions,
  maybeOptions?: MeetingPlatformRealtimeAxisReadinessOptions,
): MeetingPlatformRealtimeAxisReadiness;

export function assertMeetingPlatformRealtimeAxisReadinessMatrix(
  input?: Record<string, unknown> | MeetingPlatformRealtimeAxisReadinessMatrix,
  options?: MeetingPlatformRealtimeAxisReadinessOptions,
): MeetingPlatformRealtimeAxisReadinessMatrix;
