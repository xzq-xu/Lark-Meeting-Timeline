import type { NormalizedMeetingSignal } from './core.mjs';
import type { MeetingPlatformRolloutOptions } from './platform-rollout.mjs';

export const MEETING_PLATFORM_PARTICIPANT_TRACK_PLAN_SCHEMA: 'meeting_platform_participant_track_plan';
export const MEETING_PLATFORM_PARTICIPANT_TRACK_MATRIX_SCHEMA: 'meeting_platform_participant_track_matrix';
export const MEETING_PLATFORM_PARTICIPANT_TRACK_SCHEMA: 'meeting_platform_participant_track';
export const MEETING_PLATFORM_PARTICIPANT_TRACK_SCHEMA_VERSION: number;

export interface MeetingPlatformParticipantTrackOptions extends MeetingPlatformRolloutOptions {
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  duplicateWindowMs?: number;
  duplicate_window_ms?: number;
  suppressReconnectGapMs?: number;
  suppress_reconnect_gap_ms?: number;
  leaveStableMs?: number;
  leave_stable_ms?: number;
  emitInitialRoster?: boolean;
  emit_initial_roster?: boolean;
  outputJoinKind?: string;
  output_join_kind?: string;
  outputLeftKind?: string;
  output_left_kind?: string;
  closePendingLeavesAtMs?: number;
  close_pending_leaves_at_ms?: number;
  source?: string;
  detectorSource?: string;
  detector_source?: string;
  defaults?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingPlatformParticipantSnapshot {
  meeting?: Record<string, unknown>;
  observed_at_ms?: number;
  observedAtMs?: number;
  occurred_at_ms?: number;
  occurredAtMs?: number;
  timestamp_ms?: number;
  timestampMs?: number;
  participants?: Array<Record<string, unknown>>;
  roster?: Array<Record<string, unknown>>;
  members?: Array<Record<string, unknown>>;
  people?: Array<Record<string, unknown>>;
  source_event_id?: string;
  sourceEventId?: string;
  [key: string]: unknown;
}

export interface MeetingPlatformParticipantTrackInput {
  snapshots?: MeetingPlatformParticipantSnapshot[];
  participantSnapshots?: MeetingPlatformParticipantSnapshot[];
  participant_snapshots?: MeetingPlatformParticipantSnapshot[];
  rosterSnapshots?: MeetingPlatformParticipantSnapshot[];
  roster_snapshots?: MeetingPlatformParticipantSnapshot[];
  signals?: Array<NormalizedMeetingSignal | Record<string, unknown>>;
  events?: Array<NormalizedMeetingSignal | Record<string, unknown>>;
  defaults?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingPlatformParticipantTrackPlan {
  type: 'meeting_platform_participant_track_plan';
  schema: 'meeting_platform_participant_track_plan';
  schema_version: number;
  platform: string;
  display_name?: string;
  status: string;
  realtime_ready_when_snapshots_available: boolean;
  provider_events_block_realtime: boolean;
  transcript_blocks_realtime: boolean;
  track_source_order: Array<Record<string, unknown>>;
  filter_policy: Record<string, unknown>;
  output_contract: Record<string, unknown>;
  required_input_fields: string[];
  next_actions: string[];
}

export interface MeetingPlatformParticipantTrackMatrix {
  type: 'meeting_platform_participant_track_matrix';
  schema: 'meeting_platform_participant_track_matrix';
  schema_version: number;
  platform_count: number;
  realtime_ready_when_snapshots_available_count: number;
  provider_blocking_count: number;
  transcript_blocking_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  plans: MeetingPlatformParticipantTrackPlan[];
  next_actions: string[];
}

export interface MeetingPlatformParticipantTrackMark {
  id: string;
  source: string;
  captured_at_ms: number;
  kind: string;
  label: string;
  intent: 'participant_track';
  payload: Record<string, unknown>;
}

export interface MeetingPlatformParticipantTrack {
  type: 'meeting_platform_participant_track';
  schema: 'meeting_platform_participant_track';
  schema_version: number;
  platform: string;
  status: string;
  input_snapshot_count: number;
  normalized_signal_count: number;
  generated_signal_count: number;
  signal_count: number;
  mark_count: number;
  filter_policy: Record<string, unknown>;
  signals: NormalizedMeetingSignal[];
  marks: MeetingPlatformParticipantTrackMark[];
  diagnostics: Record<string, unknown>;
  next_actions: string[];
}

export function buildMeetingPlatformParticipantTrackPlan(
  platform: string,
  options?: MeetingPlatformParticipantTrackOptions,
): MeetingPlatformParticipantTrackPlan;

export function buildMeetingPlatformParticipantTrackMatrix(
  options?: MeetingPlatformParticipantTrackOptions,
): MeetingPlatformParticipantTrackMatrix;

export function buildMeetingPlatformParticipantTrack(
  platform: string,
  input?: MeetingPlatformParticipantTrackInput | Array<Record<string, unknown>>,
  options?: MeetingPlatformParticipantTrackOptions,
): MeetingPlatformParticipantTrack;
