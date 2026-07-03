import type { NormalizedMeetingSignal } from './core.mjs';
import type {
  ActiveSpeakerObserverState,
  ActiveSpeakerRawSample,
} from './active-speaker.mjs';
import type { MeetingPlatformRolloutOptions } from './platform-rollout.mjs';

export const MEETING_PLATFORM_SPEAKER_TRACK_PLAN_SCHEMA: 'meeting_platform_speaker_track_plan';
export const MEETING_PLATFORM_SPEAKER_TRACK_MATRIX_SCHEMA: 'meeting_platform_speaker_track_matrix';
export const MEETING_PLATFORM_SPEAKER_TRACK_SCHEMA: 'meeting_platform_speaker_track';
export const MEETING_PLATFORM_SPEAKER_TRACK_SCHEMA_VERSION: number;

export interface MeetingPlatformSpeakerTrackOptions extends MeetingPlatformRolloutOptions {
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  minStableMs?: number;
  min_stable_ms?: number;
  switchStableMs?: number;
  switch_stable_ms?: number;
  endIdleMs?: number;
  end_idle_ms?: number;
  silenceEndMs?: number;
  silence_end_ms?: number;
  minSegmentMs?: number;
  min_segment_ms?: number;
  mergeGapMs?: number;
  merge_gap_ms?: number;
  duplicateWindowMs?: number;
  duplicate_window_ms?: number;
  emitSpeakerEnd?: boolean;
  emit_speaker_end?: boolean;
  outputMarkKind?: string;
  output_mark_kind?: string;
  closeOpenSegmentsAtMs?: number;
  close_open_segments_at_ms?: number;
  initialState?: ActiveSpeakerObserverState | null;
  initial_state?: ActiveSpeakerObserverState | null;
  source?: string;
  detectorSource?: string;
  detector_source?: string;
  defaults?: Record<string, unknown>;
  observerOptions?: Record<string, unknown>;
  observer_options?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingPlatformSpeakerTrackInput {
  samples?: ActiveSpeakerRawSample[];
  activeSpeakerSamples?: ActiveSpeakerRawSample[];
  active_speaker_samples?: ActiveSpeakerRawSample[];
  observations?: ActiveSpeakerRawSample[];
  rows?: ActiveSpeakerRawSample[];
  signals?: Array<NormalizedMeetingSignal | Record<string, unknown>>;
  events?: Array<NormalizedMeetingSignal | Record<string, unknown>>;
  defaults?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingPlatformSpeakerTrackPlan {
  type: 'meeting_platform_speaker_track_plan';
  schema: 'meeting_platform_speaker_track_plan';
  schema_version: number;
  platform: string;
  display_name?: string;
  status: string;
  realtime_ready_when_samples_available: boolean;
  provider_events_block_realtime: boolean;
  transcript_blocks_realtime: boolean;
  track_source_order: Array<Record<string, unknown>>;
  filter_policy: Record<string, unknown>;
  output_contract: Record<string, unknown>;
  required_input_fields: string[];
  next_actions: string[];
}

export interface MeetingPlatformSpeakerTrackMatrix {
  type: 'meeting_platform_speaker_track_matrix';
  schema: 'meeting_platform_speaker_track_matrix';
  schema_version: number;
  platform_count: number;
  realtime_ready_when_samples_available_count: number;
  provider_blocking_count: number;
  transcript_blocking_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  plans: MeetingPlatformSpeakerTrackPlan[];
  next_actions: string[];
}

export interface MeetingPlatformSpeakerTrackSegment {
  id: string;
  meeting?: Record<string, unknown>;
  platform?: string;
  meeting_id?: string;
  speaker_id?: string;
  speaker_name?: string;
  participant_id?: string;
  participant_name?: string;
  start_at_ms: number;
  end_at_ms?: number;
  duration_ms?: number;
  start_source_event_id?: string;
  end_source_event_id?: string;
  source?: string;
  raw?: unknown;
}

export interface MeetingPlatformSpeakerTrackMark {
  id: string;
  source: string;
  captured_at_ms: number;
  kind: string;
  label: string;
  intent: 'speaker_track';
  payload: Record<string, unknown>;
}

export interface MeetingPlatformSpeakerTrack {
  type: 'meeting_platform_speaker_track';
  schema: 'meeting_platform_speaker_track';
  schema_version: number;
  platform: string;
  status: string;
  input_sample_count: number;
  normalized_signal_count: number;
  generated_signal_count: number;
  signal_count: number;
  segment_count: number;
  mark_count: number;
  filter_policy: Record<string, unknown>;
  signals: NormalizedMeetingSignal[];
  segments: MeetingPlatformSpeakerTrackSegment[];
  active_segment?: MeetingPlatformSpeakerTrackSegment;
  marks: MeetingPlatformSpeakerTrackMark[];
  observer_state?: ActiveSpeakerObserverState | null;
  diagnostics: Record<string, unknown>;
  next_actions: string[];
}

export function buildMeetingPlatformSpeakerTrackPlan(
  platform: string,
  options?: MeetingPlatformSpeakerTrackOptions,
): MeetingPlatformSpeakerTrackPlan;

export function buildMeetingPlatformSpeakerTrackMatrix(
  options?: MeetingPlatformSpeakerTrackOptions,
): MeetingPlatformSpeakerTrackMatrix;

export function buildMeetingPlatformSpeakerTrack(
  platform: string,
  input?: MeetingPlatformSpeakerTrackInput | ActiveSpeakerRawSample[],
  options?: MeetingPlatformSpeakerTrackOptions,
): MeetingPlatformSpeakerTrack;
