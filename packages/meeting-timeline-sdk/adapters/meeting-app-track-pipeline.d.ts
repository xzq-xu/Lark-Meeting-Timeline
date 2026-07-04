import type { MeetingPlatformParticipantTrack } from './platform-participant-track.mjs';
import type { MeetingPlatformSpeakerTrack } from './platform-speaker-track.mjs';

export const MEETING_APP_TRACK_PIPELINE_SCHEMA: 'meeting_app_track_pipeline';
export const MEETING_APP_TRACK_PIPELINE_SCHEMA_VERSION: number;

export interface MeetingAppTrackPipelineOptions {
  platform?: string;
  provider?: string;
  meeting_id?: string;
  meetingId?: string;
  observedAtMs?: number | string | Date;
  observed_at_ms?: number | string | Date;
  speakerTrackOptions?: Record<string, unknown>;
  speaker_track_options?: Record<string, unknown>;
  participantTrackOptions?: Record<string, unknown>;
  participant_track_options?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingAppTrackPipelineInput {
  snapshots?: Array<Record<string, unknown>>;
  samples?: Array<Record<string, unknown>>;
  captures?: Array<Record<string, unknown>>;
  rows?: Array<Record<string, unknown>>;
  items?: Array<Record<string, unknown>>;
  observations?: Array<Record<string, unknown>>;
  [key: string]: unknown;
}

export interface MeetingAppTrackPipeline {
  type: 'meeting_app_track_pipeline';
  schema: 'meeting_app_track_pipeline';
  schema_version: number;
  status: string;
  platform?: string;
  meeting?: Record<string, unknown>;
  snapshot_count: number;
  mark_count: number;
  snapshots: Array<Record<string, unknown>>;
  speaker_samples: Array<Record<string, unknown>>;
  participant_snapshots: Array<Record<string, unknown>>;
  speaker_track?: MeetingPlatformSpeakerTrack | null;
  participant_track?: MeetingPlatformParticipantTrack | null;
  marks: Array<Record<string, unknown>>;
  coverage: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingAppTrackPipelineAccumulator {
  push(input?: MeetingAppTrackPipelineInput | Array<Record<string, unknown>>, options?: MeetingAppTrackPipelineOptions): MeetingAppTrackPipeline;
  build(options?: MeetingAppTrackPipelineOptions): MeetingAppTrackPipeline;
  getSnapshots(): Array<Record<string, unknown>>;
  reset(nextSnapshots?: MeetingAppTrackPipelineInput | Array<Record<string, unknown>>): MeetingAppTrackPipeline;
}

export function buildMeetingAppTrackPipeline(
  input?: MeetingAppTrackPipelineInput | Array<Record<string, unknown>>,
  options?: MeetingAppTrackPipelineOptions,
): MeetingAppTrackPipeline;

export function createMeetingAppTrackPipeline(
  options?: MeetingAppTrackPipelineOptions,
): MeetingAppTrackPipelineAccumulator;
