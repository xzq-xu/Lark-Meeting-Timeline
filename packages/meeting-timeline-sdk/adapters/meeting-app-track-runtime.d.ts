import type { MeetingTimelineClient, MeetingTimelineClientOptions } from '../index.mjs';
import type {
  MeetingAppTrackPipeline,
  MeetingAppTrackPipelineInput,
  MeetingAppTrackPipelineOptions,
} from './meeting-app-track-pipeline.mjs';

export const MEETING_APP_TRACK_RUNTIME_SCHEMA: 'meeting_app_track_runtime';
export const MEETING_APP_TRACK_RUNTIME_SCHEMA_VERSION: number;

export interface MeetingAppTrackRuntimeOptions extends MeetingAppTrackPipelineOptions {
  clientOptions?: MeetingTimelineClientOptions;
  client_options?: MeetingTimelineClientOptions;
  initialSnapshots?: Array<Record<string, unknown>>;
  initial_snapshots?: Array<Record<string, unknown>>;
  seenMarkKeys?: Iterable<string> | string[];
  seen_mark_keys?: Iterable<string> | string[];
  maxSnapshots?: number;
  max_snapshots?: number;
  insert?: boolean;
  apply?: boolean;
  write?: boolean;
  dryRun?: boolean;
  dry_run?: boolean;
  insertOptions?: Record<string, unknown>;
  insert_options?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingAppTrackRuntimeObservation {
  type: 'meeting_app_track_runtime_observation';
  schema: 'meeting_app_track_runtime';
  schema_version: number;
  inserted: boolean;
  input_snapshot_count: number;
  retained_snapshot_count: number;
  new_mark_count: number;
  duplicate_mark_count: number;
  result_count: number;
  report: MeetingAppTrackPipeline;
  new_marks: Array<Record<string, unknown>>;
  duplicate_marks: Array<Record<string, unknown>>;
  results: unknown[];
  state: Record<string, unknown>;
}

export interface MeetingAppTrackRuntimePreview {
  report: MeetingAppTrackPipeline;
  new_marks: Array<Record<string, unknown>>;
  new_mark_count: number;
}

export interface MeetingAppTrackRuntime {
  client: MeetingTimelineClient;
  observe(input?: MeetingAppTrackPipelineInput | Array<Record<string, unknown>>, options?: MeetingAppTrackRuntimeOptions): Promise<MeetingAppTrackRuntimeObservation>;
  push(input?: MeetingAppTrackPipelineInput | Array<Record<string, unknown>>, options?: MeetingAppTrackRuntimeOptions): Promise<MeetingAppTrackRuntimeObservation>;
  preview(input?: MeetingAppTrackPipelineInput | Array<Record<string, unknown>>, options?: MeetingAppTrackRuntimeOptions): MeetingAppTrackRuntimePreview;
  getState(): Record<string, unknown>;
  getSnapshots(): Array<Record<string, unknown>>;
  reset(nextState?: Record<string, unknown>): Record<string, unknown>;
}

export function createMeetingAppTrackRuntime(
  clientOrOptions: MeetingTimelineClient | MeetingTimelineClientOptions,
  options?: MeetingAppTrackRuntimeOptions,
): MeetingAppTrackRuntime;
