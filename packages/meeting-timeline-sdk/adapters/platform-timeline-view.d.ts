import type { MeetingPlatformRolloutOptions } from './platform-rollout.mjs';

export const MEETING_PLATFORM_TIMELINE_VIEW_PLAN_SCHEMA: 'meeting_platform_timeline_view_plan';
export const MEETING_PLATFORM_TIMELINE_VIEW_MATRIX_SCHEMA: 'meeting_platform_timeline_view_matrix';
export const MEETING_PLATFORM_TIMELINE_VIEW_SCHEMA: 'meeting_platform_timeline_view';
export const MEETING_PLATFORM_TIMELINE_VIEW_SCHEMA_VERSION: number;

export interface MeetingPlatformTimelineViewOptions extends MeetingPlatformRolloutOptions {
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  fullDurationMs?: number;
  full_duration_ms?: number;
  defaultDurationMs?: number;
  default_duration_ms?: number;
  tailPaddingMs?: number;
  tail_padding_ms?: number;
  minViewMs?: number;
  min_view_ms?: number;
  viewportStartMs?: number;
  viewport_start_ms?: number;
  windowStartMs?: number;
  window_start_ms?: number;
  viewportDurationMs?: number;
  viewport_duration_ms?: number;
  windowDurationMs?: number;
  window_duration_ms?: number;
  rails?: Array<Record<string, unknown> | string>;
  timelineRails?: Array<Record<string, unknown> | string>;
  timeline_rails?: Array<Record<string, unknown> | string>;
  [key: string]: unknown;
}

export interface MeetingPlatformTimelineViewInput {
  meeting?: Record<string, unknown>;
  session?: Record<string, unknown>;
  annotations?: Array<Record<string, unknown>>;
  sequence?: Array<Record<string, unknown>>;
  marks?: Array<Record<string, unknown>>;
  events?: Array<Record<string, unknown>>;
  signals?: Array<Record<string, unknown>>;
  speakerTrack?: Record<string, unknown>;
  speaker_track?: Record<string, unknown>;
  participantTrack?: Record<string, unknown>;
  participant_track?: Record<string, unknown>;
  artifactHandoff?: Record<string, unknown>;
  artifact_handoff?: Record<string, unknown>;
  transcript?: Record<string, unknown>;
  transcript_segments?: Array<Record<string, unknown>>;
  tracks?: Record<string, Record<string, unknown>>;
  [key: string]: unknown;
}

export interface MeetingPlatformTimelineRail {
  id: string;
  label: string;
  role?: string;
  order: number;
  marker_count: number;
}

export interface MeetingPlatformTimelineMarker {
  id: string;
  rail: string;
  source: string;
  kind: string;
  label: string;
  captured_at_ms?: number;
  time_ms?: number;
  calibrated: boolean;
  warnings: string[];
  visible: boolean;
  x_ratio?: number;
  raw?: unknown;
}

export interface MeetingPlatformTimelineViewport {
  start_ms: number;
  end_ms: number;
  duration_ms: number;
  full_duration_ms: number;
  min_view_ms: number;
  is_full: boolean;
}

export interface MeetingPlatformTimelineTick {
  ms: number;
  label: string;
}

export interface MeetingPlatformTimelineViewPlan {
  type: 'meeting_platform_timeline_view_plan';
  schema: 'meeting_platform_timeline_view_plan';
  schema_version: number;
  platform: string;
  display_name?: string;
  status: string;
  rails: MeetingPlatformTimelineRail[];
  input_contract: Record<string, unknown>;
  output_contract: Record<string, unknown>;
  realtime_policy: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingPlatformTimelineViewMatrix {
  type: 'meeting_platform_timeline_view_matrix';
  schema: 'meeting_platform_timeline_view_matrix';
  schema_version: number;
  platform_count: number;
  renderer_agnostic_count: number;
  provider_blocking_count: number;
  transcript_blocking_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  plans: MeetingPlatformTimelineViewPlan[];
  next_actions: string[];
}

export interface MeetingPlatformTimelineView {
  type: 'meeting_platform_timeline_view';
  schema: 'meeting_platform_timeline_view';
  schema_version: number;
  platform: string;
  status: string;
  meeting: Record<string, unknown>;
  viewport: MeetingPlatformTimelineViewport;
  ticks: MeetingPlatformTimelineTick[];
  rails: MeetingPlatformTimelineRail[];
  markers: MeetingPlatformTimelineMarker[];
  visible_markers: MeetingPlatformTimelineMarker[];
  uncalibrated_markers: MeetingPlatformTimelineMarker[];
  diagnostics: Record<string, unknown>;
  next_actions: string[];
}

export function buildMeetingPlatformTimelineViewPlan(
  platform: string,
  options?: MeetingPlatformTimelineViewOptions,
): MeetingPlatformTimelineViewPlan;

export function buildMeetingPlatformTimelineViewMatrix(
  options?: MeetingPlatformTimelineViewOptions,
): MeetingPlatformTimelineViewMatrix;

export function buildMeetingPlatformTimelineView(
  platform: string,
  input?: MeetingPlatformTimelineViewInput,
  options?: MeetingPlatformTimelineViewOptions,
): MeetingPlatformTimelineView;

export function zoomMeetingPlatformTimelineViewport(
  viewport: Partial<MeetingPlatformTimelineViewport>,
  factor?: number,
  anchorRatio?: number,
  options?: MeetingPlatformTimelineViewOptions,
): MeetingPlatformTimelineViewport;
