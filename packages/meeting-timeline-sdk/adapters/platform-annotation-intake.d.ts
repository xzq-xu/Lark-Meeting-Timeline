import type { TimelineMarkInput } from '../index.mjs';
import type { MeetingPlatformRolloutOptions } from './platform-rollout.mjs';

export const MEETING_PLATFORM_ANNOTATION_INTAKE_PLAN_SCHEMA: 'meeting_platform_annotation_intake_plan';
export const MEETING_PLATFORM_ANNOTATION_INTAKE_MATRIX_SCHEMA: 'meeting_platform_annotation_intake_matrix';
export const MEETING_PLATFORM_ANNOTATION_INTAKE_SCHEMA: 'meeting_platform_annotation_intake';
export const MEETING_PLATFORM_ANNOTATION_INTAKE_SCHEMA_VERSION: number;

export interface MeetingPlatformAnnotationIntakeOptions extends MeetingPlatformRolloutOptions {
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  allowPendingWhenNoAxis?: boolean;
  allow_pending_when_no_axis?: boolean;
  allowOpenSessionWhenNoAxis?: boolean;
  allow_open_session_when_no_axis?: boolean;
  allowLocalSimulationAxis?: boolean;
  allow_local_simulation_axis?: boolean;
  allowedBeforeStartMs?: number;
  allowed_before_start_ms?: number;
  openSessionSource?: string;
  open_session_source?: string;
  defaultTitle?: string;
  default_title?: string;
  [key: string]: unknown;
}

export interface MeetingPlatformAnnotationIntakeInput {
  platform?: string;
  annotation?: TimelineMarkInput | Record<string, unknown>;
  mark?: TimelineMarkInput | Record<string, unknown>;
  item?: TimelineMarkInput | Record<string, unknown>;
  current_meeting?: Record<string, unknown>;
  currentMeeting?: Record<string, unknown>;
  currentAxis?: Record<string, unknown>;
  current_axis?: Record<string, unknown>;
  meeting?: Record<string, unknown>;
  state?: Record<string, unknown>;
  timelineState?: Record<string, unknown>;
  timeline_state?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingPlatformAnnotationIntakePlan {
  type: 'meeting_platform_annotation_intake_plan';
  schema: 'meeting_platform_annotation_intake_plan';
  schema_version: number;
  platform: string;
  display_name?: string;
  status: string;
  accepted_time_fields: string[];
  route_policy: Record<string, unknown>;
  output_contract: Record<string, unknown>;
  realtime_policy: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingPlatformAnnotationIntakeMatrix {
  type: 'meeting_platform_annotation_intake_matrix';
  schema: 'meeting_platform_annotation_intake_matrix';
  schema_version: number;
  platform_count: number;
  provider_blocking_count: number;
  transcript_blocking_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  plans: MeetingPlatformAnnotationIntakePlan[];
  next_actions: string[];
}

export interface MeetingPlatformAnnotationIntake {
  type: 'meeting_platform_annotation_intake';
  schema: 'meeting_platform_annotation_intake';
  schema_version: number;
  platform: string;
  status: string;
  action: string;
  accepted_for_realtime: boolean;
  timing_reliable: boolean;
  requires_device_captured_at: boolean;
  axis_state: string;
  current_meeting?: Record<string, unknown>;
  captured_at_ms?: number;
  captured_at_source?: string | null;
  normalized_time_ms?: number;
  after_meeting_end_ms?: number;
  warnings: string[];
  annotation: Record<string, unknown>;
  insert_payload?: Record<string, unknown>;
  open_session_payload?: Record<string, unknown>;
  route_policy: Record<string, unknown>;
  next_actions: string[];
}

export function buildMeetingPlatformAnnotationIntakePlan(
  platform: string,
  options?: MeetingPlatformAnnotationIntakeOptions,
): MeetingPlatformAnnotationIntakePlan;

export function buildMeetingPlatformAnnotationIntakeMatrix(
  options?: MeetingPlatformAnnotationIntakeOptions,
): MeetingPlatformAnnotationIntakeMatrix;

export function buildMeetingPlatformAnnotationIntake(
  platform: string,
  input?: MeetingPlatformAnnotationIntakeInput | Record<string, unknown>,
  options?: MeetingPlatformAnnotationIntakeOptions,
): MeetingPlatformAnnotationIntake;
