import type { MeetingPlatformRolloutOptions } from './platform-rollout.mjs';

export const MEETING_PLATFORM_RUNTIME_PROFILE_SCHEMA: 'meeting_platform_runtime_profile';
export const MEETING_PLATFORM_RUNTIME_PROFILE_MATRIX_SCHEMA: 'meeting_platform_runtime_profile_matrix';
export const MEETING_PLATFORM_RUNTIME_PROFILE_SCHEMA_VERSION: 1;

export interface MeetingPlatformRuntimeProfileOptions extends MeetingPlatformRolloutOptions {
  platforms?: string[];
  platform_keys?: string[];
  speakerFilter?: Record<string, unknown>;
  speaker_filter?: Record<string, unknown>;
  activeSpeakerFilter?: Record<string, unknown>;
  active_speaker_filter?: Record<string, unknown>;
}

export interface MeetingPlatformRuntimeProfile {
  type: 'meeting_platform_runtime_profile';
  schema: 'meeting_platform_runtime_profile';
  schema_version: 1;
  platform: string;
  display_name?: string;
  objective: string;
  runtime_contract: {
    annotation_timestamp_field: string;
    axis_timebase: string;
    relative_mark_time_allowed: boolean;
    per_meeting_annotation_isolation_required: boolean;
    provider_events_block_realtime: boolean;
    transcript_blocks_realtime: boolean;
    can_insert_annotation_before_provider_start_event: boolean;
  };
  axis: {
    primary_source?: string;
    start: Record<string, unknown>;
    end: Record<string, unknown>;
    provider_reconcile: Record<string, unknown>;
  };
  adapter_surfaces: {
    primary: string;
    recommended_order: string[];
    launch_context: string;
    realtime_axis_surface: string;
    provider_reconcile_surface?: string | null;
    candidate_detection: string[];
    rationale: string;
  };
  launch_requirements: Record<string, unknown>;
  evidence_thresholds: Record<string, unknown>;
  fallback_policy: Record<string, unknown>;
  annotations: Record<string, unknown>;
  speaker_markers: {
    enabled: boolean;
    content_policy: string;
    primary_source?: string;
    provider_support?: string;
    provider_realtime_required: boolean;
    event_types: string[];
    sdk_module: string;
    filter: Record<string, unknown>;
    marker_payload_contract: Record<string, unknown>;
    backfill?: string;
  };
  provider_events: Record<string, unknown>;
  transcript: Record<string, unknown>;
  implementation_order: string[];
  risks: string[];
  next_actions: string[];
}

export interface MeetingPlatformRuntimeProfileMatrix {
  type: 'meeting_platform_runtime_profile_matrix';
  schema: 'meeting_platform_runtime_profile_matrix';
  schema_version: 1;
  platform_count: number;
  non_blocking_provider_count: number;
  transcript_non_blocking_count: number;
  speaker_marker_enabled_count: number;
  platforms: string[];
  rows: Record<string, unknown>[];
  profiles: MeetingPlatformRuntimeProfile[];
}

export function buildMeetingPlatformRuntimeProfile(
  platform: string,
  options?: MeetingPlatformRuntimeProfileOptions,
): MeetingPlatformRuntimeProfile;

export function buildMeetingPlatformRuntimeProfileMatrix(
  options?: MeetingPlatformRuntimeProfileOptions,
): MeetingPlatformRuntimeProfileMatrix;
