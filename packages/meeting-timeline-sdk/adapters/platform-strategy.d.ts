import type { MeetingPlatformRolloutOptions } from './platform-rollout.mjs';

export const MEETING_PLATFORM_ADAPTATION_STRATEGY_SCHEMA: 'meeting_platform_adaptation_strategy';
export const MEETING_PLATFORM_ADAPTATION_STRATEGY_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdaptationStrategy {
  schema: 'meeting_platform_adaptation_strategy';
  schema_version: 1;
  type: 'meeting_platform_adaptation_strategy';
  platform: string;
  display_name?: string;
  rollout_status?: string;
  production_ready: boolean;
  ready_for_realtime_annotations: boolean;
  recommendation: string;
  source_priority?: string[];
  realtime_axis: {
    primary_source?: string;
    provider_events_block_realtime: boolean;
    transcript_blocks_realtime: boolean;
    timestamp_invariant?: string;
    annotation_time_field: string;
    signal_types: string[];
    delay_policy?: string;
  };
  local_observer?: Record<string, unknown>;
  provider_events?: Record<string, unknown>;
  speaker_activity?: Record<string, unknown>;
  post_meeting_transcript?: Record<string, unknown>;
  evidence_contract: Record<string, unknown>;
  next_actions: string[];
  rollout_plan?: Record<string, unknown>;
}

export interface MeetingPlatformAdaptationStrategyMatrix {
  type: 'meeting_platform_adaptation_strategy_matrix';
  schema: 'meeting_platform_adaptation_strategy';
  schema_version: 1;
  strategy_count: number;
  production_ready_count: number;
  realtime_ready_count: number;
  local_first_count: number;
  non_blocking_provider_count: number;
  platforms: string[];
  rows: Record<string, unknown>[];
  strategies: MeetingPlatformAdaptationStrategy[];
}

export function buildMeetingPlatformAdaptationStrategy(
  platform: string,
  options?: MeetingPlatformRolloutOptions & {
    includeRolloutPlan?: boolean;
    include_rollout_plan?: boolean;
  },
): MeetingPlatformAdaptationStrategy;

export function buildAllMeetingPlatformAdaptationStrategies(
  options?: MeetingPlatformRolloutOptions & {
    platforms?: string[];
    platform_keys?: string[];
  },
): MeetingPlatformAdaptationStrategy[];

export function buildMeetingPlatformAdaptationStrategyMatrix(
  options?: MeetingPlatformRolloutOptions & {
    platforms?: string[];
    platform_keys?: string[];
  },
): MeetingPlatformAdaptationStrategyMatrix;
