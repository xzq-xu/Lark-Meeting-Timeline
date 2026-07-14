import type { TimelineMarkInput } from '../index.mjs';
import type { MeetingPlatformRolloutOptions } from './platform-rollout.mjs';

export const MEETING_PLATFORM_CLOCK_SYNC_PLAN_SCHEMA: 'meeting_platform_clock_sync_plan';
export const MEETING_PLATFORM_CLOCK_SYNC_MATRIX_SCHEMA: 'meeting_platform_clock_sync_matrix';
export const MEETING_PLATFORM_CLOCK_SYNC_REPORT_SCHEMA: 'meeting_platform_clock_sync_report';
export const MEETING_PLATFORM_CLOCK_SYNC_SCHEMA_VERSION: number;

export interface MeetingPlatformClockSyncOptions extends MeetingPlatformRolloutOptions {
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  endpoint?: string;
  timeSyncEndpoint?: string;
  time_sync_endpoint?: string;
  sampleCount?: number;
  sample_count?: number;
  maxRecommendedSkewMs?: number;
  max_recommended_skew_ms?: number;
  maxRecommendedRttMs?: number;
  max_recommended_rtt_ms?: number;
  maxUncertaintyMs?: number;
  max_uncertainty_ms?: number;
  applyOffsetToAnnotation?: boolean;
  apply_offset_to_annotation?: boolean;
  [key: string]: unknown;
}

export interface MeetingPlatformClockSyncSample {
  id?: string;
  sample_id?: string;
  sampleId?: string;
  source?: string;
  method?: string;
  client_send_at_ms?: number | string | Date;
  clientSendAtMs?: number | string | Date;
  client_receive_at_ms?: number | string | Date;
  clientReceiveAtMs?: number | string | Date;
  server_time_ms?: number | string | Date;
  serverTimeMs?: number | string | Date;
  offset_ms?: number;
  offsetMs?: number;
  clock_offset_ms?: number;
  clockOffsetMs?: number;
  rtt_ms?: number;
  rttMs?: number;
  [key: string]: unknown;
}

export interface MeetingPlatformClockSyncInput {
  platform?: string;
  samples?: MeetingPlatformClockSyncSample[];
  clock_samples?: MeetingPlatformClockSyncSample[];
  clockSamples?: MeetingPlatformClockSyncSample[];
  time_sync_samples?: MeetingPlatformClockSyncSample[];
  timeSyncSamples?: MeetingPlatformClockSyncSample[];
  clock_sync?: MeetingPlatformClockSyncSample;
  clockSync?: MeetingPlatformClockSyncSample;
  time_sync?: MeetingPlatformClockSyncSample;
  timeSync?: MeetingPlatformClockSyncSample;
  annotation?: TimelineMarkInput | Record<string, unknown>;
  mark?: TimelineMarkInput | Record<string, unknown>;
  item?: TimelineMarkInput | Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingPlatformClockSyncPlan {
  type: 'meeting_platform_clock_sync_plan';
  schema: 'meeting_platform_clock_sync_plan';
  schema_version: number;
  platform: string;
  display_name?: string;
  status: string;
  required: boolean;
  endpoint: string;
  algorithm: Record<string, unknown>;
  thresholds: Record<string, number>;
  input_contract: Record<string, unknown>;
  realtime_policy: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingPlatformClockSyncMatrix {
  type: 'meeting_platform_clock_sync_matrix';
  schema: 'meeting_platform_clock_sync_matrix';
  schema_version: number;
  platform_count: number;
  required_count: number;
  provider_blocking_count: number;
  transcript_blocking_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  plans: MeetingPlatformClockSyncPlan[];
  next_actions: string[];
}

export interface MeetingPlatformClockSyncReport {
  type: 'meeting_platform_clock_sync_report';
  schema: 'meeting_platform_clock_sync_report';
  schema_version: number;
  platform: string;
  status: string;
  accepted_for_realtime: boolean;
  endpoint: string;
  sample_count: number;
  usable_sample_count: number;
  selected_sample_id?: string;
  selected_rtt_ms?: number;
  selected_uncertainty_ms?: number;
  recommended_offset_ms?: number;
  median_offset_ms?: number;
  offset_jitter_ms?: number;
  max_recommended_skew_ms: number;
  max_recommended_rtt_ms: number;
  max_uncertainty_ms: number;
  operator_attention_required: boolean;
  warnings: string[];
  samples: Array<Record<string, unknown>>;
  annotation_time?: Record<string, unknown>;
  calibrated_annotation?: Record<string, unknown>;
  next_actions: string[];
}

export function buildMeetingPlatformClockSyncPlan(
  platform: string,
  options?: MeetingPlatformClockSyncOptions,
): MeetingPlatformClockSyncPlan;

export function buildMeetingPlatformClockSyncMatrix(
  options?: MeetingPlatformClockSyncOptions,
): MeetingPlatformClockSyncMatrix;

export function buildMeetingPlatformClockSyncReport(
  platform: string,
  input?: MeetingPlatformClockSyncInput,
  options?: MeetingPlatformClockSyncOptions,
): MeetingPlatformClockSyncReport;
