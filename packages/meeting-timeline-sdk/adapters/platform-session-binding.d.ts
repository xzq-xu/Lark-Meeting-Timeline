import type { MeetingPlatformRolloutOptions } from './platform-rollout.mjs';

export const MEETING_PLATFORM_SESSION_BINDING_PLAN_SCHEMA: 'meeting_platform_session_binding_plan';
export const MEETING_PLATFORM_SESSION_BINDING_MATRIX_SCHEMA: 'meeting_platform_session_binding_matrix';
export const MEETING_PLATFORM_SESSION_BINDING_SCHEMA: 'meeting_platform_session_binding';
export const MEETING_PLATFORM_SESSION_BINDING_SCHEMA_VERSION: number;

export interface MeetingPlatformSessionBindingOptions extends MeetingPlatformRolloutOptions {
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  acceptScore?: number;
  accept_score?: number;
  conflictScore?: number;
  conflict_score?: number;
  maxStartDeltaMs?: number;
  max_start_delta_ms?: number;
  includeEmptyCandidates?: boolean;
  include_empty_candidates?: boolean;
  [key: string]: unknown;
}

export interface MeetingPlatformSessionBindingInput {
  platform?: string;
  current_meeting?: Record<string, unknown>;
  currentMeeting?: Record<string, unknown>;
  current_axis?: Record<string, unknown>;
  currentAxis?: Record<string, unknown>;
  meeting?: Record<string, unknown>;
  local_observer?: Record<string, unknown>;
  localObserver?: Record<string, unknown>;
  local_observers?: Array<Record<string, unknown>>;
  localObservers?: Array<Record<string, unknown>>;
  local_meeting?: Record<string, unknown>;
  localMeeting?: Record<string, unknown>;
  local_meetings?: Array<Record<string, unknown>>;
  localMeetings?: Array<Record<string, unknown>>;
  browser?: Record<string, unknown>;
  window?: Record<string, unknown>;
  tab?: Record<string, unknown>;
  provider_signal?: Record<string, unknown>;
  providerSignal?: Record<string, unknown>;
  provider_signals?: Array<Record<string, unknown>>;
  providerSignals?: Array<Record<string, unknown>>;
  provider_event?: Record<string, unknown>;
  providerEvent?: Record<string, unknown>;
  provider_events?: Array<Record<string, unknown>>;
  providerEvents?: Array<Record<string, unknown>>;
  signal?: Record<string, unknown>;
  signals?: Array<Record<string, unknown>>;
  annotation?: Record<string, unknown>;
  mark?: Record<string, unknown>;
  item?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingPlatformSessionBindingPlan {
  type: 'meeting_platform_session_binding_plan';
  schema: 'meeting_platform_session_binding_plan';
  schema_version: number;
  platform: string;
  display_name?: string;
  status: string;
  identity_fields: string[];
  match_policy: Record<string, unknown>;
  realtime_policy: Record<string, unknown>;
  output_contract: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingPlatformSessionBindingMatrix {
  type: 'meeting_platform_session_binding_matrix';
  schema: 'meeting_platform_session_binding_matrix';
  schema_version: number;
  platform_count: number;
  provider_blocking_count: number;
  transcript_blocking_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  plans: MeetingPlatformSessionBindingPlan[];
  next_actions: string[];
}

export interface MeetingPlatformSessionBinding {
  type: 'meeting_platform_session_binding';
  schema: 'meeting_platform_session_binding';
  schema_version: number;
  platform: string;
  status: string;
  accepted_for_realtime: boolean;
  bind_to_current_axis: boolean;
  should_start_axis: boolean;
  should_store_pending: boolean;
  reference_candidate_id?: string;
  selected_meeting?: Record<string, unknown>;
  start_payload?: Record<string, unknown>;
  candidate_count: number;
  match_count: number;
  positive_match_count: number;
  conflict_count: number;
  candidates: Array<Record<string, unknown>>;
  matches: Array<Record<string, unknown>>;
  conflicts: Array<Record<string, unknown>>;
  policy: Record<string, unknown>;
  next_actions: string[];
}

export function buildMeetingPlatformSessionBindingPlan(
  platform: string,
  options?: MeetingPlatformSessionBindingOptions,
): MeetingPlatformSessionBindingPlan;

export function buildMeetingPlatformSessionBindingMatrix(
  options?: MeetingPlatformSessionBindingOptions,
): MeetingPlatformSessionBindingMatrix;

export function buildMeetingPlatformSessionBinding(
  platform: string,
  input?: MeetingPlatformSessionBindingInput,
  options?: MeetingPlatformSessionBindingOptions,
): MeetingPlatformSessionBinding;
