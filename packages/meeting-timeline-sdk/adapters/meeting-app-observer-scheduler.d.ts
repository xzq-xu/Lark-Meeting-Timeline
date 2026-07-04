import type {
  MeetingAppRuntimeObserverPlan,
} from './meeting-app-profile.mjs';

export const MEETING_APP_OBSERVER_SCHEDULER_CONFIG_SCHEMA: 'meeting_app_observer_scheduler_config';
export const MEETING_APP_OBSERVER_SCHEDULER_CONFIG_MATRIX_SCHEMA: 'meeting_app_observer_scheduler_config_matrix';
export const MEETING_APP_OBSERVER_SCHEDULER_STATE_SCHEMA: 'meeting_app_observer_scheduler_state';
export const MEETING_APP_OBSERVER_SCHEDULER_SCHEMA_VERSION: 1;

export interface MeetingAppObserverSchedulerConfig {
  type: 'meeting_app_observer_scheduler_config';
  schema: 'meeting_app_observer_scheduler_config';
  version: 1;
  platform: string;
  surface: string;
  display_name?: string;
  plan_schema: string;
  sdk_ready: boolean;
  preflight_status: string;
  runtime_factory?: string;
  observe_method: string;
  track_observe_method: string;
  timestamp_field: string;
  input_contract?: Record<string, unknown>;
  cadence: Record<string, number>;
  trigger_names: string[];
  trigger_policy: Array<Record<string, unknown>>;
  trigger_policy_map: Record<string, Record<string, unknown>>;
  sample_options: Record<string, unknown>;
  track_sample_policy: Record<string, unknown>;
  plan: MeetingAppRuntimeObserverPlan;
  next_actions: string[];
}

export interface MeetingAppObserverSchedulerConfigMatrix {
  type: 'meeting_app_observer_scheduler_config_matrix';
  schema: 'meeting_app_observer_scheduler_config_matrix';
  version: 1;
  platform_count: number;
  sdk_ready_count: number;
  track_enabled_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  configs: MeetingAppObserverSchedulerConfig[];
  observer_plan_matrix: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingAppObserverSchedulerState {
  type: 'meeting_app_observer_scheduler_state';
  schema: 'meeting_app_observer_scheduler_state';
  version: 1;
  running: boolean;
  trigger_count: number;
  sample_count: number;
  track_sample_count: number;
  scheduled_count: number;
  pending_timer_count: number;
  last_trigger: string | null;
  last_sample_at_ms: number | null;
  last_result: unknown;
  last_track_result: unknown;
  last_error: string | null;
}

export interface MeetingAppObserverSchedulerRuntime {
  sample?: (inputOrOptions?: Record<string, unknown>, options?: Record<string, unknown>) => unknown | Promise<unknown>;
  tick?: (inputOrOptions?: Record<string, unknown>, options?: Record<string, unknown>) => unknown | Promise<unknown>;
  sampleTracks?: (inputOrOptions?: Record<string, unknown>, options?: Record<string, unknown>) => unknown | Promise<unknown>;
  tickTracks?: (inputOrOptions?: Record<string, unknown>, options?: Record<string, unknown>) => unknown | Promise<unknown>;
  inputProvider?: (...args: unknown[]) => unknown;
  [key: string]: unknown;
}

export interface MeetingAppObserverSchedulerOptions {
  config?: MeetingAppObserverSchedulerConfig;
  observeTracks?: boolean;
  observe_tracks?: boolean;
  trackTriggers?: string[];
  track_triggers?: string[];
  runtimeInputMode?: 'explicit' | 'provider' | string;
  runtime_input_mode?: 'explicit' | 'provider' | string;
  input?: Record<string, unknown>;
  inputProvider?: () => Record<string, unknown> | Promise<Record<string, unknown>>;
  input_provider?: () => Record<string, unknown> | Promise<Record<string, unknown>>;
  setTimeout?: typeof setTimeout;
  clearTimeout?: typeof clearTimeout;
  setInterval?: typeof setInterval;
  clearInterval?: typeof clearInterval;
  now?: (() => number | string | Date) | number | string | Date;
  nowMs?: number | string | Date;
  now_ms?: number | string | Date;
  initialState?: Record<string, unknown>;
  initial_state?: Record<string, unknown>;
  [key: string]: unknown;
}

export function normalizeMeetingAppObserverTrigger(trigger?: string): string;

export function buildMeetingAppObserverSchedulerConfig(
  planOrPlatform?: MeetingAppRuntimeObserverPlan | string | Record<string, unknown>,
  options?: MeetingAppObserverSchedulerOptions,
): MeetingAppObserverSchedulerConfig;

export function buildMeetingAppObserverSchedulerConfigMatrix(
  options?: MeetingAppObserverSchedulerOptions,
): MeetingAppObserverSchedulerConfigMatrix;

export function createMeetingAppObserverScheduler(
  runtime: MeetingAppObserverSchedulerRuntime,
  planOrPlatform?: MeetingAppRuntimeObserverPlan | string | Record<string, unknown>,
  options?: MeetingAppObserverSchedulerOptions,
): {
  config: MeetingAppObserverSchedulerConfig;
  sample(trigger?: string, input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<Record<string, unknown>>;
  schedule(trigger?: string, input?: Record<string, unknown>, options?: Record<string, unknown>): Record<string, unknown>;
  clearScheduled(id?: string | null): MeetingAppObserverSchedulerState;
  handleTrigger(trigger?: string, input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<Record<string, unknown>> | Record<string, unknown>;
  triggerChanged(input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<Record<string, unknown>> | Record<string, unknown>;
  triggerKeepAlive(input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<Record<string, unknown>> | Record<string, unknown>;
  triggerSpeakerCandidate(input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<Record<string, unknown>> | Record<string, unknown>;
  triggerCandidateMissing(input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<Record<string, unknown>> | Record<string, unknown>;
  start(inputProvider?: (() => Record<string, unknown> | Promise<Record<string, unknown>>) | null, options?: Record<string, unknown>): MeetingAppObserverSchedulerState;
  stop(): MeetingAppObserverSchedulerState;
  getState(): MeetingAppObserverSchedulerState;
  reset(nextState?: Record<string, unknown>): MeetingAppObserverSchedulerState;
};
