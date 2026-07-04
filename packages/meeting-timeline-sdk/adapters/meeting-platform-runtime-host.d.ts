import type { MeetingTimelineClient } from '../index.mjs';
import type {
  MeetingAppObserverSchedulerConfig,
  MeetingAppObserverSchedulerRuntime,
  MeetingAppObserverSchedulerState,
} from './meeting-app-observer-scheduler.mjs';
import type { MeetingAppRuntimeObserverPlan } from './meeting-app-profile.mjs';
import type {
  MeetingPlatformRuntimeBundle,
  MeetingPlatformRuntimeBundleOptions,
} from './platform-runtime-bundle.mjs';

export const MEETING_PLATFORM_RUNTIME_HOST_CONFIG_SCHEMA: 'meeting_platform_runtime_host_config';
export const MEETING_PLATFORM_RUNTIME_HOST_CONFIG_MATRIX_SCHEMA: 'meeting_platform_runtime_host_config_matrix';
export const MEETING_PLATFORM_RUNTIME_HOST_STATE_SCHEMA: 'meeting_platform_runtime_host_state';
export const MEETING_PLATFORM_RUNTIME_HOST_SCHEMA_VERSION: 1;

export interface MeetingPlatformRuntimeHostConfig {
  type: 'meeting_platform_runtime_host_config';
  schema: 'meeting_platform_runtime_host_config';
  version: 1;
  platform: string;
  display_name?: string;
  runtime_bundle_id: string;
  runtime_bundle_schema: string;
  runtime_bundle: MeetingPlatformRuntimeBundle;
  runtime_factory: string;
  runtime_options: Record<string, unknown>;
  observer_plan?: MeetingAppRuntimeObserverPlan | null;
  observer_scheduler_config?: MeetingAppObserverSchedulerConfig | null;
  driver: {
    change_observer: {
      enabled: boolean;
      trigger: string;
      schedule_by_default: boolean;
      observe_options: Record<string, unknown>;
    };
    keep_alive: {
      enabled: boolean;
      interval_ms: number;
    };
    lifecycle: {
      enabled: boolean;
      stop_events: string[];
    };
    runtime_start_default: boolean;
  };
  host_api: string[];
  readiness: {
    runtime_ready: boolean;
    observer_plan_ready: boolean;
    observer_scheduler_ready: boolean;
    host_ready: boolean;
    provider_required_for_realtime: boolean;
    transcript_blocks_realtime: boolean;
  };
  next_actions: string[];
}

export interface MeetingPlatformRuntimeHostConfigMatrix {
  type: 'meeting_platform_runtime_host_config_matrix';
  schema: 'meeting_platform_runtime_host_config_matrix';
  version: 1;
  platform_count: number;
  host_ready_count: number;
  change_observer_count: number;
  keep_alive_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  configs: MeetingPlatformRuntimeHostConfig[];
  runtime_bundle_matrix: Record<string, unknown>;
  observer_scheduler_config_matrix: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingPlatformRuntimeHostState {
  type: 'meeting_platform_runtime_host_state';
  schema: 'meeting_platform_runtime_host_state';
  version: 1;
  running: boolean;
  platform: string | null;
  change_observer_installed: boolean;
  lifecycle_installed: boolean;
  scheduled_trigger_count: number;
  last_trigger: string | null;
  last_error: string | null;
  scheduler?: MeetingAppObserverSchedulerState;
  runtime?: unknown;
}

export interface MeetingPlatformRuntimeHostOptions extends MeetingPlatformRuntimeBundleOptions {
  runtime?: MeetingAppObserverSchedulerRuntime;
  scheduler?: Record<string, unknown>;
  runtimeOptions?: Record<string, unknown>;
  runtime_options?: Record<string, unknown>;
  browserRuntimeOptions?: Record<string, unknown>;
  browser_runtime_options?: Record<string, unknown>;
  input?: Record<string, unknown>;
  snapshot?: Record<string, unknown>;
  inputProvider?: () => Record<string, unknown> | Promise<Record<string, unknown>>;
  input_provider?: () => Record<string, unknown> | Promise<Record<string, unknown>>;
  runtimeInputMode?: 'explicit' | 'provider' | string;
  runtime_input_mode?: 'explicit' | 'provider' | string;
  window?: unknown;
  win?: unknown;
  document?: unknown;
  doc?: unknown;
  eventTarget?: unknown;
  event_target?: unknown;
  MutationObserver?: unknown;
  mutationObserver?: unknown;
  mutation_observer?: unknown;
  mutationObserverCtor?: unknown;
  mutation_observer_ctor?: unknown;
  mutationRoot?: unknown;
  mutation_root?: unknown;
  mutationObserverOptions?: Record<string, unknown>;
  mutation_observer_options?: Record<string, unknown>;
  mutationFilter?: (record: unknown, options: Record<string, unknown>) => boolean;
  mutation_filter?: (record: unknown, options: Record<string, unknown>) => boolean;
  setTimeout?: typeof setTimeout;
  clearTimeout?: typeof clearTimeout;
  setInterval?: typeof setInterval;
  clearInterval?: typeof clearInterval;
  [key: string]: unknown;
}

export interface MeetingPlatformRuntimeHost {
  config: MeetingPlatformRuntimeHostConfig;
  runtime: MeetingAppObserverSchedulerRuntime;
  scheduler: Record<string, unknown>;
  inputProvider: (...args: unknown[]) => unknown;
  start(options?: Record<string, unknown>): MeetingPlatformRuntimeHostState;
  stop(): MeetingPlatformRuntimeHostState;
  dispose(): MeetingPlatformRuntimeHostState;
  installChangeObserver(options?: Record<string, unknown>): Record<string, unknown>;
  removeChangeObserver(): MeetingPlatformRuntimeHostState;
  installLifecycleHandlers(options?: Record<string, unknown>): Record<string, unknown>;
  removeLifecycleHandlers(): MeetingPlatformRuntimeHostState;
  changed(input?: Record<string, unknown>, options?: Record<string, unknown>): unknown;
  keepAlive(input?: Record<string, unknown>, options?: Record<string, unknown>): unknown;
  speakerCandidate(input?: Record<string, unknown>, options?: Record<string, unknown>): unknown;
  candidateMissing(input?: Record<string, unknown>, options?: Record<string, unknown>): unknown;
  sample(trigger?: string, input?: Record<string, unknown>, options?: Record<string, unknown>): unknown;
  getState(): MeetingPlatformRuntimeHostState;
  reset(nextState?: Record<string, unknown>): MeetingPlatformRuntimeHostState;
  scheduleTrigger(trigger?: string, input?: Record<string, unknown>, options?: Record<string, unknown>): unknown;
}

export function buildMeetingPlatformRuntimeHostConfig(
  bundleOrPlatform?: MeetingPlatformRuntimeBundle | string | Record<string, unknown>,
  options?: MeetingPlatformRuntimeHostOptions,
): MeetingPlatformRuntimeHostConfig;

export function buildMeetingPlatformRuntimeHostConfigMatrix(
  options?: MeetingPlatformRuntimeHostOptions,
): MeetingPlatformRuntimeHostConfigMatrix;

export function createMeetingPlatformRuntimeHost(
  clientOrRuntime: MeetingTimelineClient | MeetingAppObserverSchedulerRuntime | Record<string, unknown>,
  bundleOrPlatform?: MeetingPlatformRuntimeHostConfig | MeetingPlatformRuntimeBundle | string | Record<string, unknown>,
  options?: MeetingPlatformRuntimeHostOptions,
): MeetingPlatformRuntimeHost;
