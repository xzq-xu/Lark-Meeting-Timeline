import type { MeetingPlatformAdapterInstallManifest } from './platform-adapter-install-manifest.mjs';
import type {
  MeetingPlatformAdapterRuntimeManifest,
  MeetingPlatformAdapterRuntimeRecipeOptions,
  MeetingPlatformAdapterRuntimeTarget,
} from './platform-adapter-runtime-recipe.mjs';
import type {
  MeetingPlatformAdapterCandidateLaunchPlan,
  MeetingPlatformAdapterLaunchPlan,
  MeetingPlatformAdapterLaunchPlanOptions,
} from './platform-adapter-launch-plan.mjs';
import type {
  MeetingPlatformAdapterSession,
  MeetingPlatformAdapterSessionClient,
  MeetingPlatformAdapterSessionEvent,
  MeetingPlatformAdapterSessionOptions,
} from './platform-adapter-session.mjs';

export const MEETING_PLATFORM_ADAPTER_RUNNER_SCHEMA: 'meeting_platform_adapter_runner';
export const MEETING_PLATFORM_ADAPTER_OPEN_SESSION_EVENT_SCHEMA: 'meeting_platform_adapter_open_session_event';
export const MEETING_PLATFORM_ADAPTER_RUNNER_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdapterRunnerOptions extends MeetingPlatformAdapterLaunchPlanOptions, MeetingPlatformAdapterRuntimeRecipeOptions, MeetingPlatformAdapterSessionOptions {
  client?: MeetingPlatformAdapterSessionClient;
  sdk?: MeetingPlatformAdapterSessionClient;
  installManifest?: MeetingPlatformAdapterInstallManifest;
  install_manifest?: MeetingPlatformAdapterInstallManifest;
  runtimeManifest?: MeetingPlatformAdapterRuntimeManifest;
  runtime_manifest?: MeetingPlatformAdapterRuntimeManifest;
  runtimeTarget?: MeetingPlatformAdapterRuntimeTarget;
  runtime_target?: MeetingPlatformAdapterRuntimeTarget;
  manifest?: MeetingPlatformAdapterInstallManifest;
  runnerId?: string;
  runner_id?: string;
  observe?: boolean;
  observeAxis?: boolean;
  observe_axis?: boolean;
  autoObserve?: boolean;
  auto_observe?: boolean;
  validateRawSignal?: boolean;
  validate_raw_signal?: boolean;
  autoValidateRawSignal?: boolean;
  auto_validate_raw_signal?: boolean;
  readAdapterSelection?: boolean;
  read_adapter_selection?: boolean;
  autoReadAdapterSelection?: boolean;
  auto_read_adapter_selection?: boolean;
  [key: string]: unknown;
}

export interface MeetingPlatformAdapterOpenSessionEvent {
  type: 'meeting_platform_adapter_open_session_event';
  schema: 'meeting_platform_adapter_open_session_event';
  schema_version: 1;
  runner_id: string;
  action: 'open_session' | string;
  platform?: string;
  selected_surface?: string;
  plan_kind?: 'launch_plan' | 'candidate_launch_plan' | 'runtime_target';
  payload: {
    launch_plan?: MeetingPlatformAdapterLaunchPlan;
    candidate_launch_plan?: MeetingPlatformAdapterCandidateLaunchPlan;
    runtime_target?: MeetingPlatformAdapterRuntimeTarget;
    session: {
      schema: 'meeting_platform_adapter_session';
      id: string;
      plan_kind?: 'launch_plan' | 'candidate_launch_plan' | 'runtime_target' | 'unknown';
      platform?: string;
      selected_surface?: string;
      host_kind?: string;
      bridge_kind?: string;
      selected_evidence?: Record<string, unknown>;
    };
    adapter_selection_event?: MeetingPlatformAdapterSessionEvent;
    raw_signal_event?: MeetingPlatformAdapterSessionEvent;
    observe_event?: MeetingPlatformAdapterSessionEvent;
  };
  result?: unknown;
  captured_at_ms?: number;
}

export interface MeetingPlatformAdapterRunner {
  type: 'meeting_platform_adapter_runner';
  schema: 'meeting_platform_adapter_runner';
  schema_version: 1;
  id: string;
  install_manifest?: MeetingPlatformAdapterInstallManifest;
  runtime_manifest?: MeetingPlatformAdapterRuntimeManifest;
  getState(): Record<string, unknown>;
  runtimeTarget(input?: MeetingPlatformAdapterRuntimeTarget | string | URL | Record<string, unknown>, options?: MeetingPlatformAdapterRunnerOptions): MeetingPlatformAdapterRuntimeTarget;
  launchPlan(input?: MeetingPlatformAdapterLaunchPlanOptions | MeetingPlatformAdapterLaunchPlan | string | Record<string, unknown>, options?: MeetingPlatformAdapterRunnerOptions): MeetingPlatformAdapterLaunchPlan;
  candidateLaunchPlan(input?: MeetingPlatformAdapterLaunchPlanOptions | MeetingPlatformAdapterCandidateLaunchPlan | string | Record<string, unknown>, options?: MeetingPlatformAdapterRunnerOptions): MeetingPlatformAdapterCandidateLaunchPlan;
  open(input?: MeetingPlatformAdapterLaunchPlanOptions | MeetingPlatformAdapterLaunchPlan | MeetingPlatformAdapterCandidateLaunchPlan | MeetingPlatformAdapterRuntimeTarget | string | URL | Record<string, unknown>, options?: MeetingPlatformAdapterRunnerOptions): Promise<MeetingPlatformAdapterOpenSessionEvent>;
  openCandidate(input?: MeetingPlatformAdapterLaunchPlanOptions | MeetingPlatformAdapterCandidateLaunchPlan | string | Record<string, unknown>, options?: MeetingPlatformAdapterRunnerOptions): Promise<MeetingPlatformAdapterOpenSessionEvent>;
  ensureOpen(input?: MeetingPlatformAdapterLaunchPlanOptions | MeetingPlatformAdapterLaunchPlan | MeetingPlatformAdapterCandidateLaunchPlan | MeetingPlatformAdapterRuntimeTarget | string | URL | Record<string, unknown>, options?: MeetingPlatformAdapterRunnerOptions): Promise<MeetingPlatformAdapterOpenSessionEvent | undefined>;
  currentSession(): MeetingPlatformAdapterSession;
  readAdapterSelection(input?: Record<string, unknown>, options?: MeetingPlatformAdapterRunnerOptions): Promise<MeetingPlatformAdapterSessionEvent>;
  validateRawSignal(input?: Record<string, unknown>, options?: MeetingPlatformAdapterRunnerOptions): Promise<MeetingPlatformAdapterSessionEvent>;
  observeAxis(input?: Record<string, unknown>, options?: MeetingPlatformAdapterRunnerOptions): Promise<MeetingPlatformAdapterSessionEvent>;
  insertAnnotation(input?: Record<string, unknown>, options?: MeetingPlatformAdapterRunnerOptions): Promise<MeetingPlatformAdapterSessionEvent>;
  insertMark(input?: Record<string, unknown>, options?: MeetingPlatformAdapterRunnerOptions): Promise<MeetingPlatformAdapterSessionEvent>;
  speakerTrack(input?: Record<string, unknown>, options?: MeetingPlatformAdapterRunnerOptions): Promise<MeetingPlatformAdapterSessionEvent>;
  participantTrack(input?: Record<string, unknown>, options?: MeetingPlatformAdapterRunnerOptions): Promise<MeetingPlatformAdapterSessionEvent>;
  providerReconcile(input?: Record<string, unknown>, options?: MeetingPlatformAdapterRunnerOptions): Promise<MeetingPlatformAdapterSessionEvent>;
  reset(): Record<string, unknown>;
}

export interface MeetingPlatformAdapterRunnerHandoff {
  type: 'meeting_platform_adapter_runner_handoff';
  schema: 'meeting_platform_adapter_runner_handoff';
  schema_version: 1;
  runner_factory: 'createMeetingPlatformAdapterRunner';
  convenience_method: 'openMeetingPlatformAdapterSession' | 'openMeetingPlatformAdapterRuntimeSession';
  install_manifest_schema?: string;
  runtime_manifest_schema?: string;
  runtime_target_schema?: string;
  required_client_methods: string[];
  optional_client_methods: string[];
  runtime_sequence: string[];
  next_actions: string[];
}

export function createMeetingPlatformAdapterRunner(
  manifestOrInput?: MeetingPlatformAdapterInstallManifest | MeetingPlatformAdapterRuntimeManifest | MeetingPlatformAdapterRuntimeTarget | MeetingPlatformAdapterLaunchPlan | MeetingPlatformAdapterCandidateLaunchPlan | Record<string, unknown>,
  clientOrOptions?: MeetingPlatformAdapterSessionClient | MeetingPlatformAdapterRunnerOptions,
  options?: MeetingPlatformAdapterRunnerOptions,
): MeetingPlatformAdapterRunner;

export function openMeetingPlatformAdapterSession(
  manifestOrInput?: MeetingPlatformAdapterInstallManifest | MeetingPlatformAdapterRuntimeManifest | MeetingPlatformAdapterRuntimeTarget | MeetingPlatformAdapterLaunchPlan | MeetingPlatformAdapterCandidateLaunchPlan | Record<string, unknown>,
  clientOrOptions?: MeetingPlatformAdapterSessionClient | MeetingPlatformAdapterRunnerOptions,
  input?: MeetingPlatformAdapterLaunchPlanOptions | MeetingPlatformAdapterLaunchPlan | MeetingPlatformAdapterCandidateLaunchPlan | MeetingPlatformAdapterRuntimeTarget | string | URL | Record<string, unknown>,
  options?: MeetingPlatformAdapterRunnerOptions,
): Promise<MeetingPlatformAdapterOpenSessionEvent>;

export function openMeetingPlatformAdapterRuntimeSession(
  runtimeManifestOrTarget?: MeetingPlatformAdapterRuntimeManifest | MeetingPlatformAdapterRuntimeTarget | Record<string, unknown>,
  clientOrOptions?: MeetingPlatformAdapterSessionClient | MeetingPlatformAdapterRunnerOptions,
  input?: MeetingPlatformAdapterRuntimeTarget | string | URL | Record<string, unknown>,
  options?: MeetingPlatformAdapterRunnerOptions,
): Promise<MeetingPlatformAdapterOpenSessionEvent>;

export function buildMeetingPlatformAdapterRunnerHandoff(
  manifestOrInput?: MeetingPlatformAdapterInstallManifest | MeetingPlatformAdapterRuntimeManifest | MeetingPlatformAdapterRuntimeTarget | MeetingPlatformAdapterLaunchPlan | MeetingPlatformAdapterCandidateLaunchPlan | Record<string, unknown>,
  options?: MeetingPlatformAdapterRunnerOptions,
): MeetingPlatformAdapterRunnerHandoff;
