import type { MeetingPlatformAdapterInstallManifest } from './platform-adapter-install-manifest.mjs';
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

export interface MeetingPlatformAdapterRunnerOptions extends MeetingPlatformAdapterLaunchPlanOptions, MeetingPlatformAdapterSessionOptions {
  client?: MeetingPlatformAdapterSessionClient;
  sdk?: MeetingPlatformAdapterSessionClient;
  installManifest?: MeetingPlatformAdapterInstallManifest;
  install_manifest?: MeetingPlatformAdapterInstallManifest;
  manifest?: MeetingPlatformAdapterInstallManifest;
  runnerId?: string;
  runner_id?: string;
  observe?: boolean;
  observeAxis?: boolean;
  observe_axis?: boolean;
  autoObserve?: boolean;
  auto_observe?: boolean;
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
  payload: {
    launch_plan: MeetingPlatformAdapterLaunchPlan;
    session: {
      schema: 'meeting_platform_adapter_session';
      id: string;
      platform?: string;
      selected_surface?: string;
    };
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
  getState(): Record<string, unknown>;
  launchPlan(input?: MeetingPlatformAdapterLaunchPlanOptions | MeetingPlatformAdapterLaunchPlan | string | Record<string, unknown>, options?: MeetingPlatformAdapterRunnerOptions): MeetingPlatformAdapterLaunchPlan;
  candidateLaunchPlan(input?: MeetingPlatformAdapterLaunchPlanOptions | MeetingPlatformAdapterCandidateLaunchPlan | string | Record<string, unknown>, options?: MeetingPlatformAdapterRunnerOptions): MeetingPlatformAdapterCandidateLaunchPlan;
  open(input?: MeetingPlatformAdapterLaunchPlanOptions | MeetingPlatformAdapterLaunchPlan | MeetingPlatformAdapterCandidateLaunchPlan | string | Record<string, unknown>, options?: MeetingPlatformAdapterRunnerOptions): Promise<MeetingPlatformAdapterOpenSessionEvent>;
  openCandidate(input?: MeetingPlatformAdapterLaunchPlanOptions | MeetingPlatformAdapterCandidateLaunchPlan | string | Record<string, unknown>, options?: MeetingPlatformAdapterRunnerOptions): Promise<MeetingPlatformAdapterOpenSessionEvent>;
  ensureOpen(input?: MeetingPlatformAdapterLaunchPlanOptions | MeetingPlatformAdapterLaunchPlan | MeetingPlatformAdapterCandidateLaunchPlan | string | Record<string, unknown>, options?: MeetingPlatformAdapterRunnerOptions): Promise<MeetingPlatformAdapterOpenSessionEvent | undefined>;
  currentSession(): MeetingPlatformAdapterSession;
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
  convenience_method: 'openMeetingPlatformAdapterSession';
  install_manifest_schema?: string;
  required_client_methods: string[];
  optional_client_methods: string[];
  runtime_sequence: string[];
  next_actions: string[];
}

export function createMeetingPlatformAdapterRunner(
  manifestOrInput?: MeetingPlatformAdapterInstallManifest | MeetingPlatformAdapterLaunchPlan | Record<string, unknown>,
  clientOrOptions?: MeetingPlatformAdapterSessionClient | MeetingPlatformAdapterRunnerOptions,
  options?: MeetingPlatformAdapterRunnerOptions,
): MeetingPlatformAdapterRunner;

export function openMeetingPlatformAdapterSession(
  manifestOrInput?: MeetingPlatformAdapterInstallManifest | MeetingPlatformAdapterLaunchPlan | MeetingPlatformAdapterCandidateLaunchPlan | Record<string, unknown>,
  clientOrOptions?: MeetingPlatformAdapterSessionClient | MeetingPlatformAdapterRunnerOptions,
  input?: MeetingPlatformAdapterLaunchPlanOptions | MeetingPlatformAdapterLaunchPlan | MeetingPlatformAdapterCandidateLaunchPlan | string | Record<string, unknown>,
  options?: MeetingPlatformAdapterRunnerOptions,
): Promise<MeetingPlatformAdapterOpenSessionEvent>;

export function buildMeetingPlatformAdapterRunnerHandoff(
  manifestOrInput?: MeetingPlatformAdapterInstallManifest | MeetingPlatformAdapterLaunchPlan | Record<string, unknown>,
  options?: MeetingPlatformAdapterRunnerOptions,
): MeetingPlatformAdapterRunnerHandoff;
