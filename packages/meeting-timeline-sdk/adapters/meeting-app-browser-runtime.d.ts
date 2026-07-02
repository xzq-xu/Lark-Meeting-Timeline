import type { MeetingTimelineClient, MeetingTimelineClientOptions } from '../index.mjs';
import type { MeetingAppDomCaptureInput } from './meeting-app-capture.mjs';
import type {
  MeetingAppTimelineRuntimeOptions,
} from './meeting-app-runtime.mjs';
import type { MeetingAppMonitorOptions, MeetingAppMonitorSampleResult, MeetingAppMonitorState } from './meeting-app-monitor.mjs';
import type { MeetingSourceResult } from './meeting-source.mjs';

export interface MeetingAppBrowserRuntimeOptions extends MeetingAppTimelineRuntimeOptions {
  window?: Window | Record<string, unknown>;
  win?: Window | Record<string, unknown>;
  document?: Document | Record<string, unknown>;
  doc?: Document | Record<string, unknown>;
  location?: Location | string | Record<string, unknown>;
  navigator?: Navigator | Record<string, unknown>;
  eventTarget?: EventTarget | Record<string, unknown>;
  event_target?: EventTarget | Record<string, unknown>;
  stopEvents?: string[];
  stop_events?: string[];
  browserName?: string;
  browser_name?: string;
  mutationObserver?: boolean | unknown;
  mutation_observer?: boolean | unknown;
  observeMutations?: boolean;
  observe_mutations?: boolean;
  MutationObserver?: unknown;
  mutationObserverCtor?: unknown;
  mutation_observer_ctor?: unknown;
  mutationRoot?: unknown;
  mutation_root?: unknown;
  mutationObserverOptions?: Record<string, unknown>;
  mutation_observer_options?: Record<string, unknown>;
  debounceMs?: number;
  debounce_ms?: number;
  mutationDebounceMs?: number;
  mutation_debounce_ms?: number;
  speakerStableFollowup?: boolean;
  speaker_stable_followup?: boolean;
  mutationSpeakerFollowup?: boolean;
  mutation_speaker_followup?: boolean;
  speakerStableFollowupMs?: number;
  speaker_stable_followup_ms?: number;
  mutationStableFollowupMs?: number;
  mutation_stable_followup_ms?: number;
  keepMutationObserverOnStop?: boolean;
  keep_mutation_observer_on_stop?: boolean;
  [key: string]: unknown;
}

export interface MeetingAppBrowserRuntimeMessage {
  type?: string;
  action?: string;
  kind?: string;
  payload?: Record<string, unknown>;
  data?: Record<string, unknown>;
  [key: string]: unknown;
}

export function meetingAppBrowserInput(options?: MeetingAppBrowserRuntimeOptions): MeetingAppDomCaptureInput;

export function createMeetingAppBrowserRuntime(
  clientOrOptions: MeetingTimelineClient | MeetingTimelineClientOptions | MeetingAppBrowserRuntimeOptions,
  options?: MeetingAppBrowserRuntimeOptions,
): {
  client: MeetingTimelineClient;
  sources: Record<string, unknown>;
  monitor: Record<string, unknown>;
  inputProvider(extra?: MeetingAppBrowserRuntimeOptions): MeetingAppDomCaptureInput;
  start(options?: MeetingAppMonitorOptions & {
    lifecycle?: boolean;
    installLifecycleHandlers?: boolean;
    install_lifecycle_handlers?: boolean;
  }): MeetingAppMonitorState;
  stop(): MeetingAppMonitorState;
  sample(options?: MeetingAppMonitorOptions): Promise<MeetingAppMonitorSampleResult>;
  tick(options?: MeetingAppMonitorOptions): Promise<MeetingAppMonitorSampleResult>;
  handleMessage(message?: MeetingAppBrowserRuntimeMessage, options?: Record<string, unknown>): Promise<{
    handled: boolean;
    action?: string;
    reason?: string;
    type?: string;
    result?: unknown;
  }>;
  installLifecycleHandlers(options?: MeetingAppBrowserRuntimeOptions): Record<string, unknown>;
  removeLifecycleHandlers(): Record<string, unknown>;
  installMutationObserver(options?: MeetingAppBrowserRuntimeOptions): Record<string, unknown>;
  removeMutationObserver(): Record<string, unknown>;
  flushMutationObserver(options?: MeetingAppMonitorOptions & {
    force?: boolean;
    sampleOptions?: MeetingAppMonitorOptions;
    sample_options?: MeetingAppMonitorOptions;
  }): Promise<Record<string, unknown>>;
  dispose(): MeetingAppMonitorState;
  observeMeetingApp(input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<MeetingSourceResult>;
  insertMark(input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  insertMarks(inputs?: Record<string, unknown>[], options?: Record<string, unknown>): Promise<unknown>;
  ingestProvider(platformOrInput: string | Record<string, unknown>, payload?: unknown, options?: Record<string, unknown>): Promise<MeetingSourceResult>;
  getState(): Record<string, unknown>;
  reset(nextState?: Record<string, unknown>): Record<string, unknown>;
};
