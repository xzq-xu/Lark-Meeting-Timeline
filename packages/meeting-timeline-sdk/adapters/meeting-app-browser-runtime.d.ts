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
  dispose(): MeetingAppMonitorState;
  observeMeetingApp(input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<MeetingSourceResult>;
  insertMark(input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  insertMarks(inputs?: Record<string, unknown>[], options?: Record<string, unknown>): Promise<unknown>;
  ingestProvider(platformOrInput: string | Record<string, unknown>, payload?: unknown, options?: Record<string, unknown>): Promise<MeetingSourceResult>;
  getState(): Record<string, unknown>;
  reset(nextState?: Record<string, unknown>): Record<string, unknown>;
};
