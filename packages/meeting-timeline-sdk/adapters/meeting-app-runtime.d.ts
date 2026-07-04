import type { MeetingTimelineClient, MeetingTimelineClientOptions } from '../index.mjs';
import type { BrowserMeetingObserverOptions, BrowserMeetingSnapshot } from './browser-meeting.mjs';
import type { NormalizedMeetingSignal } from './core.mjs';
import type { LocalObserverOptions, LocalObserverSnapshot } from './local-observer.mjs';
import type { MeetingAppDomCaptureInput } from './meeting-app-capture.mjs';
import type {
  MeetingAppMonitorOptions,
  MeetingAppMonitorSampleResult,
  MeetingAppMonitorState,
  MeetingAppMonitorTarget,
} from './meeting-app-monitor.mjs';
import type { MeetingAppObserverOptions, MeetingAppSnapshot } from './meeting-apps.mjs';
import type { MeetingSourceAggregatorOptions, MeetingSourceResult } from './meeting-source.mjs';
import type {
  MeetingAppTrackRuntime,
  MeetingAppTrackRuntimeObservation,
  MeetingAppTrackRuntimeOptions,
  MeetingAppTrackRuntimePreview,
} from './meeting-app-track-runtime.mjs';
import type { NativeMeetingObserverOptions, NativeMeetingSnapshot } from './native-meeting.mjs';

export interface MeetingAppTimelineRuntimeOptions extends MeetingSourceAggregatorOptions, MeetingAppMonitorOptions {
  sources?: MeetingAppMonitorTarget & Record<string, unknown>;
  meetingSources?: MeetingAppMonitorTarget & Record<string, unknown>;
  meeting_sources?: MeetingAppMonitorTarget & Record<string, unknown>;
  sourceOptions?: MeetingSourceAggregatorOptions;
  source_options?: MeetingSourceAggregatorOptions;
  monitor?: {
    sample(input?: MeetingAppDomCaptureInput, options?: MeetingAppMonitorOptions): Promise<MeetingAppMonitorSampleResult>;
    tick(input?: MeetingAppDomCaptureInput, options?: MeetingAppMonitorOptions): Promise<MeetingAppMonitorSampleResult>;
    start(inputProvider?: unknown, options?: MeetingAppMonitorOptions): MeetingAppMonitorState;
    stop(): MeetingAppMonitorState;
    getState(): MeetingAppMonitorState;
    reset(nextState?: Record<string, unknown> | null): MeetingAppMonitorState;
  };
  monitorOptions?: MeetingAppMonitorOptions;
  monitor_options?: MeetingAppMonitorOptions;
  tracks?: MeetingAppTrackRuntime;
  trackRuntime?: MeetingAppTrackRuntime;
  track_runtime?: MeetingAppTrackRuntime;
  trackRuntimeOptions?: MeetingAppTrackRuntimeOptions;
  track_runtime_options?: MeetingAppTrackRuntimeOptions;
  trackMonitor?: MeetingAppTimelineRuntimeOptions['monitor'];
  track_monitor?: MeetingAppTimelineRuntimeOptions['monitor'];
  trackMonitorOptions?: MeetingAppMonitorOptions;
  track_monitor_options?: MeetingAppMonitorOptions;
}

export function createMeetingAppTimelineRuntime(
  clientOrOptions: MeetingTimelineClient | MeetingTimelineClientOptions | MeetingAppTimelineRuntimeOptions,
  options?: MeetingAppTimelineRuntimeOptions,
): {
  client: MeetingTimelineClient;
  sources: Record<string, unknown>;
  monitor: NonNullable<MeetingAppTimelineRuntimeOptions['monitor']>;
  tracks: MeetingAppTrackRuntime;
  trackMonitor: NonNullable<MeetingAppTimelineRuntimeOptions['monitor']>;
  sample(input?: MeetingAppDomCaptureInput, options?: MeetingAppMonitorOptions): Promise<MeetingAppMonitorSampleResult>;
  tick(input?: MeetingAppDomCaptureInput, options?: MeetingAppMonitorOptions): Promise<MeetingAppMonitorSampleResult>;
  start(inputProvider?: unknown, options?: MeetingAppMonitorOptions): MeetingAppMonitorState;
  stop(): MeetingAppMonitorState;
  sampleTracks(input?: MeetingAppDomCaptureInput, options?: MeetingAppMonitorOptions): Promise<MeetingAppMonitorSampleResult>;
  tickTracks(input?: MeetingAppDomCaptureInput, options?: MeetingAppMonitorOptions): Promise<MeetingAppMonitorSampleResult>;
  startTracks(inputProvider?: unknown, options?: MeetingAppMonitorOptions): MeetingAppMonitorState;
  stopTracks(): MeetingAppMonitorState;
  observeMeetingAppTracks(input?: MeetingAppSnapshot | MeetingAppSnapshot[], options?: MeetingAppTrackRuntimeOptions): Promise<MeetingAppTrackRuntimeObservation>;
  previewMeetingAppTracks(input?: MeetingAppSnapshot | MeetingAppSnapshot[], options?: MeetingAppTrackRuntimeOptions): MeetingAppTrackRuntimePreview;
  observeMeetingApp(input?: MeetingAppSnapshot, options?: MeetingAppObserverOptions): Promise<MeetingSourceResult>;
  observeApp(input?: MeetingAppSnapshot, options?: MeetingAppObserverOptions): Promise<MeetingSourceResult>;
  observeBrowser(input?: BrowserMeetingSnapshot, options?: BrowserMeetingObserverOptions): Promise<MeetingSourceResult>;
  observeNative(input?: NativeMeetingSnapshot, options?: NativeMeetingObserverOptions): Promise<MeetingSourceResult>;
  observeLocal(snapshot?: LocalObserverSnapshot, options?: LocalObserverOptions): Promise<MeetingSourceResult>;
  ingestProvider(platformOrInput: string | Record<string, unknown>, payload?: unknown, options?: Record<string, unknown>): Promise<MeetingSourceResult>;
  ingestSignals(signals?: NormalizedMeetingSignal[] | NormalizedMeetingSignal, options?: Record<string, unknown>): Promise<MeetingSourceResult>;
  insertMark(input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  insertAnnotation(input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  insertMarks(inputs?: Record<string, unknown>[], options?: Record<string, unknown>): Promise<unknown>;
  importTranscript(input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  startMeeting(input?: Record<string, unknown>): Promise<unknown>;
  endMeeting(input?: Record<string, unknown>): Promise<unknown>;
  getState(): Record<string, unknown>;
  reset(nextState?: Record<string, unknown>): Record<string, unknown>;
};
