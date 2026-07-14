import type { MeetingTimelineClient } from '../index.mjs';
import type { BrowserMeetingObserverOptions, BrowserMeetingSnapshot } from './browser-meeting.mjs';
import type { ApplyMeetingSignalOptions, ApplyMeetingSignalResult, NormalizedMeetingSignal } from './core.mjs';
import type { LocalObserverOptions, LocalObserverSnapshot } from './local-observer.mjs';
import type { MeetingAppObserverOptions, MeetingAppSnapshot } from './meeting-apps.mjs';
import type { NativeMeetingObserverOptions, NativeMeetingSnapshot } from './native-meeting.mjs';

export interface MeetingSourceAggregatorOptions {
  baseUrl?: string;
  base_url?: string;
  clientOptions?: Record<string, unknown>;
  client_options?: Record<string, unknown>;
  source?: string;
  applyOptions?: ApplyMeetingSignalOptions;
  apply_options?: ApplyMeetingSignalOptions;
  reconcileOptions?: Record<string, unknown>;
  reconcile_options?: Record<string, unknown>;
  browserOptions?: BrowserMeetingObserverOptions;
  browser_options?: BrowserMeetingObserverOptions;
  appOptions?: MeetingAppObserverOptions;
  app_options?: MeetingAppObserverOptions;
  meetingAppOptions?: MeetingAppObserverOptions;
  meeting_app_options?: MeetingAppObserverOptions;
  nativeOptions?: NativeMeetingObserverOptions;
  native_options?: NativeMeetingObserverOptions;
  localOptions?: LocalObserverOptions;
  local_options?: LocalObserverOptions;
  providerOptions?: Record<string, unknown>;
  provider_options?: Record<string, unknown>;
  speakerOptions?: Record<string, unknown>;
  speaker_options?: Record<string, unknown>;
  reconciler?: Record<string, unknown>;
  signalReconciler?: Record<string, unknown>;
  signal_reconciler?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingSourceDiagnostic {
  source: string;
  raw_signal_count: number;
  applied_signal_count: number;
  skipped_count: number;
  result_count: number;
  signals: Record<string, unknown>[];
  skipped?: Record<string, unknown>[];
}

export interface MeetingSourceResult {
  source: string;
  signals: NormalizedMeetingSignal[];
  rawSignals: NormalizedMeetingSignal[];
  reconciliation: {
    signals: NormalizedMeetingSignal[];
    skipped: Record<string, unknown>[];
    decisions: Record<string, unknown>[];
    state: Record<string, unknown>;
  };
  results: ApplyMeetingSignalResult[];
  diagnostic: MeetingSourceDiagnostic;
  [key: string]: unknown;
}

export function createMeetingSourceAggregator(
  clientOrOptions: MeetingTimelineClient | MeetingSourceAggregatorOptions,
  options?: MeetingSourceAggregatorOptions,
): {
  client: MeetingTimelineClient;
  observeBrowser(input?: BrowserMeetingSnapshot, options?: BrowserMeetingObserverOptions): Promise<MeetingSourceResult>;
  observeNative(input?: NativeMeetingSnapshot, options?: NativeMeetingObserverOptions): Promise<MeetingSourceResult>;
  observeMeetingApp(input?: MeetingAppSnapshot, options?: MeetingAppObserverOptions): Promise<MeetingSourceResult>;
  observeApp(input?: MeetingAppSnapshot, options?: MeetingAppObserverOptions): Promise<MeetingSourceResult>;
  observeLocal(snapshot?: LocalObserverSnapshot, options?: LocalObserverOptions): Promise<MeetingSourceResult>;
  observeLocalCandidates(candidates?: LocalObserverSnapshot[] | Record<string, unknown>, options?: LocalObserverOptions): Promise<MeetingSourceResult>;
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
