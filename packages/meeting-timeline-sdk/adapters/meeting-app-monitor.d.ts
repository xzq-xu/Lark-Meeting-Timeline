import type { MeetingAppDomCaptureInput, MeetingAppDomCaptureOptions, MeetingAppDomCaptureSnapshot } from './meeting-app-capture.mjs';
import type { MeetingAppObserverOptions } from './meeting-apps.mjs';
import type { MeetingSourceResult } from './meeting-source.mjs';

export interface MeetingAppMonitorOptions {
  initialState?: Record<string, unknown>;
  initial_state?: Record<string, unknown>;
  now?: (() => number | string | Date) | number | string | Date;
  nowMs?: number | string | Date;
  now_ms?: number | string | Date;
  sampleIntervalMs?: number;
  sample_interval_ms?: number;
  minObserveIntervalMs?: number;
  min_observe_interval_ms?: number;
  changedObserveEveryMs?: number;
  changed_observe_every_ms?: number;
  unchangedObserveEveryMs?: number;
  unchanged_observe_every_ms?: number;
  immediate?: boolean;
  input?: MeetingAppDomCaptureInput;
  inputProvider?: () => MeetingAppDomCaptureInput | Promise<MeetingAppDomCaptureInput>;
  input_provider?: () => MeetingAppDomCaptureInput | Promise<MeetingAppDomCaptureInput>;
  captureOptions?: MeetingAppDomCaptureOptions;
  capture_options?: MeetingAppDomCaptureOptions;
  observeOptions?: MeetingAppObserverOptions;
  observe_options?: MeetingAppObserverOptions;
  [key: string]: unknown;
}

export interface MeetingAppMonitorState {
  running: boolean;
  sampleCount: number;
  emitCount: number;
  skippedCount: number;
  lastObservedAtMs: number | null;
  lastEmittedAtMs: number | null;
  lastSignature: string | null;
  lastSnapshot: MeetingAppDomCaptureSnapshot | null;
  lastResult: unknown;
  lastSkipReason: string | null;
}

export interface MeetingAppMonitorSampleResult {
  emitted: boolean;
  skipped: boolean;
  reason: string;
  signature: string;
  snapshot: MeetingAppDomCaptureSnapshot;
  result?: MeetingSourceResult | Record<string, unknown>;
  state: MeetingAppMonitorState;
}

export interface MeetingAppMonitorTarget {
  observeMeetingApp?: (snapshot: MeetingAppDomCaptureSnapshot, options?: MeetingAppObserverOptions) => unknown | Promise<unknown>;
  observeApp?: (snapshot: MeetingAppDomCaptureSnapshot, options?: MeetingAppObserverOptions) => unknown | Promise<unknown>;
  observe?: (snapshot: MeetingAppDomCaptureSnapshot, options?: MeetingAppObserverOptions) => unknown | Promise<unknown>;
}

export function meetingAppDomSnapshotSignature(snapshot?: Record<string, unknown>): string;

export function createMeetingAppDomMonitor(
  target: MeetingAppMonitorTarget,
  options?: MeetingAppMonitorOptions,
): {
  sample(input?: MeetingAppDomCaptureInput, options?: MeetingAppMonitorOptions): Promise<MeetingAppMonitorSampleResult>;
  tick(input?: MeetingAppDomCaptureInput, options?: MeetingAppMonitorOptions): Promise<MeetingAppMonitorSampleResult>;
  start(
    inputProvider?: MeetingAppDomCaptureInput | (() => MeetingAppDomCaptureInput | Promise<MeetingAppDomCaptureInput>) | null,
    options?: MeetingAppMonitorOptions,
  ): MeetingAppMonitorState;
  stop(): MeetingAppMonitorState;
  getState(): MeetingAppMonitorState;
  reset(nextState?: Record<string, unknown> | null): MeetingAppMonitorState;
};
