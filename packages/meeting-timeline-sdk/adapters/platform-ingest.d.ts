import type { MeetingTimelineClient } from '../index.mjs';
import type { MeetingPlatformEventAdapter } from './platform-registry.mjs';
import type { NormalizedMeetingSignal, ApplyMeetingSignalOptions, ApplyMeetingSignalResult } from './core.mjs';
import type {
  MeetingSignalReconciliationResult,
  MeetingSignalReconcilerOptions,
  MeetingSignalReconcilerState,
} from './signal-reconciler.mjs';

export interface PlatformEventIngestInput {
  platform?: string;
  provider?: string;
  adapter?: string;
  payload?: unknown;
  body?: unknown;
  event?: unknown;
  raw?: unknown;
  receivedAtMs?: number | string | Date;
  received_at_ms?: number | string | Date;
  received_at?: number | string | Date;
  timestamp?: number | string | Date;
  ts?: number | string | Date;
  normalizerOptions?: Record<string, unknown>;
  normalizer_options?: Record<string, unknown>;
  applyOptions?: ApplyMeetingSignalOptions;
  apply_options?: ApplyMeetingSignalOptions;
  options?: PlatformEventIngestOptions;
}

export interface PlatformEventIngestOptions extends ApplyMeetingSignalOptions {
  receivedAtMs?: number | string | Date;
  received_at_ms?: number | string | Date;
  received_at?: number | string | Date;
  normalizerOptions?: Record<string, unknown>;
  normalizer_options?: Record<string, unknown>;
  applyOptions?: ApplyMeetingSignalOptions;
  apply_options?: ApplyMeetingSignalOptions;
  reconcileOptions?: MeetingSignalReconcilerOptions;
  reconcile_options?: MeetingSignalReconcilerOptions;
}

export interface NormalizedPlatformEvent {
  adapter: MeetingPlatformEventAdapter;
  platform: string;
  source: string;
  signals: NormalizedMeetingSignal[];
}

export interface PlatformEventIngestResult extends NormalizedPlatformEvent {
  results: ApplyMeetingSignalResult[];
}

export interface ReconciledPlatformEventIngestResult extends NormalizedPlatformEvent {
  rawSignals: NormalizedMeetingSignal[];
  reconciliation: MeetingSignalReconciliationResult;
  results: ApplyMeetingSignalResult[];
}

export interface PlatformEventDiagnosticIssue {
  severity: 'info' | 'warning' | 'error';
  code: string;
  message: string;
  [key: string]: unknown;
}

export interface PlatformEventDiagnosticResult {
  ok: boolean;
  supported: boolean;
  actionable: boolean;
  platform?: string;
  source?: string;
  adapter?: {
    key?: string;
    source?: string;
    aliases?: readonly string[];
  };
  signal_count?: number;
  signal_types?: string[];
  coverage?: {
    realtime_axis: boolean;
    meeting_start: boolean;
    meeting_end: boolean;
    participant_track: boolean;
    speaker_activity: boolean;
    artifact_ready: boolean;
    subscription_lifecycle: boolean;
  };
  meetings?: Array<Record<string, unknown>>;
  signals?: Array<Record<string, unknown>>;
  raw_signals?: NormalizedMeetingSignal[];
  issues: PlatformEventDiagnosticIssue[];
  error?: string;
  details?: Record<string, unknown>;
}

export function normalizePlatformEvent(
  platformOrInput: string | PlatformEventIngestInput,
  payload?: unknown,
  options?: PlatformEventIngestOptions,
): NormalizedPlatformEvent;

export function diagnosePlatformEvent(
  platformOrInput: string | PlatformEventIngestInput,
  payload?: unknown,
  options?: PlatformEventIngestOptions & {
    includeRawSignals?: boolean;
    include_raw_signals?: boolean;
  },
): PlatformEventDiagnosticResult;

export function ingestPlatformEvent(
  client: MeetingTimelineClient,
  platformOrInput: string | PlatformEventIngestInput,
  payload?: unknown,
  options?: PlatformEventIngestOptions,
): Promise<PlatformEventIngestResult>;

export function createReconciledPlatformEventIngestor(
  client: MeetingTimelineClient,
  options?: PlatformEventIngestOptions & {
    reconciler?: {
      reconcile(signals?: NormalizedMeetingSignal[], options?: MeetingSignalReconcilerOptions): MeetingSignalReconciliationResult;
      getState(): MeetingSignalReconcilerState;
      reset(nextState?: Partial<MeetingSignalReconcilerState>): MeetingSignalReconcilerState;
    };
    signalReconciler?: {
      reconcile(signals?: NormalizedMeetingSignal[], options?: MeetingSignalReconcilerOptions): MeetingSignalReconciliationResult;
      getState(): MeetingSignalReconcilerState;
      reset(nextState?: Partial<MeetingSignalReconcilerState>): MeetingSignalReconcilerState;
    };
    signal_reconciler?: {
      reconcile(signals?: NormalizedMeetingSignal[], options?: MeetingSignalReconcilerOptions): MeetingSignalReconciliationResult;
      getState(): MeetingSignalReconcilerState;
      reset(nextState?: Partial<MeetingSignalReconcilerState>): MeetingSignalReconcilerState;
    };
  },
): {
  ingest(
    platformOrInput: string | PlatformEventIngestInput,
    payload?: unknown,
    options?: PlatformEventIngestOptions,
  ): Promise<ReconciledPlatformEventIngestResult>;
  getState(): MeetingSignalReconcilerState;
  reset(nextState?: Partial<MeetingSignalReconcilerState>): MeetingSignalReconcilerState;
};
