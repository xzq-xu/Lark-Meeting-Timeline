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

export function normalizePlatformEvent(
  platformOrInput: string | PlatformEventIngestInput,
  payload?: unknown,
  options?: PlatformEventIngestOptions,
): NormalizedPlatformEvent;

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
