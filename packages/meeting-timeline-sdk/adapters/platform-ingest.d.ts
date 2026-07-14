import type { MeetingTimelineClient } from '../index.mjs';
import type { MeetingPlatformEventAdapter } from './platform-registry.mjs';
import type { NormalizedMeetingSignal, ApplyMeetingSignalOptions, ApplyMeetingSignalResult } from './core.mjs';
import type {
  MeetingSignalReconciliationResult,
  MeetingSignalReconcilerOptions,
  MeetingSignalReconcilerState,
} from './signal-reconciler.mjs';

export const MEETING_PLATFORM_PROVIDER_REPLAY_REPORT_SCHEMA: 'meeting_platform_provider_replay_report';
export const MEETING_PLATFORM_PROVIDER_REPLAY_MATRIX_SCHEMA: 'meeting_platform_provider_replay_matrix';
export const MEETING_PLATFORM_PROVIDER_REPLAY_ACCEPTANCE_SCHEMA: 'meeting_platform_provider_replay_acceptance_report';
export const MEETING_PLATFORM_PROVIDER_REPLAY_SCHEMA_VERSION: 1;

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

export interface MeetingPlatformProviderReplayRecord {
  id?: string;
  kind?: string;
  required?: boolean;
  received_at_ms?: number | string | Date;
  receivedAtMs?: number | string | Date;
  payload?: unknown;
  body?: unknown;
  event?: unknown;
  raw?: unknown;
  provider_event?: unknown;
  providerEvent?: unknown;
  [key: string]: unknown;
}

export interface MeetingPlatformProviderReplayRow {
  id: string;
  kind?: string;
  required: boolean;
  accepted: boolean;
  platform?: string;
  source?: string;
  runtime_event_action?: string;
  runtime_event_schema?: string;
  runtime_event_platform?: string;
  sent_at_ms?: number;
  received_at_ms?: number | string | Date;
  signal_count: number;
  signal_types: string[];
  coverage?: PlatformEventDiagnosticResult['coverage'];
  actionable_for_axis: boolean;
  meetings?: Array<Record<string, unknown>>;
  signals?: Array<Record<string, unknown>>;
  issues: PlatformEventDiagnosticIssue[];
  runtime_error?: Record<string, unknown>;
}

export interface MeetingPlatformProviderReplayReport {
  type: 'meeting_platform_provider_replay_report';
  schema: 'meeting_platform_provider_replay_report';
  schema_version: 1;
  accepted: boolean;
  target: string;
  platform: string;
  source?: string;
  adapter?: Record<string, unknown>;
  record_count: number;
  accepted_record_count: number;
  runtime_event_count: number;
  signal_count: number;
  signal_types: string[];
  coverage: Record<string, boolean>;
  required_coverage: string[];
  runtime_contract: Record<string, unknown>;
  rows: MeetingPlatformProviderReplayRow[];
  issue_count: number;
  issues: string[];
  next_actions: string[];
}

export interface MeetingPlatformProviderReplayAcceptanceReport {
  type: 'meeting_platform_provider_replay_acceptance_report';
  schema: 'meeting_platform_provider_replay_acceptance_report';
  schema_version: 1;
  accepted: boolean;
  target?: string;
  platform?: string;
  record_count: number;
  accepted_record_count: number;
  runtime_event_count: number;
  signal_count: number;
  required_coverage: string[];
  coverage?: Record<string, boolean>;
  issue_count: number;
  issues: PlatformEventDiagnosticIssue[];
  next_actions: string[];
}

export interface MeetingPlatformProviderReplayMatrix {
  type: 'meeting_platform_provider_replay_matrix';
  schema: 'meeting_platform_provider_replay_matrix';
  schema_version: 1;
  accepted: boolean;
  target: string;
  platform_count: number;
  accepted_count: number;
  runtime_event_count: number;
  signal_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  reports: MeetingPlatformProviderReplayReport[];
  issue_count: number;
  issues: string[];
  next_actions: string[];
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

export function sampleMeetingPlatformProviderEvents(
  platform: string,
  options?: Record<string, unknown>,
): MeetingPlatformProviderReplayRecord[];

export function buildMeetingPlatformProviderReplayReport(
  platformOrInput?: string | (PlatformEventIngestInput & {
    records?: MeetingPlatformProviderReplayRecord[];
    events?: MeetingPlatformProviderReplayRecord[];
    samples?: MeetingPlatformProviderReplayRecord[];
    provider_events?: MeetingPlatformProviderReplayRecord[];
    providerEvents?: MeetingPlatformProviderReplayRecord[];
  }),
  recordsOrOptions?: MeetingPlatformProviderReplayRecord[] | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingPlatformProviderReplayReport;

export function buildMeetingPlatformProviderReplayAcceptanceReport(
  reportOrInput?: MeetingPlatformProviderReplayReport | string | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingPlatformProviderReplayAcceptanceReport;

export function buildMeetingPlatformProviderReplayMatrix(
  options?: Record<string, unknown> & {
    platforms?: string[];
    platform_keys?: string[];
    recordsByPlatform?: Record<string, MeetingPlatformProviderReplayRecord[]>;
    records_by_platform?: Record<string, MeetingPlatformProviderReplayRecord[]>;
    providerEventsByPlatform?: Record<string, MeetingPlatformProviderReplayRecord[]>;
    provider_events_by_platform?: Record<string, MeetingPlatformProviderReplayRecord[]>;
  },
): MeetingPlatformProviderReplayMatrix;

export function assertMeetingPlatformProviderReplayReport(
  reportOrInput?: MeetingPlatformProviderReplayReport | string | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingPlatformProviderReplayReport;

export function assertMeetingPlatformProviderReplayMatrix(
  matrixOrOptions?: MeetingPlatformProviderReplayMatrix | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingPlatformProviderReplayMatrix;

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
