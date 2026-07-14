import type {
  MeetingEndInput,
  MeetingStartInput,
  MeetingTimelineClient,
  MeetingTimelineClientOptions,
  TimelineMarkInput,
  TranscriptImportInput,
} from '../index.mjs';
import type {
  LocalObserverCandidateResult,
  LocalObserverOptions,
  LocalObserverResult,
  LocalObserverSnapshot,
  LocalObserverState,
  LocalTimelineObserverOptions,
} from './local-observer.mjs';
import type {
  PlatformEventIngestInput,
  PlatformEventIngestOptions,
  ReconciledPlatformEventIngestResult,
} from './platform-ingest.mjs';
import type { ApplyMeetingSignalResult, NormalizedMeetingSignal } from './core.mjs';
import type { MeetingSignalReconcilerState } from './signal-reconciler.mjs';

export interface MeetingTimelineBridgeOptions {
  source?: string;
  clientOptions?: MeetingTimelineClientOptions;
  client_options?: MeetingTimelineClientOptions;
  applyOptions?: Record<string, unknown>;
  apply_options?: Record<string, unknown>;
  localObserverOptions?: LocalTimelineObserverOptions;
  local_observer_options?: LocalTimelineObserverOptions;
  ingestOptions?: PlatformEventIngestOptions;
  ingest_options?: PlatformEventIngestOptions;
  [key: string]: unknown;
}

export interface MeetingTimelineBridgeState {
  local_observer: LocalObserverState;
  signal_reconciler: MeetingSignalReconcilerState;
}

export interface MeetingTimelineBridgeObserveResult extends LocalObserverResult {
  rawSignals?: NormalizedMeetingSignal[];
  results: ApplyMeetingSignalResult[];
  reconciliation?: ReconciledPlatformEventIngestResult['reconciliation'];
  ingestion?: ReconciledPlatformEventIngestResult;
}

export interface MeetingTimelineBridgeCandidateResult extends LocalObserverCandidateResult {
  rawSignals?: NormalizedMeetingSignal[];
  results: ApplyMeetingSignalResult[];
  reconciliation?: ReconciledPlatformEventIngestResult['reconciliation'];
  ingestion?: ReconciledPlatformEventIngestResult;
}

export interface MeetingTimelineBridge {
  client: MeetingTimelineClient;
  observe(snapshot?: LocalObserverSnapshot, options?: LocalTimelineObserverOptions): Promise<MeetingTimelineBridgeObserveResult>;
  observeCandidates(
    candidates?: LocalObserverSnapshot[] | { snapshots?: LocalObserverSnapshot[]; candidates?: LocalObserverSnapshot[]; items?: LocalObserverSnapshot[]; tabs?: LocalObserverSnapshot[]; windows?: Array<LocalObserverSnapshot & { tabs?: LocalObserverSnapshot[] }> },
    options?: LocalObserverOptions,
  ): Promise<MeetingTimelineBridgeCandidateResult>;
  ingest(
    platformOrInput: string | PlatformEventIngestInput,
    payload?: unknown,
    options?: PlatformEventIngestOptions,
  ): Promise<ReconciledPlatformEventIngestResult>;
  insertMark(input?: TimelineMarkInput, options?: Record<string, unknown>): Promise<unknown>;
  insertAnnotation(input?: TimelineMarkInput, options?: Record<string, unknown>): Promise<unknown>;
  insertMarks(inputs?: TimelineMarkInput[] | { annotations?: TimelineMarkInput[]; items?: TimelineMarkInput[] }, options?: Record<string, unknown>): Promise<unknown>;
  importTranscript(input?: TranscriptImportInput & { raw?: unknown; content?: unknown; platform?: string }, options?: Record<string, unknown>): Promise<unknown>;
  startMeeting(input?: MeetingStartInput): Promise<unknown>;
  endMeeting(input?: MeetingEndInput): Promise<unknown>;
  getState(): Promise<unknown> | undefined;
  getBridgeState(): MeetingTimelineBridgeState;
  reset(nextState?: Partial<MeetingTimelineBridgeState> & {
    localObserver?: LocalObserverState;
    signalReconciler?: Partial<MeetingSignalReconcilerState>;
  }): MeetingTimelineBridgeState;
}

export function createMeetingTimelineBridge(
  clientOrOptions: MeetingTimelineClient | MeetingTimelineClientOptions,
  options?: MeetingTimelineBridgeOptions,
): MeetingTimelineBridge;
