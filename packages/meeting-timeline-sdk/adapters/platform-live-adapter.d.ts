import type { MeetingTimelineClient } from '../index.mjs';
import type {
  MeetingSourceAggregatorOptions,
  MeetingSourceResult,
} from './meeting-source.mjs';
import type {
  MeetingPlatformEvidenceSession,
  MeetingPlatformEvidenceSessionOptions,
  MeetingPlatformEvidenceSessionSummary,
  MeetingPlatformEvidenceSessionState,
} from './platform-evidence-session.mjs';
import type {
  MeetingPlatformEvidencePackage,
  MeetingPlatformEvidencePackageVerification,
} from './platform-evidence-package.mjs';
import type { MeetingPlatformEvidenceCorrelation } from './platform-evidence-correlation.mjs';
import type { MeetingPlatformAdaptationStrategy } from './platform-strategy.mjs';

export const MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA: string;
export const MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION: number;

export interface MeetingPlatformLiveAdapterOptions extends MeetingSourceAggregatorOptions, MeetingPlatformEvidenceSessionOptions {
  id?: string;
  sessionId?: string;
  session_id?: string;
  captureEvidence?: boolean;
  capture_evidence?: boolean;
}

export interface MeetingPlatformLiveAdapterResult<T = unknown> {
  action: string;
  result: T;
  live_evidence: MeetingPlatformEvidenceSessionSummary;
  evidence_state: MeetingPlatformEvidenceSessionState;
}

export interface MeetingPlatformLiveSourceResult extends MeetingSourceResult {
  action: string;
  live_evidence: MeetingPlatformEvidenceSessionSummary;
  evidence_state: MeetingPlatformEvidenceSessionState;
}

export interface MeetingPlatformLiveAdapter {
  schema: string;
  schema_version: number;
  platform: string;
  source: Record<string, unknown>;
  evidenceSession: MeetingPlatformEvidenceSession;
  session: MeetingPlatformEvidenceSession;
  observeMeetingApp(input?: Record<string, unknown>, options?: MeetingPlatformLiveAdapterOptions): Promise<MeetingPlatformLiveSourceResult>;
  observeApp(input?: Record<string, unknown>, options?: MeetingPlatformLiveAdapterOptions): Promise<MeetingPlatformLiveSourceResult>;
  observeBrowser(input?: Record<string, unknown>, options?: MeetingPlatformLiveAdapterOptions): Promise<MeetingPlatformLiveSourceResult>;
  observeNative(input?: Record<string, unknown>, options?: MeetingPlatformLiveAdapterOptions): Promise<MeetingPlatformLiveSourceResult>;
  observeLocal(input?: Record<string, unknown>, options?: MeetingPlatformLiveAdapterOptions): Promise<MeetingPlatformLiveSourceResult>;
  observeLocalCandidates(candidates?: unknown[] | Record<string, unknown>, options?: MeetingPlatformLiveAdapterOptions): Promise<MeetingPlatformLiveSourceResult>;
  ingestProvider(input?: Record<string, unknown> | string, payload?: unknown, options?: MeetingPlatformLiveAdapterOptions): Promise<MeetingPlatformLiveSourceResult>;
  ingestSignals(signals?: unknown[] | Record<string, unknown>, options?: MeetingPlatformLiveAdapterOptions): Promise<MeetingPlatformLiveSourceResult>;
  captureMeetingAppSnapshot(input?: Record<string, unknown>, options?: MeetingPlatformLiveAdapterOptions): MeetingPlatformLiveAdapterResult<Record<string, unknown>>;
  captureProviderWebhook(input?: Record<string, unknown> | string, payload?: unknown, options?: MeetingPlatformLiveAdapterOptions): MeetingPlatformLiveAdapterResult<Record<string, unknown>>;
  insertAnnotation(input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<MeetingPlatformLiveAdapterResult<unknown>>;
  insertMark(input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<MeetingPlatformLiveAdapterResult<unknown>>;
  insertMarks(inputs?: Record<string, unknown>[], options?: Record<string, unknown>): Promise<MeetingPlatformLiveAdapterResult<unknown>>;
  importTranscript(input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  startMeeting(input?: Record<string, unknown>): Promise<unknown>;
  endMeeting(input?: Record<string, unknown>): Promise<unknown>;
  summary(options?: MeetingPlatformEvidenceSessionOptions): MeetingPlatformEvidenceSessionSummary;
  correlation(options?: Record<string, unknown>): MeetingPlatformEvidenceCorrelation;
  strategy(options?: MeetingPlatformEvidenceSessionOptions): MeetingPlatformAdaptationStrategy;
  exportPackage(options?: MeetingPlatformEvidenceSessionOptions): MeetingPlatformEvidencePackage;
  verify(options?: MeetingPlatformEvidenceSessionOptions): MeetingPlatformEvidencePackageVerification;
  getState(): Record<string, unknown>;
  reset(nextState?: Record<string, unknown>): Record<string, unknown>;
}

export function createMeetingPlatformLiveAdapter(
  platform: string,
  clientOrOptions: MeetingTimelineClient | MeetingPlatformLiveAdapterOptions,
  options?: MeetingPlatformLiveAdapterOptions,
): MeetingPlatformLiveAdapter;
