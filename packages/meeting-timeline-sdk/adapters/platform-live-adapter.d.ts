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
export const MEETING_PLATFORM_LIVE_ADAPTER_PLAN_SCHEMA: string;
export const MEETING_PLATFORM_LIVE_ADAPTER_MATRIX_SCHEMA: string;

export interface MeetingPlatformLiveAdapterOptions extends MeetingSourceAggregatorOptions, MeetingPlatformEvidenceSessionOptions {
  id?: string;
  sessionId?: string;
  session_id?: string;
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
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

export interface MeetingPlatformLiveAdapterPlan {
  schema: string;
  schema_version: number;
  type: 'meeting_platform_live_adapter_plan';
  platform: string;
  display_name?: string;
  rollout_status?: string;
  production_ready: boolean;
  ready_for_realtime_annotations: boolean;
  pilot_ready: boolean;
  recommended_mode?: string;
  source_priority?: string[];
  live_adapter: Record<string, unknown>;
  realtime_axis?: Record<string, unknown>;
  local_observer?: Record<string, unknown>;
  provider_events?: Record<string, unknown>;
  speaker_activity?: Record<string, unknown>;
  post_meeting_transcript?: Record<string, unknown>;
  handoff: Record<string, unknown>;
  endpoints?: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingPlatformLiveAdapterMatrix {
  type: 'meeting_platform_live_adapter_matrix';
  schema: string;
  schema_version: number;
  platform_count: number;
  production_ready_count: number;
  pilot_ready_count: number;
  local_first_count: number;
  non_blocking_provider_count: number;
  platforms: string[];
  rows: Record<string, unknown>[];
  plans: MeetingPlatformLiveAdapterPlan[];
}

export interface MeetingPlatformLiveAdapterSuite {
  schema: string;
  schema_version: number;
  platforms: string[];
  adapter(platform: string, adapterOptions?: MeetingPlatformLiveAdapterOptions): MeetingPlatformLiveAdapter;
  liveAdapter(platform: string, adapterOptions?: MeetingPlatformLiveAdapterOptions): MeetingPlatformLiveAdapter;
  adapters(adapterOptions?: MeetingPlatformLiveAdapterOptions): Record<string, MeetingPlatformLiveAdapter>;
  plan(platform: string, planOptions?: MeetingPlatformLiveAdapterOptions): MeetingPlatformLiveAdapterPlan;
  matrix(matrixOptions?: MeetingPlatformLiveAdapterOptions): MeetingPlatformLiveAdapterMatrix;
  summary(summaryOptions?: MeetingPlatformLiveAdapterOptions): Record<string, unknown>;
  getState(): Record<string, unknown>;
  reset(nextState?: Record<string, unknown>): Record<string, unknown>;
}

export function buildMeetingPlatformLiveAdapterPlan(
  platform: string,
  options?: MeetingPlatformLiveAdapterOptions,
): MeetingPlatformLiveAdapterPlan;

export function buildMeetingPlatformLiveAdapterMatrix(
  options?: MeetingPlatformLiveAdapterOptions,
): MeetingPlatformLiveAdapterMatrix;

export function createMeetingPlatformLiveAdapter(
  platform: string,
  clientOrOptions: MeetingTimelineClient | MeetingPlatformLiveAdapterOptions,
  options?: MeetingPlatformLiveAdapterOptions,
): MeetingPlatformLiveAdapter;

export function createMeetingPlatformLiveAdapterSuite(
  clientOrOptions: MeetingTimelineClient | MeetingPlatformLiveAdapterOptions,
  options?: MeetingPlatformLiveAdapterOptions,
): MeetingPlatformLiveAdapterSuite;
