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
  MeetingPlatformEvidencePackageOptions,
  MeetingPlatformEvidencePackageVerification,
} from './platform-evidence-package.mjs';
import type { MeetingPlatformEvidenceCorrelation } from './platform-evidence-correlation.mjs';
import type { MeetingPlatformAdaptationStrategy } from './platform-strategy.mjs';
import type { MeetingPlatformAdaptationRunbook } from './platform-rollout.mjs';

export const MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA: string;
export const MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION: number;
export const MEETING_PLATFORM_LIVE_ADAPTER_PLAN_SCHEMA: string;
export const MEETING_PLATFORM_LIVE_ADAPTER_MATRIX_SCHEMA: string;
export const MEETING_PLATFORM_LIVE_ADAPTER_READINESS_SCHEMA: string;
export const MEETING_PLATFORM_LIVE_ADAPTER_READINESS_MATRIX_SCHEMA: string;
export const MEETING_PLATFORM_LIVE_ADAPTER_HANDOFF_SCHEMA: string;
export const MEETING_PLATFORM_LIVE_ADAPTER_HANDOFF_BUNDLE_SCHEMA: string;
export const MEETING_PLATFORM_LIVE_ADAPTER_REQUIRED_METHODS: readonly string[];

export interface MeetingPlatformLiveAdapterOptions extends MeetingSourceAggregatorOptions, MeetingPlatformEvidenceSessionOptions, MeetingPlatformEvidencePackageOptions {
  id?: string;
  sessionId?: string;
  session_id?: string;
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  captureEvidence?: boolean;
  capture_evidence?: boolean;
}

export interface MeetingPlatformLiveAdapterReadinessOptions extends MeetingPlatformLiveAdapterOptions {
  target?: 'pilot' | 'production';
  readinessTarget?: 'pilot' | 'production';
  readiness_target?: 'pilot' | 'production';
  adapter?: MeetingPlatformLiveAdapter;
  liveAdapter?: MeetingPlatformLiveAdapter;
  live_adapter?: MeetingPlatformLiveAdapter;
  evidencePackage?: MeetingPlatformEvidencePackage | MeetingPlatformEvidencePackage[] | Record<string, MeetingPlatformEvidencePackage>;
  evidence_package?: MeetingPlatformEvidencePackage | MeetingPlatformEvidencePackage[] | Record<string, MeetingPlatformEvidencePackage>;
  handoffPackage?: MeetingPlatformEvidencePackage | MeetingPlatformEvidencePackage[] | Record<string, MeetingPlatformEvidencePackage>;
  handoff_package?: MeetingPlatformEvidencePackage | MeetingPlatformEvidencePackage[] | Record<string, MeetingPlatformEvidencePackage>;
  requiredMethods?: Iterable<string> | string[];
  required_methods?: Iterable<string> | string[];
  adapterOptions?: MeetingPlatformLiveAdapterOptions;
  adapter_options?: MeetingPlatformLiveAdapterOptions;
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

export interface MeetingPlatformLiveAdapterReadiness {
  type: 'meeting_platform_live_adapter_readiness';
  schema: string;
  schema_version: number;
  platform: string;
  display_name?: string;
  target: 'pilot' | 'production';
  status: 'ready' | 'warning' | 'blocked';
  passed: boolean;
  ready_for_realtime_annotations: boolean;
  production_ready: boolean;
  rollout_status?: string;
  recommended_mode?: string;
  blocking_count: number;
  warning_count: number;
  checks: Record<string, unknown>[];
  blocking_checks: Record<string, unknown>[];
  warnings: Record<string, unknown>[];
  verification?: MeetingPlatformEvidencePackageVerification;
  plan: MeetingPlatformLiveAdapterPlan;
  next_actions: string[];
}

export interface MeetingPlatformLiveAdapterReadinessMatrix {
  type: 'meeting_platform_live_adapter_readiness_matrix';
  schema: string;
  schema_version: number;
  platform_count: number;
  passed_count: number;
  ready_count: number;
  warning_count: number;
  blocked_count: number;
  realtime_ready_count: number;
  production_ready_count: number;
  platforms: string[];
  rows: Record<string, unknown>[];
  reports: MeetingPlatformLiveAdapterReadiness[];
}

export interface MeetingPlatformLiveAdapterHandoff {
  type: 'meeting_platform_live_adapter_handoff';
  schema: string;
  schema_version: number;
  platform: string;
  display_name?: string;
  status: 'ready' | 'warning' | 'blocked';
  passed: boolean;
  target: 'pilot' | 'production';
  rollout_status?: string;
  recommended_mode?: string;
  production_ready: boolean;
  ready_for_realtime_annotations: boolean;
  sdk: Record<string, unknown>;
  host_contract: Record<string, unknown>;
  commands: Record<string, string>;
  evidence_paths: Record<string, string>;
  required_host_inputs: string[];
  realtime_flow: string[];
  outputs: string[];
  plan: MeetingPlatformLiveAdapterPlan;
  readiness: MeetingPlatformLiveAdapterReadiness;
  integration_plan: Record<string, unknown>;
  runbook: MeetingPlatformAdaptationRunbook;
  next_actions: string[];
}

export interface MeetingPlatformLiveAdapterHandoffBundle {
  type: 'meeting_platform_live_adapter_handoff_bundle';
  schema: string;
  schema_version: number;
  platform_count: number;
  passed_count: number;
  ready_count: number;
  blocked_count: number;
  realtime_ready_count: number;
  production_ready_count: number;
  platforms: string[];
  sdk: Record<string, unknown>;
  commands: Record<string, string>;
  evidence_paths: Record<string, string>;
  host_contract: Record<string, unknown>;
  rows: Record<string, unknown>[];
  handoffs: MeetingPlatformLiveAdapterHandoff[];
  live_adapter_matrix: MeetingPlatformLiveAdapterMatrix;
  readiness_matrix: MeetingPlatformLiveAdapterReadinessMatrix;
  next_actions: string[];
}

export interface MeetingPlatformLiveAdapterSuite {
  schema: string;
  schema_version: number;
  platforms: string[];
  adapter(platform: string, adapterOptions?: MeetingPlatformLiveAdapterOptions): MeetingPlatformLiveAdapter;
  liveAdapter(platform: string, adapterOptions?: MeetingPlatformLiveAdapterOptions): MeetingPlatformLiveAdapter;
  adapters(adapterOptions?: MeetingPlatformLiveAdapterOptions): Record<string, MeetingPlatformLiveAdapter>;
  plan(platform: string, planOptions?: MeetingPlatformLiveAdapterOptions): MeetingPlatformLiveAdapterPlan;
  readiness(platform: string, readinessOptions?: MeetingPlatformLiveAdapterReadinessOptions): MeetingPlatformLiveAdapterReadiness;
  assertReadiness(platform: string, readinessOptions?: MeetingPlatformLiveAdapterReadinessOptions): MeetingPlatformLiveAdapterReadiness;
  matrix(matrixOptions?: MeetingPlatformLiveAdapterOptions): MeetingPlatformLiveAdapterMatrix;
  readinessMatrix(readinessOptions?: MeetingPlatformLiveAdapterReadinessOptions): MeetingPlatformLiveAdapterReadinessMatrix;
  assertReadinessMatrix(readinessOptions?: MeetingPlatformLiveAdapterReadinessOptions): MeetingPlatformLiveAdapterReadinessMatrix;
  handoff(platform: string, handoffOptions?: MeetingPlatformLiveAdapterReadinessOptions): MeetingPlatformLiveAdapterHandoff;
  handoffBundle(handoffOptions?: MeetingPlatformLiveAdapterReadinessOptions): MeetingPlatformLiveAdapterHandoffBundle;
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

export function buildMeetingPlatformLiveAdapterReadiness(
  platform: string,
  options?: MeetingPlatformLiveAdapterReadinessOptions,
): MeetingPlatformLiveAdapterReadiness;

export function buildMeetingPlatformLiveAdapterReadinessMatrix(
  options?: MeetingPlatformLiveAdapterReadinessOptions,
): MeetingPlatformLiveAdapterReadinessMatrix;

export function assertMeetingPlatformLiveAdapterReadiness(
  platform: string,
  options?: MeetingPlatformLiveAdapterReadinessOptions,
): MeetingPlatformLiveAdapterReadiness;

export function assertMeetingPlatformLiveAdapterReadinessMatrix(
  options?: MeetingPlatformLiveAdapterReadinessOptions,
): MeetingPlatformLiveAdapterReadinessMatrix;

export function buildMeetingPlatformLiveAdapterHandoff(
  platform: string,
  options?: MeetingPlatformLiveAdapterReadinessOptions,
): MeetingPlatformLiveAdapterHandoff;

export function buildMeetingPlatformLiveAdapterHandoffBundle(
  options?: MeetingPlatformLiveAdapterReadinessOptions,
): MeetingPlatformLiveAdapterHandoffBundle;

export function createMeetingPlatformLiveAdapter(
  platform: string,
  clientOrOptions: MeetingTimelineClient | MeetingPlatformLiveAdapterOptions,
  options?: MeetingPlatformLiveAdapterOptions,
): MeetingPlatformLiveAdapter;

export function createMeetingPlatformLiveAdapterSuite(
  clientOrOptions: MeetingTimelineClient | MeetingPlatformLiveAdapterOptions,
  options?: MeetingPlatformLiveAdapterOptions,
): MeetingPlatformLiveAdapterSuite;
