import type { MeetingTimelineClient } from '../index.mjs';
import type { MeetingPlatformTimelineKit } from './platform-kit.mjs';
import type {
  MeetingPlatformRuntimeBundle,
  MeetingPlatformRuntimeBundleMatrix,
} from './platform-runtime-bundle.mjs';
import type { MeetingPlatformAdaptationStrategyMatrix } from './platform-strategy.mjs';
import type { MeetingPlatformParticipantTrackMatrix } from './platform-participant-track.mjs';
import type { MeetingPlatformSpeakerTrackMatrix } from './platform-speaker-track.mjs';
import type {
  MeetingSessionEnvironmentSnapshot,
  MeetingSessionDiscoverySnapshot,
  MeetingSessionSelection,
  NormalizedMeetingSessionCandidate,
} from './meeting-session-discovery.mjs';

export const MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA: 'meeting_platform_integration_runtime';
export const MEETING_PLATFORM_INTEGRATION_RUNTIME_MANIFEST_SCHEMA: 'meeting_platform_integration_runtime_manifest';
export const MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA_VERSION: 1;

export interface MeetingPlatformIntegrationRuntimeOptions {
  baseUrl?: string;
  base_url?: string;
  basePath?: string;
  base_path?: string;
  platforms?: string[];
  platform_keys?: string[];
  env?: Record<string, unknown>;
  clientOptions?: Record<string, unknown>;
  client_options?: Record<string, unknown>;
  requireHandoffReady?: boolean;
  require_handoff_ready?: boolean;
  handoffReadinessMatrix?: Record<string, unknown>;
  handoff_readiness_matrix?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingPlatformIntegrationRuntimeIssue {
  severity: 'error' | 'warning' | 'info';
  code: string;
  message: string;
  [key: string]: unknown;
}

export interface MeetingPlatformIntegrationRuntimeRow {
  platform: string;
  display_name?: string;
  runtime_ready: boolean;
  sdk_wiring_ready: boolean;
  browser_match_count?: number;
  recommended_mode?: string;
  primary_axis_source?: string;
  strategy_recommendation?: string;
  provider_required_for_realtime: boolean;
  provider_blocks_realtime: boolean;
  transcript_blocks_realtime: boolean;
  runtime_host_replay_required?: boolean;
  runtime_host_replay_accepted?: boolean;
  runtime_host_replay_missing?: string[];
  speaker_track_ready?: boolean;
  speaker_min_stable_ms?: number;
  speaker_switch_stable_ms?: number;
  speaker_end_idle_ms?: number;
  speaker_provider_blocks_realtime?: boolean;
  speaker_transcript_blocks_realtime?: boolean;
  participant_track_ready?: boolean;
  participant_duplicate_window_ms?: number;
  participant_leave_stable_ms?: number;
  participant_provider_blocks_realtime?: boolean;
  participant_transcript_blocks_realtime?: boolean;
  handoff_ready: boolean;
  pilot_ready: boolean;
  production_ready: boolean;
  first_next_action?: string;
}

export interface MeetingPlatformIntegrationRuntimeManifest {
  type: 'meeting_platform_integration_runtime_manifest';
  schema: typeof MEETING_PLATFORM_INTEGRATION_RUNTIME_MANIFEST_SCHEMA;
  schema_version: typeof MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA_VERSION;
  base_url: string;
  base_path: string;
  require_handoff_ready?: boolean;
  platform_count: number;
  platforms: string[];
  host_integration_ready: boolean;
  blocking_count: number;
  warning_count: number;
  issues: MeetingPlatformIntegrationRuntimeIssue[];
  next_actions: string[];
  rows: MeetingPlatformIntegrationRuntimeRow[];
  registry_acceptance: Record<string, unknown>;
  runtime_bundle_matrix: MeetingPlatformRuntimeBundleMatrix;
  adaptation_strategy_matrix: MeetingPlatformAdaptationStrategyMatrix;
  speaker_track_matrix: MeetingPlatformSpeakerTrackMatrix;
  participant_track_matrix: MeetingPlatformParticipantTrackMatrix;
  adaptation_package_matrix: Record<string, unknown>;
  live_adapter_matrix: Record<string, unknown>;
  handoff_readiness_matrix: Record<string, unknown>;
}

export interface MeetingPlatformBrowserDetection {
  type: 'meeting_platform_browser_detection';
  detected: boolean;
  platform?: string;
  reason: 'explicit' | 'url' | 'runtime_preset' | 'none';
  meeting?: Record<string, unknown>;
  browser?: {
    url?: string;
    title?: string;
  };
  runtime_preset_platform?: string;
}

export interface MeetingPlatformResolution {
  type: 'meeting_platform_resolution';
  detected: boolean;
  supported: boolean;
  platform?: string;
  reason?: MeetingPlatformBrowserDetection['reason'];
  meeting?: Record<string, unknown>;
  browser?: MeetingPlatformBrowserDetection['browser'];
  current_platforms: string[];
  display_name?: string;
  registry?: {
    aliases?: string[];
    normalize_available?: boolean;
    insert_endpoint?: string;
    runtime_event_endpoint?: string;
    [key: string]: unknown;
  };
  strategy?: {
    rollout_status?: string;
    recommendation?: string;
    primary_axis_source?: string;
    provider_blocks_realtime?: boolean;
    transcript_blocks_realtime?: boolean;
    speaker_realtime_primary?: boolean;
    [key: string]: unknown;
  };
  runtime?: {
    runtime_ready?: boolean;
    browser_match_count?: number;
    provider_required_for_realtime?: boolean;
    transcript_blocks_realtime?: boolean;
    speaker_min_stable_ms?: number;
    [key: string]: unknown;
  };
  next_actions: string[];
  detection: MeetingPlatformBrowserDetection;
}

export interface MeetingPlatformCandidateResolutionRow {
  selected?: boolean;
  rank?: number;
  score?: number;
  supported: boolean;
  platform?: string;
  meeting_id?: string;
  title?: string;
  meeting_url?: string;
  discovery?: Record<string, unknown>;
  candidate: NormalizedMeetingSessionCandidate;
  resolution: MeetingPlatformResolution;
}

export interface MeetingPlatformCandidateResolution {
  type: 'meeting_platform_candidate_resolution';
  detected: boolean;
  supported: boolean;
  platform?: string;
  meeting?: Record<string, unknown>;
  selected_candidate?: MeetingPlatformCandidateResolutionRow;
  selected_resolution?: MeetingPlatformResolution;
  candidate_count: number;
  supported_candidate_count: number;
  current_platforms: string[];
  next_actions: string[];
  candidates: MeetingPlatformCandidateResolutionRow[];
  selection: MeetingSessionSelection;
}

export interface MeetingPlatformCandidateObservation {
  type: 'meeting_platform_candidate_observation';
  action: 'observe_platform_candidates';
  source: 'platform_candidates';
  platform?: string;
  supported: boolean;
  detected: boolean;
  selected_candidate?: MeetingPlatformCandidateResolutionRow;
  selected_resolution?: MeetingPlatformResolution;
  platform_candidate_resolution: MeetingPlatformCandidateResolution;
  signals: Array<Record<string, unknown>>;
  rawSignals: Array<Record<string, unknown>>;
  raw_signals: Array<Record<string, unknown>>;
  results: Array<Record<string, unknown>>;
  platform_results: Array<{
    platform: string;
    candidate_count: number;
    selected?: boolean;
    result: Record<string, unknown>;
  }>;
  selected_result?: Record<string, unknown>;
  diagnostic: Record<string, unknown>;
}

export interface MeetingPlatformIntegrationRuntime {
  type: 'meeting_platform_integration_runtime';
  schema: typeof MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA;
  schema_version: typeof MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA_VERSION;
  platforms: string[];
  client: MeetingTimelineClient;
  kit: MeetingPlatformTimelineKit;
  liveAdapters: Record<string, unknown>;
  live_adapters: Record<string, unknown>;
  adapter(platform: string, adapterOptions?: Record<string, unknown>): Record<string, unknown>;
  platformAdapter(platform: string, adapterOptions?: Record<string, unknown>): Record<string, unknown>;
  manifest(manifestOptions?: MeetingPlatformIntegrationRuntimeOptions): MeetingPlatformIntegrationRuntimeManifest;
  runManifest(manifestOptions?: MeetingPlatformIntegrationRuntimeOptions): Promise<MeetingPlatformIntegrationRuntimeManifest>;
  assertManifest(manifestOptions?: MeetingPlatformIntegrationRuntimeOptions): MeetingPlatformIntegrationRuntimeManifest;
  runAndAssertManifest(manifestOptions?: MeetingPlatformIntegrationRuntimeOptions): Promise<MeetingPlatformIntegrationRuntimeManifest>;
  registry(registryOptions?: Record<string, unknown>): {
    manifest: Record<string, unknown>;
    acceptance: Record<string, unknown>;
  };
  runtimeBundle(platform: string, bundleOptions?: Record<string, unknown>): MeetingPlatformRuntimeBundle;
  runtimeBundles(bundleOptions?: Record<string, unknown>): MeetingPlatformRuntimeBundleMatrix;
  adaptationPackages(packageOptions?: Record<string, unknown>): Record<string, unknown>;
  adaptationStrategyMatrix(strategyOptions?: Record<string, unknown>): MeetingPlatformAdaptationStrategyMatrix;
  resolvePlatform(input?: Record<string, unknown>, resolveOptions?: Record<string, unknown>): MeetingPlatformResolution;
  resolvePlatformCandidates(
    input?: MeetingSessionEnvironmentSnapshot | MeetingSessionDiscoverySnapshot[] | Record<string, unknown>,
    resolveOptions?: Record<string, unknown>,
  ): MeetingPlatformCandidateResolution;
  observePlatformCandidates(
    input?: MeetingSessionEnvironmentSnapshot | MeetingSessionDiscoverySnapshot[] | Record<string, unknown>,
    observeOptions?: Record<string, unknown>,
  ): Promise<MeetingPlatformCandidateObservation>;
  readiness(readinessOptions?: Record<string, unknown>): Record<string, unknown>;
  handoffReadiness(readinessOptions?: Record<string, unknown>): Record<string, unknown>;
  runHandoffReadiness(readinessOptions?: Record<string, unknown>): Promise<Record<string, unknown>>;
  observeMeetingApp(platform: string, snapshot?: Record<string, unknown>, observeOptions?: Record<string, unknown>): Promise<Record<string, unknown>>;
  observeApp(platform: string, snapshot?: Record<string, unknown>, observeOptions?: Record<string, unknown>): Promise<Record<string, unknown>>;
  ingestProvider(platform: string, requestOrPayload?: Record<string, unknown>, payload?: unknown, ingestOptions?: Record<string, unknown>): Promise<Record<string, unknown>>;
  insertAnnotation(platform: string, input?: Record<string, unknown>, markOptions?: Record<string, unknown>): Promise<Record<string, unknown>>;
  insertMark(platform: string, input?: Record<string, unknown>, markOptions?: Record<string, unknown>): Promise<Record<string, unknown>>;
  speakerTrack(platform: string, input?: Record<string, unknown>, trackOptions?: Record<string, unknown>): Record<string, unknown>;
  participantTrack(platform: string, input?: Record<string, unknown>, trackOptions?: Record<string, unknown>): Record<string, unknown>;
  timelineView(platform: string, input?: Record<string, unknown>, viewOptions?: Record<string, unknown>): Record<string, unknown>;
  handleEvent(input?: Record<string, unknown>, payload?: unknown, eventOptions?: Record<string, unknown>): Promise<Record<string, unknown>>;
  summary(summaryOptions?: MeetingPlatformIntegrationRuntimeOptions): Record<string, unknown>;
  runSummary(summaryOptions?: MeetingPlatformIntegrationRuntimeOptions): Promise<Record<string, unknown>>;
  getState(): Record<string, unknown>;
  reset(nextState?: Record<string, unknown>): Record<string, unknown>;
}

export interface MeetingPlatformIntegrationBrowserRuntime {
  type: 'meeting_platform_integration_browser_runtime';
  schema: typeof MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA;
  schema_version: typeof MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA_VERSION;
  integrationRuntime: MeetingPlatformIntegrationRuntime;
  integration_runtime: MeetingPlatformIntegrationRuntime;
  detect(input?: Record<string, unknown>, detectOptions?: Record<string, unknown>): MeetingPlatformBrowserDetection;
  resolvePlatform(input?: Record<string, unknown>, resolveOptions?: Record<string, unknown>): MeetingPlatformResolution;
  resolvePlatformCandidates(
    input?: MeetingSessionEnvironmentSnapshot | MeetingSessionDiscoverySnapshot[] | Record<string, unknown>,
    resolveOptions?: Record<string, unknown>,
  ): MeetingPlatformCandidateResolution;
  observePlatformCandidates(
    input?: MeetingSessionEnvironmentSnapshot | MeetingSessionDiscoverySnapshot[] | Record<string, unknown>,
    observeOptions?: Record<string, unknown>,
  ): Promise<MeetingPlatformCandidateObservation>;
  platformFor(input?: Record<string, unknown>, platformOptions?: Record<string, unknown>): string;
  sample(sampleOptions?: Record<string, unknown>): Promise<Record<string, unknown>>;
  tick(sampleOptions?: Record<string, unknown>): Promise<Record<string, unknown>>;
  start(startOptions?: Record<string, unknown>): Record<string, unknown>;
  stop(): Record<string, unknown>;
  dispose(): Record<string, unknown>;
  handleMessage(message?: Record<string, unknown>, messageOptions?: Record<string, unknown>): Promise<Record<string, unknown>>;
  observeMeetingApp(input?: Record<string, unknown>, observeOptions?: Record<string, unknown>): Promise<Record<string, unknown>>;
  insertAnnotation(input?: Record<string, unknown>, markOptions?: Record<string, unknown>): Promise<Record<string, unknown>>;
  getState(): Record<string, unknown>;
  reset(nextState?: Record<string, unknown>): Record<string, unknown>;
}

export interface MeetingPlatformIntegrationContentScriptBridge {
  type: 'meeting_platform_integration_content_script_bridge';
  schema: typeof MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA;
  schema_version: typeof MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA_VERSION;
  runtime: MeetingPlatformIntegrationBrowserRuntime | Record<string, unknown>;
  integrationRuntime?: MeetingPlatformIntegrationRuntime;
  integration_runtime?: MeetingPlatformIntegrationRuntime;
  detect(
    input?: Record<string, unknown>,
    detectOptions?: Record<string, unknown>,
  ): MeetingPlatformBrowserDetection | undefined;
  dispatchMessage(
    message?: Record<string, unknown>,
    messageOptions?: Record<string, unknown>,
  ): Promise<Record<string, unknown>>;
  start(startOptions?: Record<string, unknown>): Record<string, unknown>;
  stop(): Record<string, unknown>;
  dispose(): Record<string, unknown>;
  installExtensionMessaging(installOptions?: Record<string, unknown>): Record<string, unknown>;
  installWindowMessaging(installOptions?: Record<string, unknown>): Record<string, unknown>;
  removeMessaging(kind?: string): Record<string, unknown>;
  getState(): Record<string, unknown>;
}

export function detectMeetingPlatformForBrowser(
  input?: Record<string, unknown>,
  options?: MeetingPlatformIntegrationRuntimeOptions,
): MeetingPlatformBrowserDetection;

export function resolveMeetingPlatformForInput(
  input?: Record<string, unknown>,
  options?: MeetingPlatformIntegrationRuntimeOptions,
): MeetingPlatformResolution;

export function resolveMeetingPlatformCandidates(
  input?: MeetingSessionEnvironmentSnapshot | MeetingSessionDiscoverySnapshot[] | Record<string, unknown>,
  options?: MeetingPlatformIntegrationRuntimeOptions,
): MeetingPlatformCandidateResolution;

export function buildMeetingPlatformIntegrationRuntimeManifest(
  options?: MeetingPlatformIntegrationRuntimeOptions,
): MeetingPlatformIntegrationRuntimeManifest;

export function runMeetingPlatformIntegrationRuntimeManifest(
  options?: MeetingPlatformIntegrationRuntimeOptions,
): Promise<MeetingPlatformIntegrationRuntimeManifest>;

export function assertMeetingPlatformIntegrationRuntimeManifest(
  manifestOrOptions?: MeetingPlatformIntegrationRuntimeManifest | MeetingPlatformIntegrationRuntimeOptions,
  options?: MeetingPlatformIntegrationRuntimeOptions,
): MeetingPlatformIntegrationRuntimeManifest;

export function createMeetingPlatformIntegrationRuntime(
  clientOrOptions: MeetingTimelineClient | MeetingPlatformIntegrationRuntimeOptions,
  options?: MeetingPlatformIntegrationRuntimeOptions,
): MeetingPlatformIntegrationRuntime;

export function createMeetingPlatformIntegrationBrowserRuntime(
  clientOrOptions: MeetingTimelineClient | MeetingPlatformIntegrationRuntimeOptions,
  options?: MeetingPlatformIntegrationRuntimeOptions,
): MeetingPlatformIntegrationBrowserRuntime;

export function createMeetingPlatformIntegrationContentScriptBridge(
  clientOrOptions: MeetingTimelineClient | MeetingPlatformIntegrationRuntimeOptions,
  options?: MeetingPlatformIntegrationRuntimeOptions,
): MeetingPlatformIntegrationContentScriptBridge;

export function installMeetingPlatformIntegrationContentScriptBridge(
  clientOrOptions: MeetingTimelineClient | MeetingPlatformIntegrationRuntimeOptions,
  options?: MeetingPlatformIntegrationRuntimeOptions,
): MeetingPlatformIntegrationContentScriptBridge;
