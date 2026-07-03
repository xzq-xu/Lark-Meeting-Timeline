import type { MeetingTimelineClient } from '../index.mjs';
import type { MeetingPlatformTimelineKit } from './platform-kit.mjs';
import type {
  MeetingPlatformRuntimeBundle,
  MeetingPlatformRuntimeBundleMatrix,
} from './platform-runtime-bundle.mjs';

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
  provider_required_for_realtime: boolean;
  transcript_blocks_realtime: boolean;
  speaker_min_stable_ms?: number;
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
  adaptation_package_matrix: Record<string, unknown>;
  live_adapter_matrix: Record<string, unknown>;
  handoff_readiness_matrix: Record<string, unknown>;
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
  assertManifest(manifestOptions?: MeetingPlatformIntegrationRuntimeOptions): MeetingPlatformIntegrationRuntimeManifest;
  registry(registryOptions?: Record<string, unknown>): {
    manifest: Record<string, unknown>;
    acceptance: Record<string, unknown>;
  };
  runtimeBundle(platform: string, bundleOptions?: Record<string, unknown>): MeetingPlatformRuntimeBundle;
  runtimeBundles(bundleOptions?: Record<string, unknown>): MeetingPlatformRuntimeBundleMatrix;
  adaptationPackages(packageOptions?: Record<string, unknown>): Record<string, unknown>;
  readiness(readinessOptions?: Record<string, unknown>): Record<string, unknown>;
  handoffReadiness(readinessOptions?: Record<string, unknown>): Record<string, unknown>;
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
  getState(): Record<string, unknown>;
  reset(nextState?: Record<string, unknown>): Record<string, unknown>;
}

export function buildMeetingPlatformIntegrationRuntimeManifest(
  options?: MeetingPlatformIntegrationRuntimeOptions,
): MeetingPlatformIntegrationRuntimeManifest;

export function assertMeetingPlatformIntegrationRuntimeManifest(
  manifestOrOptions?: MeetingPlatformIntegrationRuntimeManifest | MeetingPlatformIntegrationRuntimeOptions,
  options?: MeetingPlatformIntegrationRuntimeOptions,
): MeetingPlatformIntegrationRuntimeManifest;

export function createMeetingPlatformIntegrationRuntime(
  clientOrOptions: MeetingTimelineClient | MeetingPlatformIntegrationRuntimeOptions,
  options?: MeetingPlatformIntegrationRuntimeOptions,
): MeetingPlatformIntegrationRuntime;
