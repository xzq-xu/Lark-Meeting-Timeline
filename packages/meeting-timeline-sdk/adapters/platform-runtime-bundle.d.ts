import type { MeetingPlatformAdaptationPackage, MeetingPlatformAdaptationPackageOptions } from './platform-adaptation-package.mjs';
import type { MeetingPlatformAdapterRoute } from './platform-adapter-route.mjs';
import type { MeetingAppRuntimeObserverPlan } from './meeting-app-profile.mjs';
import type { MeetingAppObserverSchedulerConfig } from './meeting-app-observer-scheduler.mjs';
import type { MeetingPlatformRuntimeEventPlan } from './platform-runtime-event.mjs';

export const MEETING_PLATFORM_RUNTIME_BUNDLE_SCHEMA: 'meeting_platform_runtime_bundle';
export const MEETING_PLATFORM_RUNTIME_BUNDLE_MATRIX_SCHEMA: 'meeting_platform_runtime_bundle_matrix';
export const MEETING_PLATFORM_RUNTIME_BUNDLE_SCHEMA_VERSION: 1;

export interface MeetingPlatformRuntimeBundleOptions extends MeetingPlatformAdaptationPackageOptions {
  bundleId?: string;
  bundle_id?: string;
  contentScriptJs?: string | string[];
  content_script_js?: string | string[];
  js?: string | string[];
  sampleIntervalMs?: number;
  sample_interval_ms?: number;
  mutationDebounceMs?: number;
  mutation_debounce_ms?: number;
  speakerStableFollowupMs?: number;
  speaker_stable_followup_ms?: number;
  observeMutations?: boolean;
  observe_mutations?: boolean;
  url?: string;
  href?: string;
  meeting_url?: string;
  meetingUrl?: string;
  capturedAtMs?: number;
  captured_at_ms?: number;
}

export interface MeetingPlatformRuntimeBundle {
  type: 'meeting_platform_runtime_bundle';
  schema: 'meeting_platform_runtime_bundle';
  schema_version: 1;
  id: string;
  platform: string;
  display_name?: string;
  objective: string;
  runtime_contract: Record<string, unknown>;
  modules: Record<string, string>;
  adapter_route: MeetingPlatformAdapterRoute;
  browser: {
    matches: string[];
    host_permissions: string[];
    content_scripts: Record<string, unknown>[];
    manifest: Record<string, unknown>;
  };
  runtime: {
    preset: string;
    start_options: Record<string, unknown>;
    mutation_observer: Record<string, unknown>;
    capture_profile?: string;
    required_snapshots: unknown[];
    speaker_filter?: Record<string, unknown>;
    content_script_bridge?: {
      module: string;
      create_function: string;
      install_function: string;
      options: Record<string, unknown>;
    };
    lightweight_connector_bridge?: {
      module: string;
      create_hub_function: string;
      create_browser_runtime_function: string;
      create_function: string;
      install_function: string;
      options: Record<string, unknown>;
      route_source_priority: string[];
    };
    observer_plan?: MeetingAppRuntimeObserverPlan | null;
    observation_loop?: Record<string, unknown> | null;
    observer_scheduler?: {
      module: string;
      create_function: string;
      config: MeetingAppObserverSchedulerConfig;
    } | null;
    runtime_host?: {
      module: string;
      create_function: string;
      config_function: string;
      driver: string;
    } | null;
  };
  messaging: {
    message_types: Record<string, string>;
    bridge_message_types?: string[];
    lightweight_connector_message_types?: string[];
    accepted_methods: string[];
    runtime_event?: {
      schema: string;
      endpoint?: string;
      build_function: string;
      client_factory: string;
      plan_schema?: string;
      plan?: MeetingPlatformRuntimeEventPlan;
    };
    examples: Record<string, unknown>;
  };
  host: {
    endpoints: Record<string, string>;
    annotation_timestamp_field: string;
    insert_policy?: Record<string, unknown>;
  };
  provider_reconcile?: Record<string, unknown>;
  transcript: Record<string, unknown>;
  readiness: {
    sdk_wiring_ready: boolean;
    runtime_ready: boolean;
    observer_plan_ready?: boolean;
    observer_scheduler_ready?: boolean;
    runtime_host_ready?: boolean;
    lightweight_connector_ready?: boolean;
    observer_preflight_status?: string;
    provider_required_for_realtime: boolean;
    transcript_blocks_realtime: boolean;
    missing_items: string[];
  };
  observer_plan?: MeetingAppRuntimeObserverPlan | null;
  adaptation_package: MeetingPlatformAdaptationPackage;
  next_actions: string[];
}

export interface MeetingPlatformRuntimeBundleMatrix {
  type: 'meeting_platform_runtime_bundle_matrix';
  schema: 'meeting_platform_runtime_bundle_matrix';
  schema_version: 1;
  platform_count: number;
  runtime_ready_count: number;
  sdk_wiring_ready_count: number;
  observer_plan_ready_count?: number;
  runtime_host_ready_count?: number;
  lightweight_connector_ready_count?: number;
  adapter_route_ready_count?: number;
  local_observer_first_count?: number;
  provider_non_blocking_route_count?: number;
  transcript_non_blocking_route_count?: number;
  candidate_observer_count: number;
  provider_required_for_realtime_count: number;
  transcript_blocking_count: number;
  platforms: string[];
  rows: Record<string, unknown>[];
  bundles: MeetingPlatformRuntimeBundle[];
  next_actions: string[];
}

export function buildMeetingPlatformRuntimeBundle(
  platform: string,
  options?: MeetingPlatformRuntimeBundleOptions,
): MeetingPlatformRuntimeBundle;

export function buildMeetingPlatformRuntimeBundleMatrix(
  options?: MeetingPlatformRuntimeBundleOptions,
): MeetingPlatformRuntimeBundleMatrix;
