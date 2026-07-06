import type { MeetingPlatformConformanceReport } from './platform-conformance.mjs';

export const MEETING_PLATFORM_CONSUMER_HANDOFF_SCHEMA: 'meeting_platform_consumer_handoff';
export const MEETING_PLATFORM_CONSUMER_HANDOFF_SCHEMA_VERSION: 1;

export interface MeetingPlatformConsumerHandoffOptions {
  baseUrl?: string;
  base_url?: string;
  basePath?: string;
  base_path?: string;
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  env?: Record<string, unknown>;
  requireHandoffReady?: boolean;
  require_handoff_ready?: boolean;
  requireProductionReady?: boolean;
  require_production_ready?: boolean;
  includeDetails?: boolean;
  include_details?: boolean;
  handoffReadinessMatrix?: Record<string, unknown>;
  handoff_readiness_matrix?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingPlatformConsumerHandoffIssue {
  severity: 'error' | 'warning' | 'info' | string;
  code: string;
  message: string;
  platform?: string;
  [key: string]: unknown;
}

export interface MeetingPlatformConsumerHandoffRow {
  platform: string;
  display_name?: string;
  consumer_ready: boolean;
  conformance_accepted: boolean;
  conformance_blocking_count?: number;
  runtime_ready: boolean;
  sdk_wiring_ready: boolean;
  candidate_observation_ready: boolean;
  candidate_observer_message_type?: string;
  adapter_route_ready: boolean;
  adapter_first_route?: string;
  adapter_recommended_mode?: string;
  speaker_track_ready: boolean;
  speaker_min_stable_ms?: number;
  speaker_switch_stable_ms?: number;
  speaker_end_idle_ms?: number;
  participant_track_ready: boolean;
  participant_duplicate_window_ms?: number;
  participant_leave_stable_ms?: number;
  provider_blocks_realtime: boolean;
  transcript_blocks_realtime: boolean;
  handoff_ready: boolean;
  pilot_ready: boolean;
  production_ready: boolean;
  first_next_action?: string;
  issue_count: number;
  blocking_count: number;
  issues: MeetingPlatformConsumerHandoffIssue[];
}

export interface MeetingPlatformConsumerHandoffEntrypoints {
  package: string;
  primary_modules: Record<string, string | undefined>;
  kit_methods: string[];
  host_methods: string[];
  http_endpoints: Record<string, string>;
  commands: Record<string, string>;
}

export interface MeetingPlatformConsumerHandoff {
  type: 'meeting_platform_consumer_handoff';
  schema: typeof MEETING_PLATFORM_CONSUMER_HANDOFF_SCHEMA;
  schema_version: typeof MEETING_PLATFORM_CONSUMER_HANDOFF_SCHEMA_VERSION;
  accepted: boolean;
  base_url: string;
  base_path: string;
  platform_count: number;
  consumer_ready_count: number;
  conformance_accepted_count: number;
  runtime_ready_count: number;
  adapter_route_ready_count: number;
  candidate_observer_count: number;
  speaker_track_ready_count: number;
  participant_track_ready_count: number;
  handoff_ready_count: number;
  pilot_ready_count: number;
  production_ready_count: number;
  blocking_count: number;
  warning_count: number;
  require_handoff_ready: boolean;
  require_production_ready: boolean;
  platforms: string[];
  entrypoints: MeetingPlatformConsumerHandoffEntrypoints;
  hard_contracts: Record<string, unknown>;
  boot_order: Array<Record<string, unknown>>;
  endpoints: Record<string, string>;
  commands: Record<string, string>;
  rows: MeetingPlatformConsumerHandoffRow[];
  issues: MeetingPlatformConsumerHandoffIssue[];
  next_actions: string[];
  host_integration_plan: Record<string, unknown>;
  conformance_report: MeetingPlatformConformanceReport;
  adaptation_package_matrix: Record<string, unknown>;
  handoff_readiness_matrix: Record<string, unknown>;
}

export function buildMeetingPlatformConsumerHandoff(
  options?: MeetingPlatformConsumerHandoffOptions,
): MeetingPlatformConsumerHandoff;

export function assertMeetingPlatformConsumerHandoff(
  handoffOrOptions?: MeetingPlatformConsumerHandoff | MeetingPlatformConsumerHandoffOptions,
  options?: MeetingPlatformConsumerHandoffOptions,
): MeetingPlatformConsumerHandoff;
