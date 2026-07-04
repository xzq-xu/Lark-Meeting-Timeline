export const MEETING_PLATFORM_CONFORMANCE_REPORT_SCHEMA: 'meeting_platform_conformance_report';
export const MEETING_PLATFORM_CONFORMANCE_SCHEMA_VERSION: 1;

export interface MeetingPlatformConformanceOptions {
  baseUrl?: string;
  base_url?: string;
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  env?: Record<string, unknown>;
  contractTarget?: string;
  contract_target?: string;
  [key: string]: unknown;
}

export interface MeetingPlatformConformanceIssue {
  severity: 'error' | 'warning' | string;
  code: string;
  message: string;
  platform?: string;
  [key: string]: unknown;
}

export interface MeetingPlatformConformanceRow {
  platform: string;
  display_name?: string;
  accepted: boolean;
  aliases: string[];
  normalizer_available: boolean;
  setup_manifest_ready: boolean;
  capability_contract_ready: boolean;
  integration_plan_ready: boolean;
  contract_accepted: boolean;
  adapter_route_ready: boolean;
  adapter_first_route?: string;
  runtime_ready: boolean;
  registry_entry_ready: boolean;
  candidate_observation_ready: boolean;
  timestamp_field?: string;
  provider_required_for_realtime: boolean;
  transcript_blocks_realtime: boolean;
  transcript_normalizer?: string;
  transcript_normalizer_available?: boolean;
  issue_count: number;
  blocking_count: number;
  issues: MeetingPlatformConformanceIssue[];
  commands: Record<string, string>;
}

export interface MeetingPlatformConformanceReport {
  type: 'meeting_platform_conformance_report';
  schema: 'meeting_platform_conformance_report';
  schema_version: 1;
  accepted: boolean;
  platform_count: number;
  accepted_count: number;
  normalizer_count: number;
  setup_manifest_ready_count: number;
  capability_contract_ready_count: number;
  integration_plan_ready_count: number;
  adapter_contract_accepted_count: number;
  adapter_route_ready_count: number;
  runtime_ready_count: number;
  registry_entry_ready_count: number;
  candidate_observer_count: number;
  provider_required_for_realtime_count: number;
  transcript_blocking_count: number;
  blocking_count: number;
  platforms: string[];
  rows: MeetingPlatformConformanceRow[];
  registry_acceptance: Record<string, unknown>;
  next_actions: string[];
}

export function buildMeetingPlatformConformanceReport(
  options?: MeetingPlatformConformanceOptions,
): MeetingPlatformConformanceReport;

export function assertMeetingPlatformConformanceReport(
  options?: MeetingPlatformConformanceOptions,
): MeetingPlatformConformanceReport;
