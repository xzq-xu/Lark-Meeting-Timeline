import type { NormalizedMeetingSignal } from './core.mjs';

export interface MeetingPlatformEventAdapter {
  key: string;
  aliases: readonly string[];
  source: string;
  normalize(raw?: unknown, options?: Record<string, unknown>): NormalizedMeetingSignal[];
}

export const MEETING_PLATFORM_REGISTRY_ENTRY_SCHEMA: 'meeting_platform_registry_entry';
export const MEETING_PLATFORM_REGISTRY_MANIFEST_SCHEMA: 'meeting_platform_registry_manifest';
export const MEETING_PLATFORM_REGISTRY_ACCEPTANCE_SCHEMA: 'meeting_platform_registry_acceptance';
export const MEETING_PLATFORM_REGISTRY_SCHEMA_VERSION: 1;
export const MEETING_PLATFORM_EVENT_ADAPTERS: readonly MeetingPlatformEventAdapter[];

export function meetingPlatformEventAdapterFor(platform: string): MeetingPlatformEventAdapter | null;

export interface MeetingPlatformRegistryOptions {
  baseUrl?: string;
  base_url?: string;
  basePath?: string;
  base_path?: string;
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  env?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingPlatformRegistryEntry {
  type: 'meeting_platform_registry_entry';
  schema: 'meeting_platform_registry_entry';
  schema_version: 1;
  platform: string;
  display_name?: string;
  aliases: string[];
  event_adapter: Record<string, unknown>;
  supported_surfaces?: Record<string, unknown>;
  runtime: Record<string, unknown>;
  provider: Record<string, unknown>;
  annotations: Record<string, unknown>;
  adapter_route: Record<string, unknown>;
  transcript: Record<string, unknown>;
  host: Record<string, unknown>;
  sdk: Record<string, unknown>;
  readiness: Record<string, unknown>;
  commands: Record<string, string>;
  next_actions: string[];
}

export interface MeetingPlatformRegistryManifest {
  type: 'meeting_platform_registry_manifest';
  schema: 'meeting_platform_registry_manifest';
  schema_version: 1;
  platform_count: number;
  normalizer_count: number;
  runtime_ready_count: number;
  contract_accepted_count: number;
  candidate_observer_count: number;
  provider_required_for_realtime_count: number;
  transcript_blocking_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  entries: MeetingPlatformRegistryEntry[];
  next_actions: string[];
}

export interface MeetingPlatformRegistryAcceptanceReport {
  type: 'meeting_platform_registry_acceptance';
  schema: 'meeting_platform_registry_acceptance';
  schema_version: 1;
  accepted: boolean;
  platform_count: number;
  normalizer_count: number;
  runtime_ready_count: number;
  contract_accepted_count: number;
  candidate_observer_count: number;
  provider_required_for_realtime_count: number;
  transcript_blocking_count: number;
  blocking_count: number;
  warning_count: number;
  issues: Array<Record<string, unknown>>;
  manifest: MeetingPlatformRegistryManifest;
  next_actions: string[];
}

export function buildMeetingPlatformRegistryEntry(
  platform: string,
  options?: MeetingPlatformRegistryOptions,
): MeetingPlatformRegistryEntry;

export function buildMeetingPlatformRegistryManifest(
  options?: MeetingPlatformRegistryOptions,
): MeetingPlatformRegistryManifest;

export function buildMeetingPlatformRegistryAcceptanceReport(
  manifestOrOptions?: MeetingPlatformRegistryManifest | MeetingPlatformRegistryEntry | MeetingPlatformRegistryOptions,
  options?: MeetingPlatformRegistryOptions,
): MeetingPlatformRegistryAcceptanceReport;

export function assertMeetingPlatformRegistryManifest(
  manifestOrOptions?: MeetingPlatformRegistryManifest | MeetingPlatformRegistryEntry | MeetingPlatformRegistryOptions,
  options?: MeetingPlatformRegistryOptions,
): MeetingPlatformRegistryAcceptanceReport;
