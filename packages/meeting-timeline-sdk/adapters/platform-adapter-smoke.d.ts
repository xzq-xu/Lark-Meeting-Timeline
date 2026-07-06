import type { MeetingPlatformAdapterInstallManifest } from './platform-adapter-install-manifest.mjs';

export const MEETING_PLATFORM_ADAPTER_SMOKE_REPORT_SCHEMA: 'meeting_platform_adapter_smoke_report';
export const MEETING_PLATFORM_ADAPTER_SMOKE_ROW_SCHEMA: 'meeting_platform_adapter_smoke_row';
export const MEETING_PLATFORM_ADAPTER_SMOKE_REPORT_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdapterSmokeOptions {
  platforms?: string[];
  platform_keys?: string[];
  platformKeys?: string[];
  target?: 'static' | 'pilot' | 'production' | string;
  baseUrl?: string;
  base_url?: string;
  captured_at_ms?: number;
  capturedAtMs?: number;
  baseCapturedAtMs?: number;
  base_captured_at_ms?: number;
  fixtureUrls?: Record<string, string>;
  fixture_urls?: Record<string, string>;
  installManifest?: MeetingPlatformAdapterInstallManifest;
  install_manifest?: MeetingPlatformAdapterInstallManifest;
  manifest?: MeetingPlatformAdapterInstallManifest;
  availableFiles?: string[];
  available_files?: string[];
  [key: string]: unknown;
}

export interface MeetingPlatformAdapterSmokeCall {
  method: string;
  platform?: string;
  captured_at_ms?: number;
  label?: string;
  speaker_id?: string;
  participant_id?: string;
  event_type?: string;
  current_url?: string;
  candidate_count?: number;
}

export interface MeetingPlatformAdapterSmokeRow {
  type: 'meeting_platform_adapter_smoke_row';
  schema: 'meeting_platform_adapter_smoke_row';
  schema_version: 1;
  platform: string;
  fixture_url?: string;
  accepted: boolean;
  observed_platform?: string;
  inserted_platform?: string;
  observe_action?: string;
  insert_action?: string;
  speaker_action?: string;
  participant_action?: string;
  provider_action?: string;
  observe_before_insert?: boolean;
  captured_at_ms_expected?: number;
  captured_at_ms_preserved?: boolean;
  speaker_track_inserted?: boolean;
  participant_track_inserted?: boolean;
  provider_reconcile_nonblocking?: boolean;
  bridge_state_opened?: boolean;
  call_count: number;
  calls: MeetingPlatformAdapterSmokeCall[];
  issues: string[];
  error?: Record<string, unknown>;
}

export interface MeetingPlatformAdapterSmokeReport {
  type: 'meeting_platform_adapter_smoke_report';
  schema: 'meeting_platform_adapter_smoke_report';
  schema_version: 1;
  accepted: boolean;
  install_manifest_schema?: string;
  install_manifest_accepted: boolean;
  platform_count: number;
  accepted_count: number;
  failed_count: number;
  platforms: string[];
  runtime_contract: {
    local_axis_first: true;
    timestamp_field: 'captured_at_ms';
    provider_events_block_realtime: false;
    transcript_blocks_realtime: false;
    transcript_required_for_smoke: false;
  };
  verified_sequence: string[];
  rows: MeetingPlatformAdapterSmokeRow[];
  next_actions: string[];
}

export function runMeetingPlatformAdapterSmoke(
  manifestOrOptions?: MeetingPlatformAdapterInstallManifest | MeetingPlatformAdapterSmokeOptions,
  options?: MeetingPlatformAdapterSmokeOptions,
): Promise<MeetingPlatformAdapterSmokeReport>;

export function assertMeetingPlatformAdapterSmoke(
  manifestOrOptions?: MeetingPlatformAdapterInstallManifest | MeetingPlatformAdapterSmokeOptions,
  options?: MeetingPlatformAdapterSmokeOptions,
): Promise<MeetingPlatformAdapterSmokeReport>;
