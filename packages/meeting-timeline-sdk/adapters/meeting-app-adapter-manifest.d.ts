import type { MeetingAppDomCaptureProfilePlatform } from './meeting-app-capture.mjs';

export const MEETING_APP_ADAPTER_MANIFEST_SCHEMA: 'meeting_app_adapter_manifest';
export const MEETING_APP_ADAPTER_MANIFEST_MATRIX_SCHEMA: 'meeting_app_adapter_manifest_matrix';
export const MEETING_APP_ADAPTER_MANIFEST_SCHEMA_VERSION: 1;

export interface MeetingAppAdapterManifestOptions {
  baseUrl?: string;
  base_url?: string;
  platform?: string;
  provider?: string;
  key?: string;
  name?: string;
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  [key: string]: unknown;
}

export interface MeetingAppAdapterManifestIssue {
  severity: 'error' | 'warning' | 'info' | string;
  code: string;
  message: string;
  platform?: string;
  [key: string]: unknown;
}

export interface MeetingAppAdapterManifest {
  type: 'meeting_app_adapter_manifest';
  schema: typeof MEETING_APP_ADAPTER_MANIFEST_SCHEMA;
  schema_version: typeof MEETING_APP_ADAPTER_MANIFEST_SCHEMA_VERSION;
  accepted: boolean;
  platform: MeetingAppDomCaptureProfilePlatform;
  display_name: string;
  recommended_mode: string;
  surface: string;
  extension: Record<string, unknown>;
  runtime: Record<string, unknown>;
  capture: Record<string, unknown>;
  bridge: Record<string, unknown>;
  observer_plan: Record<string, unknown>;
  contracts: Record<string, unknown>;
  host_endpoints: Record<string, string>;
  implementation_order: string[];
  readiness: Record<string, unknown>;
  commands: Record<string, string>;
  acceptance_report: Record<string, unknown>;
  handoff: Record<string, unknown>;
  issue_count: number;
  blocking_count: number;
  warning_count: number;
  issues: MeetingAppAdapterManifestIssue[];
  next_actions: string[];
}

export interface MeetingAppAdapterManifestMatrix {
  type: 'meeting_app_adapter_manifest_matrix';
  schema: typeof MEETING_APP_ADAPTER_MANIFEST_MATRIX_SCHEMA;
  schema_version: typeof MEETING_APP_ADAPTER_MANIFEST_SCHEMA_VERSION;
  accepted: boolean;
  platform_count: number;
  accepted_count: number;
  extension_match_count: number;
  mutation_observer_count: number;
  candidate_observer_count: number;
  participant_selector_ready_count: number;
  platforms: MeetingAppDomCaptureProfilePlatform[];
  rows: Array<Record<string, unknown>>;
  manifests: MeetingAppAdapterManifest[];
  next_actions: string[];
}

export function buildMeetingAppAdapterManifest(
  platformOrOptions?: string | MeetingAppAdapterManifestOptions,
  options?: MeetingAppAdapterManifestOptions,
): MeetingAppAdapterManifest;

export function buildMeetingAppAdapterManifestMatrix(
  options?: MeetingAppAdapterManifestOptions,
): MeetingAppAdapterManifestMatrix;

export function assertMeetingAppAdapterManifest(
  manifestOrPlatform?: MeetingAppAdapterManifest | string | MeetingAppAdapterManifestOptions,
  options?: MeetingAppAdapterManifestOptions,
): MeetingAppAdapterManifest;

export function assertMeetingAppAdapterManifestMatrix(
  matrixOrOptions?: MeetingAppAdapterManifestMatrix | MeetingAppAdapterManifestOptions,
  options?: MeetingAppAdapterManifestOptions,
): MeetingAppAdapterManifestMatrix;
