import type { MeetingPlatformAdapterAcceptanceChecklistOptions } from './platform-adapter-acceptance-checklist.mjs';
import type { MeetingPlatformAdapterAuthoringOptions } from './platform-adapter-authoring.mjs';
import type { MeetingPlatformAdapterPortfolioOptions } from './platform-adapter-portfolio.mjs';
import type { MeetingPlatformImplementationHandoffOptions } from './platform-implementation-handoff.mjs';
import type { MeetingPlatformRuntimeBundleOptions } from './platform-runtime-bundle.mjs';
import type { MeetingPlatformProviderConnectionOptions } from './platform-provider-connection.mjs';
import type { MeetingPlatformAdapterContractOptions } from './platform-adapter-contract.mjs';
import type { MeetingPlatformAdaptationPackageOptions } from './platform-adaptation-package.mjs';

export const MEETING_PLATFORM_ADAPTER_EXPORT_PACKAGE_SCHEMA: 'meeting_platform_adapter_export_package';
export const MEETING_PLATFORM_ADAPTER_EXPORT_PACKAGE_MATRIX_SCHEMA: 'meeting_platform_adapter_export_package_matrix';
export const MEETING_PLATFORM_ADAPTER_EXPORT_PACKAGE_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdapterExportPackageOptions
  extends MeetingPlatformAdapterAcceptanceChecklistOptions,
    MeetingPlatformAdapterAuthoringOptions,
    MeetingPlatformAdapterPortfolioOptions,
    MeetingPlatformImplementationHandoffOptions,
    MeetingPlatformRuntimeBundleOptions,
    MeetingPlatformProviderConnectionOptions,
    MeetingPlatformAdapterContractOptions,
    MeetingPlatformAdaptationPackageOptions {
  platforms?: string[];
  platform_keys?: string[];
  target?: 'static' | 'pilot' | 'production' | string;
  acceptanceTarget?: 'static' | 'pilot' | 'production' | string;
  acceptance_target?: 'static' | 'pilot' | 'production' | string;
  includeArtifacts?: boolean;
  include_artifacts?: boolean;
}

export interface MeetingPlatformAdapterExportHostFile {
  path: string;
  source: string;
  required_from: string;
  purpose: string;
}

export interface MeetingPlatformAdapterExportPackage {
  type: 'meeting_platform_adapter_export_package';
  schema: 'meeting_platform_adapter_export_package';
  schema_version: 1;
  platform: string;
  display_name: string;
  built_in: boolean;
  target: string;
  accepted: boolean;
  export_ready: boolean;
  package_role: string;
  recommended_first_surface?: string;
  local_axis_first: boolean;
  timestamp_field: 'captured_at_ms';
  provider_events_block_realtime: false;
  transcript_blocks_realtime: false;
  import_paths: Record<string, string>;
  artifact_refs: Record<string, Record<string, unknown>>;
  host_files: MeetingPlatformAdapterExportHostFile[];
  surface_entrypoints: Record<string, Record<string, unknown>>;
  setup_order: Array<Record<string, unknown>>;
  commands: Record<string, string>;
  readiness: Record<string, unknown>;
  portfolio_item: Record<string, unknown>;
  acceptance_checklist: Record<string, unknown>;
  authoring_plan?: Record<string, unknown>;
  artifacts?: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingPlatformAdapterExportPackageMatrix {
  type: 'meeting_platform_adapter_export_package_matrix';
  schema: 'meeting_platform_adapter_export_package_matrix';
  schema_version: 1;
  target: string;
  platform_count: number;
  accepted_count: number;
  export_ready_count: number;
  built_in_count: number;
  custom_authoring_count: number;
  browser_extension_ready_count: number;
  provider_reconcile_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  packages: MeetingPlatformAdapterExportPackage[];
  portfolio: Record<string, unknown>;
  acceptance_checklist_matrix: Record<string, unknown>;
  next_actions: string[];
}

export function buildMeetingPlatformAdapterExportPackage(
  platform: string,
  input?: MeetingPlatformAdapterExportPackageOptions,
  options?: MeetingPlatformAdapterExportPackageOptions,
): MeetingPlatformAdapterExportPackage;

export function buildMeetingPlatformAdapterExportPackageMatrix(
  input?: MeetingPlatformAdapterExportPackageOptions,
  options?: MeetingPlatformAdapterExportPackageOptions,
): MeetingPlatformAdapterExportPackageMatrix;
