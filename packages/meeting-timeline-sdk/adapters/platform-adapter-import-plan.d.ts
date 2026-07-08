import type { MeetingPlatformAdapterExportPackage } from './platform-adapter-export-package.mjs';

export const MEETING_PLATFORM_ADAPTER_IMPORT_PLAN_SCHEMA: 'meeting_platform_adapter_import_plan';
export const MEETING_PLATFORM_ADAPTER_IMPORT_PLAN_MATRIX_SCHEMA: 'meeting_platform_adapter_import_plan_matrix';
export const MEETING_PLATFORM_ADAPTER_IMPORT_PLAN_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdapterImportPlanOptions {
  target?: 'static' | 'pilot' | 'production' | string;
  acceptanceTarget?: 'static' | 'pilot' | 'production' | string;
  acceptance_target?: 'static' | 'pilot' | 'production' | string;
  surface?: 'browser_extension' | 'webview_preload' | 'native_detector' | 'native_host' | 'provider_reconcile' | string;
  preferredSurface?: string;
  preferred_surface?: string;
  availableFiles?: string[];
  available_files?: string[];
  allowCustomAuthoring?: boolean;
  allow_custom_authoring?: boolean;
}

export interface MeetingPlatformAdapterImportPlan {
  type: 'meeting_platform_adapter_import_plan';
  schema: 'meeting_platform_adapter_import_plan';
  schema_version: 1;
  platform: string;
  display_name?: string;
  target: string;
  selected_surface: string;
  accepted: boolean;
  import_ready: boolean;
  built_in: boolean;
  export_package_schema?: string;
  export_package_target?: string;
  export_package_accepted: boolean;
  runtime_contract: Record<string, unknown>;
  adapter_blueprint: Record<string, unknown>;
  raw_signal_validation: Record<string, unknown>;
  adapter_preflight: Record<string, unknown>;
  sdk_imports: Record<string, string>;
  host_file_coverage: Record<string, unknown>;
  surface_entrypoints?: Record<string, Record<string, unknown>>;
  surface_entrypoint?: Record<string, unknown>;
  install_steps: Array<Record<string, unknown>>;
  commands: Record<string, string>;
  readiness: Record<string, unknown>;
  source_package: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingPlatformAdapterImportPlanMatrix {
  type: 'meeting_platform_adapter_import_plan_matrix';
  schema: 'meeting_platform_adapter_import_plan_matrix';
  schema_version: 1;
  target: string;
  package_count: number;
  accepted_count: number;
  blocked_count: number;
  file_coverage_checked_count: number;
  missing_file_count: number;
  adapter_preflight_startup_ready_count: number;
  adapter_preflight_realtime_ready_count: number;
  raw_signal_validation_ready_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  plans: MeetingPlatformAdapterImportPlan[];
  next_actions: string[];
}

export function buildMeetingPlatformAdapterImportPlan(
  exportPackage: MeetingPlatformAdapterExportPackage | Record<string, unknown>,
  input?: MeetingPlatformAdapterImportPlanOptions,
  options?: MeetingPlatformAdapterImportPlanOptions,
): MeetingPlatformAdapterImportPlan;

export function assertMeetingPlatformAdapterImportPlan(
  planOrPackage: MeetingPlatformAdapterImportPlan | MeetingPlatformAdapterExportPackage | Record<string, unknown>,
  input?: MeetingPlatformAdapterImportPlanOptions,
  options?: MeetingPlatformAdapterImportPlanOptions,
): MeetingPlatformAdapterImportPlan;

export function buildMeetingPlatformAdapterImportPlanMatrix(
  packagesOrInput?: MeetingPlatformAdapterExportPackage[] | { packages?: MeetingPlatformAdapterExportPackage[]; exportPackages?: MeetingPlatformAdapterExportPackage[]; export_packages?: MeetingPlatformAdapterExportPackage[] } | Record<string, unknown>,
  input?: MeetingPlatformAdapterImportPlanOptions,
  options?: MeetingPlatformAdapterImportPlanOptions,
): MeetingPlatformAdapterImportPlanMatrix;
