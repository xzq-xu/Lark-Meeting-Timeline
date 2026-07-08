import type { MeetingPlatformAdapterExportPackage } from './platform-adapter-export-package.mjs';
import type { MeetingPlatformAdapterImportPlan, MeetingPlatformAdapterImportPlanOptions } from './platform-adapter-import-plan.mjs';

export const MEETING_PLATFORM_ADAPTER_INSTALL_MANIFEST_SCHEMA: 'meeting_platform_adapter_install_manifest';
export const MEETING_PLATFORM_ADAPTER_INSTALL_MANIFEST_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdapterInstallManifestOptions extends MeetingPlatformAdapterImportPlanOptions {
  baseUrl?: string;
  base_url?: string;
  installTarget?: string;
  install_target?: string;
  contentScriptJs?: string | string[];
  content_script_js?: string | string[];
  allowPartialInstall?: boolean;
  allow_partial_install?: boolean;
}

export interface MeetingPlatformAdapterInstallManifest {
  type: 'meeting_platform_adapter_install_manifest';
  schema: 'meeting_platform_adapter_install_manifest';
  schema_version: 1;
  target: string;
  install_target: string;
  base_url?: string;
  accepted: boolean;
  platform_count: number;
  ready_platform_count: number;
  blocked_platform_count: number;
  platforms: string[];
  selected_surfaces: string[];
  sdk_imports: Record<string, string>;
  runtime_contract: Record<string, unknown>;
  platform_registry: Array<Record<string, unknown>>;
  browser_extension: Record<string, unknown>;
  webview_preload: Record<string, unknown>;
  native_detector: Record<string, unknown>;
  native_host: Record<string, unknown>;
  adapter_blueprints: Record<string, unknown>;
  adapter_preflight: Record<string, unknown>;
  provider_reconcile: Record<string, unknown>;
  install_sequence: Array<Record<string, unknown>>;
  readiness: Record<string, unknown>;
  next_actions: string[];
}

export function buildMeetingPlatformAdapterInstallManifest(
  plansOrInput?: MeetingPlatformAdapterImportPlan[] | MeetingPlatformAdapterExportPackage[] | {
    plans?: MeetingPlatformAdapterImportPlan[];
    importPlans?: MeetingPlatformAdapterImportPlan[];
    import_plans?: MeetingPlatformAdapterImportPlan[];
    packages?: MeetingPlatformAdapterExportPackage[];
    exportPackages?: MeetingPlatformAdapterExportPackage[];
    export_packages?: MeetingPlatformAdapterExportPackage[];
    importPlanMatrix?: { plans?: MeetingPlatformAdapterImportPlan[] };
    import_plan_matrix?: { plans?: MeetingPlatformAdapterImportPlan[] };
  } | Record<string, unknown>,
  input?: MeetingPlatformAdapterInstallManifestOptions,
  options?: MeetingPlatformAdapterInstallManifestOptions,
): MeetingPlatformAdapterInstallManifest;

export function assertMeetingPlatformAdapterInstallManifest(
  manifestOrInput?: MeetingPlatformAdapterInstallManifest | MeetingPlatformAdapterImportPlan[] | MeetingPlatformAdapterExportPackage[] | Record<string, unknown>,
  input?: MeetingPlatformAdapterInstallManifestOptions,
  options?: MeetingPlatformAdapterInstallManifestOptions,
): MeetingPlatformAdapterInstallManifest;
