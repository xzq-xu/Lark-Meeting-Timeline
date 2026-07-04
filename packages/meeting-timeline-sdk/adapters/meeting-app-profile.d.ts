import type { MeetingAppDomCaptureProfilePlatform } from './meeting-app-capture.mjs';
import type { MeetingAppLaunchGateOptions } from './meeting-app-gate.mjs';

export const MEETING_APP_INTEGRATION_PROFILE_SCHEMA: 'meeting_app_integration_profile';
export const MEETING_APP_INTEGRATION_PROFILE_SCHEMA_VERSION: number;
export const MEETING_APP_INTEGRATION_PROFILE_PLATFORMS: readonly MeetingAppDomCaptureProfilePlatform[];
export const MEETING_APP_RUNTIME_ADAPTER_CONFIG_SCHEMA: 'meeting_app_runtime_adapter_config';
export const MEETING_APP_RUNTIME_ADAPTER_PROFILE_RESOLUTION_SCHEMA: 'meeting_app_runtime_adapter_profile_resolution';
export const MEETING_APP_RUNTIME_ADAPTER_PROFILE_MATRIX_SCHEMA: 'meeting_app_runtime_adapter_profile_matrix';
export const MEETING_APP_RUNTIME_ADAPTER_SELECTION_SCHEMA: 'meeting_app_runtime_adapter_selection';
export const MEETING_APP_RUNTIME_ADAPTER_HANDOFF_SCHEMA: 'meeting_app_runtime_adapter_handoff';
export const MEETING_APP_LIVE_SNAPSHOT_CAPTURE_PLAN_SCHEMA: 'meeting_app_live_snapshot_capture_plan';
export const MEETING_APP_DEPLOYMENT_MANIFEST_SCHEMA: 'meeting_app_deployment_manifest';
export const MEETING_APP_LIVE_EVIDENCE_PACKAGE_SCHEMA: 'meeting_app_live_evidence_package';
export const MEETING_APP_DOM_ADAPTATION_DIAGNOSIS_SCHEMA: 'meeting_app_dom_adaptation_diagnosis';
export const MEETING_APP_DOM_ADAPTATION_DIAGNOSIS_MATRIX_SCHEMA: 'meeting_app_dom_adaptation_diagnosis_matrix';

export interface MeetingAppIntegrationProfileOptions extends MeetingAppLaunchGateOptions {
  platform?: MeetingAppDomCaptureProfilePlatform | string;
  provider?: MeetingAppDomCaptureProfilePlatform | string;
  key?: MeetingAppDomCaptureProfilePlatform | string;
  name?: MeetingAppDomCaptureProfilePlatform | string;
  platforms?: Array<MeetingAppDomCaptureProfilePlatform | string> | MeetingAppDomCaptureProfilePlatform | string;
  platform_keys?: Array<MeetingAppDomCaptureProfilePlatform | string> | MeetingAppDomCaptureProfilePlatform | string;
  platformKeys?: Array<MeetingAppDomCaptureProfilePlatform | string> | MeetingAppDomCaptureProfilePlatform | string;
  includeLaunchGate?: boolean;
  include_launch_gate?: boolean;
  launchGateOptions?: MeetingAppLaunchGateOptions;
  launch_gate_options?: MeetingAppLaunchGateOptions;
  [key: string]: unknown;
}

export interface MeetingAppIntegrationProfile {
  type: 'meeting_app_integration_profile';
  schema: 'meeting_app_integration_profile';
  version: number;
  platform: MeetingAppDomCaptureProfilePlatform;
  display_name: string;
  recommended_mode: string;
  extension: Record<string, unknown>;
  runtime: Record<string, unknown>;
  capture: Record<string, unknown>;
  event_model: Record<string, unknown>;
  provider_reconciliation: Record<string, unknown>;
  readiness?: Record<string, unknown>;
  launch_gate?: Record<string, unknown>;
  implementation_steps: string[];
  limitations: string[];
}

export interface MeetingAppIntegrationMatrix {
  type: 'meeting_app_integration_matrix';
  schema: 'meeting_app_integration_profile';
  version: number;
  platform_count: number;
  platforms: MeetingAppDomCaptureProfilePlatform[];
  rows: Array<Record<string, unknown>>;
}

export interface MeetingAppRuntimeAdapterConfig {
  type: 'meeting_app_runtime_adapter_config';
  schema: 'meeting_app_runtime_adapter_config';
  version: number;
  platform: MeetingAppDomCaptureProfilePlatform;
  display_name: string;
  source: string;
  extension: Record<string, unknown>;
  bridge_options: Record<string, unknown>;
  runtime_options: Record<string, unknown>;
  capture_options: Record<string, unknown>;
  startup: Record<string, unknown>;
  supported_client_methods: string[];
  readiness: Record<string, unknown>;
}

export interface MeetingAppRuntimeAdapterProfileResolution {
  type: 'meeting_app_runtime_adapter_profile_resolution';
  schema: 'meeting_app_runtime_adapter_profile_resolution';
  version: number;
  detected: boolean;
  reason?: string;
  platform: MeetingAppDomCaptureProfilePlatform | null;
  display_name?: string;
  url?: string;
  title?: string;
  extension?: Record<string, unknown>;
  runtime_config?: MeetingAppRuntimeAdapterConfig;
  capture?: Record<string, unknown>;
  runtime?: Record<string, unknown>;
  tracks?: Record<string, unknown>;
  host?: Record<string, unknown>;
  readiness: Record<string, unknown>;
  issues?: Array<Record<string, unknown>>;
  next_actions: string[];
}

export interface MeetingAppRuntimeAdapterProfileMatrix {
  type: 'meeting_app_runtime_adapter_profile_matrix';
  schema: 'meeting_app_runtime_adapter_profile_matrix';
  version: number;
  platform_count: number;
  detected_count: number;
  runtime_ready_count: number;
  track_profile_count: number;
  platforms: MeetingAppDomCaptureProfilePlatform[];
  rows: Array<Record<string, unknown>>;
  profiles: MeetingAppRuntimeAdapterProfileResolution[];
  next_actions: string[];
}

export interface MeetingAppRuntimeAdapterSelection {
  type: 'meeting_app_runtime_adapter_selection';
  schema: 'meeting_app_runtime_adapter_selection';
  version: number;
  selected: boolean;
  detected: boolean;
  supported: boolean;
  platform?: MeetingAppDomCaptureProfilePlatform | null;
  display_name?: string;
  reason?: string;
  current_platforms: MeetingAppDomCaptureProfilePlatform[];
  candidate_count: number;
  supported_candidate_count: number;
  selected_candidate?: Record<string, unknown>;
  selected_meeting?: Record<string, unknown> | null;
  adapter?: Record<string, unknown>;
  launch?: Record<string, unknown>;
  readiness: Record<string, unknown>;
  profile: MeetingAppRuntimeAdapterProfileResolution;
  candidates: Array<Record<string, unknown>>;
  selection: Record<string, unknown>;
  issues: Array<Record<string, unknown>>;
  next_actions: string[];
}

export interface MeetingAppRuntimeAdapterHandoff {
  type: 'meeting_app_runtime_adapter_handoff';
  schema: 'meeting_app_runtime_adapter_handoff';
  version: number;
  selected: boolean;
  platform?: MeetingAppDomCaptureProfilePlatform | null;
  display_name?: string;
  surface: 'browser_extension' | 'electron_webview' | 'webview' | 'native_detector' | 'custom_host' | string;
  meeting?: Record<string, unknown> | null;
  adapter?: Record<string, unknown>;
  install: Record<string, unknown>;
  runtime?: Record<string, unknown>;
  tracks?: Record<string, unknown>;
  annotations: Record<string, unknown>;
  host_contract: Record<string, unknown>;
  readiness: Record<string, unknown>;
  selection: MeetingAppRuntimeAdapterSelection;
  issues: Array<Record<string, unknown>>;
  next_actions: string[];
}

export interface MeetingAppRuntimeAdapterAcceptanceReport {
  type: 'meeting_app_runtime_adapter_acceptance_report';
  schema: 'meeting_app_runtime_adapter_config';
  version: number;
  accepted: boolean;
  platform?: MeetingAppDomCaptureProfilePlatform | string;
  config: MeetingAppRuntimeAdapterConfig | Record<string, unknown>;
  coverage: Record<string, boolean>;
  issues: Array<Record<string, unknown>>;
}

export interface MeetingAppRuntimeAdapterValidationReport {
  type: 'meeting_app_runtime_adapter_validation_report';
  schema: 'meeting_app_runtime_adapter_config';
  version: number;
  accepted: boolean;
  production_ready: boolean;
  platform?: MeetingAppDomCaptureProfilePlatform | string;
  evidence_level: string;
  evidence_count: number;
  config_acceptance: MeetingAppRuntimeAdapterAcceptanceReport;
  launch_gate?: Record<string, unknown> | null;
  next_actions: string[];
  issues: Array<Record<string, unknown>>;
}

export interface MeetingAppDomAdaptationDiagnosis {
  type: 'meeting_app_dom_adaptation_diagnosis';
  schema: 'meeting_app_dom_adaptation_diagnosis';
  version: number;
  platform: MeetingAppDomCaptureProfilePlatform;
  display_name: string;
  accepted: boolean;
  production_ready: boolean;
  evidence_level: string;
  evidence_count: number;
  record_count: number;
  phases: string[];
  selector_probe: Record<string, unknown>;
  observer_probe: Record<string, unknown>;
  runtime_probe: Record<string, unknown>;
  recommended_capture: Record<string, unknown>;
  next_actions: string[];
  issues: Array<Record<string, unknown>>;
}

export interface MeetingAppDomAdaptationDiagnosisMatrix {
  type: 'meeting_app_dom_adaptation_diagnosis_matrix';
  schema: 'meeting_app_dom_adaptation_diagnosis_matrix';
  version: number;
  platform_count: number;
  accepted_count: number;
  production_ready_count: number;
  active_speaker_ready_count: number;
  meeting_start_ready_count: number;
  meeting_end_ready_count: number;
  platforms: MeetingAppDomCaptureProfilePlatform[];
  rows: Array<Record<string, unknown>>;
  diagnoses: MeetingAppDomAdaptationDiagnosis[];
  next_actions: string[];
}

export interface MeetingAppLiveSnapshotCapturePlan {
  type: 'meeting_app_live_snapshot_capture_plan';
  schema: 'meeting_app_live_snapshot_capture_plan';
  version: number;
  platform: MeetingAppDomCaptureProfilePlatform;
  display_name: string;
  runtime_config: MeetingAppRuntimeAdapterConfig;
  recorder: Record<string, unknown>;
  validation: Record<string, unknown>;
  required_snapshots: Array<Record<string, unknown>>;
  recommended_snapshots: Array<Record<string, unknown>>;
  minimum_record_count: number;
  handoff: Record<string, unknown>;
}

export interface MeetingAppDeploymentManifest {
  type: 'meeting_app_deployment_manifest';
  schema: 'meeting_app_deployment_manifest';
  version: number;
  platform: MeetingAppDomCaptureProfilePlatform;
  display_name: string;
  recommended_mode: string;
  profile: MeetingAppIntegrationProfile;
  runtime_config: MeetingAppRuntimeAdapterConfig;
  extension_install_plan: Record<string, unknown>;
  live_snapshot_capture_plan: MeetingAppLiveSnapshotCapturePlan;
  validation_report: MeetingAppRuntimeAdapterValidationReport;
  runtime_contract: Record<string, unknown>;
  integration_targets: Array<Record<string, unknown>>;
  production_gate: Record<string, unknown>;
  handoff: Record<string, unknown>;
  rollout_checklist: string[];
}

export interface MeetingAppDeploymentManifestAcceptanceReport {
  type: 'meeting_app_deployment_manifest_acceptance_report';
  schema: 'meeting_app_deployment_manifest';
  version: number;
  accepted: boolean;
  production_ready: boolean;
  platform?: MeetingAppDomCaptureProfilePlatform | string;
  manifest: MeetingAppDeploymentManifest | Record<string, unknown>;
  coverage: Record<string, boolean>;
  issues: Array<Record<string, unknown>>;
}

export interface MeetingAppDeploymentManifestAcceptanceSummary {
  type: 'meeting_app_deployment_manifest_acceptance_summary';
  schema: 'meeting_app_deployment_manifest';
  version: number;
  accepted: boolean;
  production_ready: boolean;
  platform_count: number;
  accepted_count: number;
  production_ready_count: number;
  rows: Array<Record<string, unknown>>;
}

export interface MeetingAppLiveEvidencePackage {
  type: 'meeting_app_live_evidence_package';
  schema: 'meeting_app_live_evidence_package';
  version: number;
  id: string;
  createdAtMs: number;
  created_at_ms: number;
  source: string;
  label?: string;
  accepted: boolean;
  production_ready: boolean;
  platform_count: number;
  platforms: MeetingAppDomCaptureProfilePlatform[];
  record_count: number;
  record_set: Record<string, unknown>;
  records_by_platform: Partial<Record<MeetingAppDomCaptureProfilePlatform, Record<string, unknown>>>;
  manifest_acceptance: Partial<Record<MeetingAppDomCaptureProfilePlatform, MeetingAppDeploymentManifestAcceptanceReport>>;
  summary: Record<string, unknown>;
  issues: Array<Record<string, unknown>>;
  handoff: Record<string, unknown>;
}

export interface MeetingAppLiveEvidencePackageSummary {
  type: 'meeting_app_live_evidence_package_summary';
  schema: 'meeting_app_live_evidence_package';
  version: number;
  id: string;
  accepted: boolean;
  production_ready: boolean;
  platform_count: number;
  record_count: number;
  accepted_count: number;
  production_ready_count: number;
  rows: Array<Record<string, unknown>>;
  issues: Array<Record<string, unknown>>;
}

export function buildMeetingAppIntegrationProfile(
  platformOrInput?: MeetingAppDomCaptureProfilePlatform | string | MeetingAppIntegrationProfileOptions,
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppIntegrationProfile;

export function buildAllMeetingAppIntegrationProfiles(
  options?: MeetingAppIntegrationProfileOptions,
): Partial<Record<MeetingAppDomCaptureProfilePlatform, MeetingAppIntegrationProfile>>;

export function buildMeetingAppIntegrationMatrix(
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppIntegrationMatrix;

export function buildMeetingAppRuntimeAdapterConfig(
  platformOrInput?: MeetingAppDomCaptureProfilePlatform | string | MeetingAppIntegrationProfileOptions,
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppRuntimeAdapterConfig;

export function resolveMeetingAppRuntimeAdapterProfile(
  input?: string | Record<string, unknown>,
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppRuntimeAdapterProfileResolution;

export function buildMeetingAppRuntimeAdapterProfileMatrix(
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppRuntimeAdapterProfileMatrix;

export function selectMeetingAppRuntimeAdapter(
  input?: string | Record<string, unknown>,
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppRuntimeAdapterSelection;

export function buildMeetingAppRuntimeAdapterHandoff(
  selectionOrInput?: string | Record<string, unknown> | MeetingAppRuntimeAdapterSelection,
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppRuntimeAdapterHandoff;

export function buildAllMeetingAppRuntimeAdapterConfigs(
  options?: MeetingAppIntegrationProfileOptions,
): Partial<Record<MeetingAppDomCaptureProfilePlatform, MeetingAppRuntimeAdapterConfig>>;

export function buildMeetingAppRuntimeAdapterAcceptanceReport(
  configOrPlatform?: MeetingAppRuntimeAdapterConfig | MeetingAppDomCaptureProfilePlatform | string | MeetingAppIntegrationProfileOptions,
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppRuntimeAdapterAcceptanceReport;

export function assertMeetingAppRuntimeAdapterConfig(
  configOrPlatform?: MeetingAppRuntimeAdapterConfig | MeetingAppDomCaptureProfilePlatform | string | MeetingAppIntegrationProfileOptions,
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppRuntimeAdapterAcceptanceReport;

export function buildAllMeetingAppRuntimeAdapterAcceptanceReports(
  options?: MeetingAppIntegrationProfileOptions,
): Partial<Record<MeetingAppDomCaptureProfilePlatform, MeetingAppRuntimeAdapterAcceptanceReport>>;

export function buildMeetingAppLiveSnapshotCapturePlan(
  platformOrInput?: MeetingAppRuntimeAdapterConfig | MeetingAppDomCaptureProfilePlatform | string | MeetingAppIntegrationProfileOptions,
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppLiveSnapshotCapturePlan;

export function buildAllMeetingAppLiveSnapshotCapturePlans(
  options?: MeetingAppIntegrationProfileOptions,
): Partial<Record<MeetingAppDomCaptureProfilePlatform, MeetingAppLiveSnapshotCapturePlan>>;

export function buildMeetingAppDeploymentManifest(
  platformOrInput?: MeetingAppRuntimeAdapterConfig | MeetingAppDomCaptureProfilePlatform | string | MeetingAppIntegrationProfileOptions,
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppDeploymentManifest;

export function buildAllMeetingAppDeploymentManifests(
  options?: MeetingAppIntegrationProfileOptions,
): Partial<Record<MeetingAppDomCaptureProfilePlatform, MeetingAppDeploymentManifest>>;

export function buildMeetingAppDeploymentManifestAcceptanceReport(
  manifestOrPlatform?: MeetingAppDeploymentManifest | MeetingAppRuntimeAdapterConfig | MeetingAppDomCaptureProfilePlatform | string | MeetingAppIntegrationProfileOptions,
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppDeploymentManifestAcceptanceReport;

export function assertMeetingAppDeploymentManifest(
  manifestOrPlatform?: MeetingAppDeploymentManifest | MeetingAppRuntimeAdapterConfig | MeetingAppDomCaptureProfilePlatform | string | MeetingAppIntegrationProfileOptions,
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppDeploymentManifestAcceptanceReport;

export function buildAllMeetingAppDeploymentManifestAcceptanceReports(
  options?: MeetingAppIntegrationProfileOptions,
): Partial<Record<MeetingAppDomCaptureProfilePlatform, MeetingAppDeploymentManifestAcceptanceReport>>;

export function buildMeetingAppDeploymentManifestAcceptanceSummary(
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppDeploymentManifestAcceptanceSummary;

export function buildMeetingAppLiveEvidencePackage(
  input?: Record<string, unknown> | unknown[],
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppLiveEvidencePackage;

export function buildMeetingAppLiveEvidencePackageSummary(
  input?: MeetingAppLiveEvidencePackage | Record<string, unknown> | unknown[],
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppLiveEvidencePackageSummary;

export function buildMeetingAppDomAdaptationDiagnosis(
  platformOrInput?: MeetingAppDomCaptureProfilePlatform | string | MeetingAppIntegrationProfileOptions | Record<string, unknown> | unknown[],
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppDomAdaptationDiagnosis;

export function buildAllMeetingAppDomAdaptationDiagnoses(
  options?: MeetingAppIntegrationProfileOptions,
): Partial<Record<MeetingAppDomCaptureProfilePlatform, MeetingAppDomAdaptationDiagnosis>>;

export function buildMeetingAppDomAdaptationDiagnosisMatrix(
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppDomAdaptationDiagnosisMatrix;

export function buildMeetingAppRuntimeAdapterValidationReport(
  configOrPlatform?: MeetingAppRuntimeAdapterConfig | MeetingAppDomCaptureProfilePlatform | string | MeetingAppIntegrationProfileOptions,
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppRuntimeAdapterValidationReport;

export function assertMeetingAppRuntimeAdapterValidation(
  configOrPlatform?: MeetingAppRuntimeAdapterConfig | MeetingAppDomCaptureProfilePlatform | string | MeetingAppIntegrationProfileOptions,
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppRuntimeAdapterValidationReport;

export function buildAllMeetingAppRuntimeAdapterValidationReports(
  options?: MeetingAppIntegrationProfileOptions,
): Partial<Record<MeetingAppDomCaptureProfilePlatform, MeetingAppRuntimeAdapterValidationReport>>;

export default buildMeetingAppIntegrationProfile;
