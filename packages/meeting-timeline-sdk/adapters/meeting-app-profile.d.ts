import type { MeetingAppDomCaptureProfilePlatform } from './meeting-app-capture.mjs';
import type { MeetingAppLaunchGateOptions } from './meeting-app-gate.mjs';

export const MEETING_APP_INTEGRATION_PROFILE_SCHEMA: 'meeting_app_integration_profile';
export const MEETING_APP_INTEGRATION_PROFILE_SCHEMA_VERSION: number;
export const MEETING_APP_INTEGRATION_PROFILE_PLATFORMS: readonly MeetingAppDomCaptureProfilePlatform[];
export const MEETING_APP_RUNTIME_ADAPTER_CONFIG_SCHEMA: 'meeting_app_runtime_adapter_config';

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
