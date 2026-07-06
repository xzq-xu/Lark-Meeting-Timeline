import type {
  MeetingAppAdapterCapabilityMatrix,
  MeetingAppAdapterCapabilityOptions,
  MeetingAppAdapterCapabilityReport,
  MeetingAppAdapterExecutionGate,
  MeetingAppAdapterExecutionPlan,
} from './meeting-app-adapter-capability.mjs';
import type {
  MeetingAppAdapterHandoffPackage,
} from './meeting-app-adapter-handoff-package.mjs';

export const MEETING_APP_ADAPTER_INTEGRATION_PACKAGE_SCHEMA: 'meeting_app_adapter_integration_package';
export const MEETING_APP_ADAPTER_INTEGRATION_PACKAGE_MATRIX_SCHEMA: 'meeting_app_adapter_integration_package_matrix';
export const MEETING_APP_ADAPTER_INTEGRATION_PACKAGE_SCHEMA_VERSION: 1;

export interface MeetingAppAdapterIntegrationReadiness {
  accepted: boolean;
  static_ready: boolean;
  pilot_ready: boolean;
  production_ready: boolean;
  realtime_ready: boolean;
  first_blocked_step?: string;
  risk_level?: string;
  recommended_mode?: string;
  missing_evidence_count?: number;
}

export interface MeetingAppAdapterIntegrationStep {
  id: string;
  source?: string;
  status?: string;
  role?: string;
  required_for_realtime?: boolean;
  required_for_pilot?: boolean;
  required_for_production?: boolean;
  sdk_modules?: string[];
  evidence?: string[];
}

export interface MeetingAppAdapterIntegrationPackage {
  type: 'meeting_app_adapter_integration_package';
  schema: 'meeting_app_adapter_integration_package';
  schema_version: 1;
  platform: string;
  display_name?: string;
  accepted: boolean;
  static_ready: boolean;
  pilot_ready: boolean;
  production_ready: boolean;
  realtime_ready: boolean;
  recommended_mode?: string;
  risk_level?: string;
  first_blocked_step?: string;
  artifact_count: number;
  file_count: number;
  files: Array<Record<string, unknown>>;
  file_paths: string[];
  entrypoints: Record<string, string>;
  commands: Record<string, string>;
  evidence_contract: Record<string, unknown>;
  integration_steps: MeetingAppAdapterIntegrationStep[];
  gates: MeetingAppAdapterExecutionGate[];
  readiness: MeetingAppAdapterIntegrationReadiness;
  handoff_package: MeetingAppAdapterHandoffPackage;
  capability_report: MeetingAppAdapterCapabilityReport;
  execution_plan: MeetingAppAdapterExecutionPlan;
  next_actions: string[];
}

export interface MeetingAppAdapterIntegrationPackageMatrix {
  type: 'meeting_app_adapter_integration_package_matrix';
  schema: 'meeting_app_adapter_integration_package_matrix';
  schema_version: 1;
  accepted: boolean;
  platform_count: number;
  accepted_count: number;
  static_ready_count: number;
  pilot_ready_count: number;
  production_ready_count: number;
  realtime_ready_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  capability_matrix: MeetingAppAdapterCapabilityMatrix;
  packages: MeetingAppAdapterIntegrationPackage[];
  next_actions: string[];
}

export function buildMeetingAppAdapterIntegrationPackage(
  platformOrCapability?: string | MeetingAppAdapterCapabilityReport | MeetingAppAdapterExecutionPlan | MeetingAppAdapterCapabilityOptions,
  options?: MeetingAppAdapterCapabilityOptions,
): MeetingAppAdapterIntegrationPackage;

export function buildMeetingAppAdapterIntegrationPackageMatrix(
  options?: MeetingAppAdapterCapabilityOptions & { capabilityMatrix?: MeetingAppAdapterCapabilityMatrix },
): MeetingAppAdapterIntegrationPackageMatrix;
