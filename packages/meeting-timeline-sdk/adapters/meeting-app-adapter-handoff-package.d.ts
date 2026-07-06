import type {
  MeetingAppAdapterRuntimeConfig,
  MeetingAppAdapterRuntimeConfigOptions,
} from './meeting-app-adapter-runtime-config.mjs';
import type {
  MeetingAppAdapterSpec,
  MeetingAppAdapterSpecInput,
  MeetingAppAdapterSpecMatrixOptions,
} from './meeting-app-adapter-spec.mjs';
import type {
  MeetingAppAdapterManifest,
} from './meeting-app-adapter-manifest.mjs';

export const MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA: 'meeting_app_adapter_handoff_package';
export const MEETING_APP_ADAPTER_HANDOFF_PACKAGE_MATRIX_SCHEMA: 'meeting_app_adapter_handoff_package_matrix';
export const MEETING_APP_ADAPTER_VERIFICATION_PLAN_SCHEMA: 'meeting_app_adapter_verification_plan';
export const MEETING_APP_ADAPTER_VERIFICATION_REPORT_SCHEMA: 'meeting_app_adapter_verification_report';
export const MEETING_APP_ADAPTER_VERIFICATION_REPORT_MATRIX_SCHEMA: 'meeting_app_adapter_verification_report_matrix';
export const MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA_VERSION: 1;

export interface MeetingAppAdapterHandoffPackageOptions extends MeetingAppAdapterRuntimeConfigOptions {}

export interface MeetingAppAdapterHandoffPackageMatrixOptions extends MeetingAppAdapterSpecMatrixOptions {
  packages?: Iterable<string | MeetingAppAdapterSpec | MeetingAppAdapterSpecInput> | Array<string | MeetingAppAdapterSpec | MeetingAppAdapterSpecInput>;
  configs?: Iterable<string | MeetingAppAdapterSpec | MeetingAppAdapterSpecInput> | Array<string | MeetingAppAdapterSpec | MeetingAppAdapterSpecInput>;
}

export interface MeetingAppAdapterVerificationOptions extends MeetingAppAdapterHandoffPackageOptions {
  evidence?: Record<string, unknown>;
  evidence_package?: Record<string, unknown>;
  evidencePackage?: Record<string, unknown>;
  live_evidence?: Record<string, unknown>;
  liveEvidence?: Record<string, unknown>;
  target?: 'pilot' | 'production' | string;
  acceptance_target?: 'pilot' | 'production' | string;
  acceptanceTarget?: 'pilot' | 'production' | string;
  requireProductionReady?: boolean;
  require_production_ready?: boolean;
}

export interface MeetingAppAdapterVerificationMatrixOptions extends MeetingAppAdapterHandoffPackageMatrixOptions, MeetingAppAdapterVerificationOptions {
  reports?: Iterable<MeetingAppAdapterVerificationReport> | MeetingAppAdapterVerificationReport[];
  verification_reports?: Iterable<MeetingAppAdapterVerificationReport> | MeetingAppAdapterVerificationReport[];
  verificationReports?: Iterable<MeetingAppAdapterVerificationReport> | MeetingAppAdapterVerificationReport[];
  evidence_by_adapter?: Record<string, Record<string, unknown>>;
  evidenceByAdapter?: Record<string, Record<string, unknown>>;
}

export interface MeetingAppAdapterHandoffPackageFile {
  path: string;
  media_type: 'application/json' | 'text/markdown' | string;
  schema: string;
  content: unknown;
}

export interface MeetingAppAdapterHandoffPackageIssue {
  severity: 'error' | 'warning' | 'info' | string;
  code: string;
  message: string;
  [key: string]: unknown;
}

export interface MeetingAppAdapterVerificationPlan {
  type: 'meeting_app_adapter_verification_plan';
  schema: typeof MEETING_APP_ADAPTER_VERIFICATION_PLAN_SCHEMA;
  schema_version: typeof MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA_VERSION;
  adapter_key: string;
  display_name: string;
  source: string;
  surface: string;
  generated_from: Record<string, unknown>;
  commands: Record<string, string>;
  acceptance_policy: Record<string, unknown>;
  required_evidence: Array<Record<string, unknown>>;
  acceptance_checks: Array<Record<string, unknown>>;
  next_actions: string[];
}

export interface MeetingAppAdapterHandoffPackage {
  type: 'meeting_app_adapter_handoff_package';
  schema: typeof MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA;
  schema_version: typeof MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA_VERSION;
  accepted: boolean;
  adapter_key: string;
  display_name: string;
  source: string;
  surface: string;
  artifact_count: number;
  file_count: number;
  files: MeetingAppAdapterHandoffPackageFile[];
  file_paths: string[];
  adapter_spec: MeetingAppAdapterSpec;
  adapter_manifest?: MeetingAppAdapterManifest;
  runtime_config: MeetingAppAdapterRuntimeConfig;
  extension_manifest_fragment: Record<string, unknown>;
  verification_plan: MeetingAppAdapterVerificationPlan;
  consumer_entrypoints: Record<string, string>;
  contracts: Record<string, unknown>;
  validation: Record<string, unknown>;
  implementation_steps: string[];
  issue_count: number;
  blocking_count: number;
  warning_count: number;
  issues: MeetingAppAdapterHandoffPackageIssue[];
  next_actions: string[];
}

export interface MeetingAppAdapterVerificationReport {
  type: 'meeting_app_adapter_verification_report';
  schema: typeof MEETING_APP_ADAPTER_VERIFICATION_REPORT_SCHEMA;
  schema_version: typeof MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA_VERSION;
  accepted: boolean;
  target: string;
  adapter_key: string;
  display_name: string;
  static_ready: boolean;
  live_evidence_ready: boolean;
  pilot_ready: boolean;
  production_ready: boolean;
  check_count: number;
  passed_check_count: number;
  missing_evidence_count: number;
  checks: Array<Record<string, unknown>>;
  evidence_summary: Record<string, number>;
  handoff_package: Record<string, unknown>;
  verification_plan: MeetingAppAdapterVerificationPlan;
  issue_count: number;
  blocking_count: number;
  warning_count: number;
  issues: MeetingAppAdapterHandoffPackageIssue[];
  next_actions: string[];
}

export interface MeetingAppAdapterVerificationReportMatrix {
  type: 'meeting_app_adapter_verification_report_matrix';
  schema: typeof MEETING_APP_ADAPTER_VERIFICATION_REPORT_MATRIX_SCHEMA;
  schema_version: typeof MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA_VERSION;
  accepted: boolean;
  target: string;
  report_count: number;
  accepted_count: number;
  pilot_ready_count: number;
  production_ready_count: number;
  missing_evidence_count: number;
  rows: Array<Record<string, unknown>>;
  reports: MeetingAppAdapterVerificationReport[];
  next_actions: string[];
}

export interface MeetingAppAdapterHandoffPackageMatrix {
  type: 'meeting_app_adapter_handoff_package_matrix';
  schema: typeof MEETING_APP_ADAPTER_HANDOFF_PACKAGE_MATRIX_SCHEMA;
  schema_version: typeof MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA_VERSION;
  accepted: boolean;
  package_count: number;
  accepted_count: number;
  custom_count: number;
  built_in_count: number;
  runtime_config_ready_count: number;
  content_script_ready_count: number;
  live_evidence_required_count: number;
  rows: Array<Record<string, unknown>>;
  packages: MeetingAppAdapterHandoffPackage[];
  next_actions: string[];
}

export function buildMeetingAppAdapterHandoffPackage(
  specOrPlatform?: MeetingAppAdapterSpec | string | MeetingAppAdapterSpecInput,
  options?: MeetingAppAdapterHandoffPackageOptions,
): MeetingAppAdapterHandoffPackage;

export function buildMeetingAppAdapterHandoffPackageMatrix(
  options?: MeetingAppAdapterHandoffPackageMatrixOptions,
): MeetingAppAdapterHandoffPackageMatrix;

export function assertMeetingAppAdapterHandoffPackage(
  packageOrSpec?: MeetingAppAdapterHandoffPackage | MeetingAppAdapterSpec | string | MeetingAppAdapterSpecInput,
  options?: MeetingAppAdapterHandoffPackageOptions,
): MeetingAppAdapterHandoffPackage;

export function buildMeetingAppAdapterVerificationReport(
  packageOrSpec?: MeetingAppAdapterHandoffPackage | MeetingAppAdapterSpec | string | MeetingAppAdapterSpecInput,
  options?: MeetingAppAdapterVerificationOptions,
): MeetingAppAdapterVerificationReport;

export function buildMeetingAppAdapterVerificationReportMatrix(
  options?: MeetingAppAdapterVerificationMatrixOptions,
): MeetingAppAdapterVerificationReportMatrix;

export function assertMeetingAppAdapterVerificationReport(
  reportOrPackage?: MeetingAppAdapterVerificationReport | MeetingAppAdapterHandoffPackage | MeetingAppAdapterSpec | string | MeetingAppAdapterSpecInput,
  options?: MeetingAppAdapterVerificationOptions,
): MeetingAppAdapterVerificationReport;

export function assertMeetingAppAdapterVerificationReportMatrix(
  matrixOrOptions?: MeetingAppAdapterVerificationReportMatrix | MeetingAppAdapterVerificationMatrixOptions,
  options?: MeetingAppAdapterVerificationMatrixOptions,
): MeetingAppAdapterVerificationReportMatrix;

export function assertMeetingAppAdapterHandoffPackageMatrix(
  matrixOrOptions?: MeetingAppAdapterHandoffPackageMatrix | MeetingAppAdapterHandoffPackageMatrixOptions,
  options?: MeetingAppAdapterHandoffPackageMatrixOptions,
): MeetingAppAdapterHandoffPackageMatrix;
