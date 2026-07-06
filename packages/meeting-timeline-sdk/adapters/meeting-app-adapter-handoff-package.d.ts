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
export const MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA_VERSION: 1;

export interface MeetingAppAdapterHandoffPackageOptions extends MeetingAppAdapterRuntimeConfigOptions {}

export interface MeetingAppAdapterHandoffPackageMatrixOptions extends MeetingAppAdapterSpecMatrixOptions {
  packages?: Iterable<string | MeetingAppAdapterSpec | MeetingAppAdapterSpecInput> | Array<string | MeetingAppAdapterSpec | MeetingAppAdapterSpecInput>;
  configs?: Iterable<string | MeetingAppAdapterSpec | MeetingAppAdapterSpecInput> | Array<string | MeetingAppAdapterSpec | MeetingAppAdapterSpecInput>;
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

export function assertMeetingAppAdapterHandoffPackageMatrix(
  matrixOrOptions?: MeetingAppAdapterHandoffPackageMatrix | MeetingAppAdapterHandoffPackageMatrixOptions,
  options?: MeetingAppAdapterHandoffPackageMatrixOptions,
): MeetingAppAdapterHandoffPackageMatrix;
