import type {
  MeetingAppAdapterSpec,
  MeetingAppAdapterSpecInput,
  MeetingAppAdapterSpecMatrixOptions,
} from './meeting-app-adapter-spec.mjs';

export const MEETING_APP_ADAPTER_RUNTIME_CONFIG_SCHEMA: 'meeting_app_adapter_runtime_config';
export const MEETING_APP_ADAPTER_RUNTIME_CONFIG_MATRIX_SCHEMA: 'meeting_app_adapter_runtime_config_matrix';
export const MEETING_APP_ADAPTER_RUNTIME_CONFIG_SCHEMA_VERSION: 1;

export interface MeetingAppAdapterRuntimeConfigOptions extends MeetingAppAdapterSpecInput {
  maxControls?: number;
  max_controls?: number;
  maxParticipants?: number;
  max_participants?: number;
  maxTexts?: number;
  max_texts?: number;
  includeShadowDom?: boolean;
  include_shadow_dom?: boolean;
  extensionMessaging?: boolean;
  extension_messaging?: boolean;
  windowMessaging?: boolean;
  window_messaging?: boolean;
  allowedOrigins?: Iterable<string> | string[];
  allowed_origins?: Iterable<string> | string[];
  applyOptions?: Record<string, unknown>;
  apply_options?: Record<string, unknown>;
  speakerOptions?: Record<string, unknown>;
  speaker_options?: Record<string, unknown>;
  participantOptions?: Record<string, unknown>;
  participant_options?: Record<string, unknown>;
}

export interface MeetingAppAdapterRuntimeConfigMatrixOptions extends MeetingAppAdapterSpecMatrixOptions {
  configs?: Iterable<string | MeetingAppAdapterSpec | MeetingAppAdapterSpecInput> | Array<string | MeetingAppAdapterSpec | MeetingAppAdapterSpecInput>;
}

export interface MeetingAppAdapterRuntimeConfigIssue {
  severity: 'error' | 'warning' | 'info' | string;
  code: string;
  message: string;
  [key: string]: unknown;
}

export interface MeetingAppAdapterRuntimeConfig {
  type: 'meeting_app_adapter_runtime_config';
  schema: typeof MEETING_APP_ADAPTER_RUNTIME_CONFIG_SCHEMA;
  schema_version: typeof MEETING_APP_ADAPTER_RUNTIME_CONFIG_SCHEMA_VERSION;
  accepted: boolean;
  adapter_key: string;
  display_name: string;
  source_spec: Record<string, unknown>;
  entrypoints: Record<string, string>;
  extension: Record<string, unknown>;
  capture_options: Record<string, unknown>;
  browser_runtime_options: Record<string, unknown>;
  content_script_options: Record<string, unknown>;
  host_endpoints: Record<string, string>;
  contracts: Record<string, unknown>;
  usage: Record<string, string>;
  readiness: Record<string, unknown>;
  issue_count: number;
  blocking_count: number;
  warning_count: number;
  issues: MeetingAppAdapterRuntimeConfigIssue[];
  next_actions: string[];
}

export interface MeetingAppAdapterRuntimeConfigMatrix {
  type: 'meeting_app_adapter_runtime_config_matrix';
  schema: typeof MEETING_APP_ADAPTER_RUNTIME_CONFIG_MATRIX_SCHEMA;
  schema_version: typeof MEETING_APP_ADAPTER_RUNTIME_CONFIG_SCHEMA_VERSION;
  accepted: boolean;
  config_count: number;
  accepted_count: number;
  custom_count: number;
  built_in_count: number;
  capture_ready_count: number;
  mutation_ready_count: number;
  content_script_ready_count: number;
  rows: Array<Record<string, unknown>>;
  configs: MeetingAppAdapterRuntimeConfig[];
  next_actions: string[];
}

export function buildMeetingAppAdapterRuntimeConfig(
  specOrPlatform?: MeetingAppAdapterSpec | string | MeetingAppAdapterSpecInput,
  options?: MeetingAppAdapterRuntimeConfigOptions,
): MeetingAppAdapterRuntimeConfig;

export function buildMeetingAppAdapterRuntimeConfigMatrix(
  options?: MeetingAppAdapterRuntimeConfigMatrixOptions,
): MeetingAppAdapterRuntimeConfigMatrix;

export function assertMeetingAppAdapterRuntimeConfig(
  configOrSpec?: MeetingAppAdapterRuntimeConfig | MeetingAppAdapterSpec | string | MeetingAppAdapterSpecInput,
  options?: MeetingAppAdapterRuntimeConfigOptions,
): MeetingAppAdapterRuntimeConfig;

export function assertMeetingAppAdapterRuntimeConfigMatrix(
  matrixOrOptions?: MeetingAppAdapterRuntimeConfigMatrix | MeetingAppAdapterRuntimeConfigMatrixOptions,
  options?: MeetingAppAdapterRuntimeConfigMatrixOptions,
): MeetingAppAdapterRuntimeConfigMatrix;
