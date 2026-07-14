export const MEETING_APP_ADAPTER_SPEC_SCHEMA: 'meeting_app_adapter_spec';
export const MEETING_APP_ADAPTER_SPEC_MATRIX_SCHEMA: 'meeting_app_adapter_spec_matrix';
export const MEETING_APP_ADAPTER_SPEC_SCHEMA_VERSION: 1;

export interface MeetingAppAdapterSpecInput {
  adapter_key?: string;
  adapterKey?: string;
  platform?: string;
  provider?: string;
  key?: string;
  name?: string;
  display_name?: string;
  displayName?: string;
  source?: string;
  surface?: string;
  matches?: Iterable<string> | string[];
  url_matches?: Iterable<string> | string[];
  urlMatches?: Iterable<string> | string[];
  host_permissions?: Iterable<string> | string[];
  hostPermissions?: Iterable<string> | string[];
  permissions?: Iterable<string> | string[];
  control_selectors?: Iterable<string> | string[];
  controlSelectors?: Iterable<string> | string[];
  participant_selectors?: Iterable<string> | string[];
  participantSelectors?: Iterable<string> | string[];
  text_selectors?: Iterable<string> | string[];
  textSelectors?: Iterable<string> | string[];
  mutation_track_selectors?: Iterable<string> | string[];
  mutationTrackSelectors?: Iterable<string> | string[];
  mutation_ignore_selectors?: Iterable<string> | string[];
  mutationIgnoreSelectors?: Iterable<string> | string[];
  required_signals?: Iterable<string> | string[];
  requiredSignals?: Iterable<string> | string[];
  provider_events_block_realtime?: boolean;
  providerEventsBlockRealtime?: boolean;
  transcript_blocks_realtime?: boolean;
  transcriptBlocksRealtime?: boolean;
  includeBuiltInManifest?: boolean;
  include_built_in_manifest?: boolean;
  [key: string]: unknown;
}

export interface MeetingAppAdapterSpecMatrixOptions extends MeetingAppAdapterSpecInput {
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  platformKeys?: Iterable<string> | string[];
  adapters?: Iterable<string | MeetingAppAdapterSpecInput> | Array<string | MeetingAppAdapterSpecInput>;
  adapter_specs?: Iterable<string | MeetingAppAdapterSpecInput> | Array<string | MeetingAppAdapterSpecInput>;
  adapterSpecs?: Iterable<string | MeetingAppAdapterSpecInput> | Array<string | MeetingAppAdapterSpecInput>;
  specs?: Iterable<string | MeetingAppAdapterSpecInput> | Array<string | MeetingAppAdapterSpecInput>;
  includeBuiltIns?: boolean;
  include_built_ins?: boolean;
}

export interface MeetingAppAdapterSpecIssue {
  severity: 'error' | 'warning' | 'info' | string;
  code: string;
  message: string;
  [key: string]: unknown;
}

export interface MeetingAppAdapterSpec {
  type: 'meeting_app_adapter_spec';
  schema: typeof MEETING_APP_ADAPTER_SPEC_SCHEMA;
  schema_version: typeof MEETING_APP_ADAPTER_SPEC_SCHEMA_VERSION;
  accepted: boolean;
  adapter_key: string;
  platform: string;
  display_name: string;
  source: string;
  surface: string;
  url_detection: Record<string, unknown>;
  extension: Record<string, unknown>;
  capture: Record<string, unknown>;
  runtime: Record<string, unknown>;
  contracts: Record<string, unknown>;
  host_endpoints: Record<string, string>;
  implementation_steps: string[];
  readiness: Record<string, unknown>;
  built_in_manifest?: Record<string, unknown>;
  issue_count: number;
  blocking_count: number;
  warning_count: number;
  issues: MeetingAppAdapterSpecIssue[];
  next_actions: string[];
}

export interface MeetingAppAdapterSpecMatrix {
  type: 'meeting_app_adapter_spec_matrix';
  schema: typeof MEETING_APP_ADAPTER_SPEC_MATRIX_SCHEMA;
  schema_version: typeof MEETING_APP_ADAPTER_SPEC_SCHEMA_VERSION;
  accepted: boolean;
  spec_count: number;
  accepted_count: number;
  custom_count: number;
  built_in_count: number;
  url_match_ready_count: number;
  capture_selector_ready_count: number;
  mutation_observer_ready_count: number;
  rows: Array<Record<string, unknown>>;
  specs: MeetingAppAdapterSpec[];
  next_actions: string[];
}

export function buildMeetingAppAdapterSpec(
  specOrPlatform?: string | MeetingAppAdapterSpecInput,
  options?: MeetingAppAdapterSpecInput,
): MeetingAppAdapterSpec;

export function buildMeetingAppAdapterSpecMatrix(
  options?: MeetingAppAdapterSpecMatrixOptions,
): MeetingAppAdapterSpecMatrix;

export function buildMeetingAppAdapterSpecTemplate(
  input?: MeetingAppAdapterSpecInput,
): MeetingAppAdapterSpecInput;

export function assertMeetingAppAdapterSpec(
  specOrPlatform?: MeetingAppAdapterSpec | string | MeetingAppAdapterSpecInput,
  options?: MeetingAppAdapterSpecInput,
): MeetingAppAdapterSpec;

export function assertMeetingAppAdapterSpecMatrix(
  matrixOrOptions?: MeetingAppAdapterSpecMatrix | MeetingAppAdapterSpecMatrixOptions,
  options?: MeetingAppAdapterSpecMatrixOptions,
): MeetingAppAdapterSpecMatrix;

export function builtInMeetingAppAdapterSpecMatrix(
  options?: MeetingAppAdapterSpecMatrixOptions,
): MeetingAppAdapterSpecMatrix;
