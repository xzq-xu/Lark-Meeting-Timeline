export const MEETING_PLATFORM_HOST_INTEGRATION_SCHEMA: string;
export const MEETING_PLATFORM_HOST_INTEGRATION_SCHEMA_VERSION: number;
export const MEETING_PLATFORM_HOST_INTEGRATION_SCAFFOLD_SCHEMA: string;
export const MEETING_PLATFORM_HOST_INTEGRATION_ACCEPTANCE_SCHEMA: string;

export interface MeetingPlatformHostIntegrationOptions {
  baseUrl?: string;
  base_url?: string;
  basePath?: string;
  base_path?: string;
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  packageName?: string;
  package_name?: string;
  env?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingPlatformHostIntegrationPlan {
  type: 'meeting_platform_host_integration';
  schema: string;
  schema_version: number;
  base_url: string;
  base_path: string;
  platforms: string[];
  platform_imports: Array<Record<string, string>>;
  sdk: Record<string, unknown>;
  runtime_contract: Record<string, unknown>;
  endpoints: Record<string, string>;
  commands: Record<string, string>;
  evidence_paths: Record<string, string>;
  integration_plans: Record<string, Record<string, unknown>>;
  extension_install_plan: Record<string, unknown>;
  adaptation_strategy_matrix: Record<string, unknown>;
  runtime_event_plan_matrix: Record<string, unknown>;
  runtime_bundle_matrix: Record<string, unknown>;
  handoff_bundle: Record<string, unknown>;
  adapter_contract_matrix: Record<string, unknown>;
  adapter_contract_acceptance_matrix: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingPlatformHostIntegrationScaffoldFile {
  path: string;
  role: string;
  mime: string;
  content: string;
}

export interface MeetingPlatformHostIntegrationScaffold {
  type: 'meeting_platform_host_integration_scaffold';
  schema: string;
  schema_version: number;
  platforms: string[];
  plan: MeetingPlatformHostIntegrationPlan;
  files: MeetingPlatformHostIntegrationScaffoldFile[];
}

export interface MeetingPlatformHostIntegrationScaffoldAcceptanceReport {
  type: 'meeting_platform_host_integration_acceptance';
  schema: string;
  schema_version: number;
  accepted: boolean;
  platform_count: number;
  file_count: number;
  required_files: string[];
  blocking_count: number;
  warning_count: number;
  issues: Array<Record<string, unknown>>;
  scaffold: MeetingPlatformHostIntegrationScaffold;
}

export function buildMeetingPlatformHostIntegrationPlan(
  options?: MeetingPlatformHostIntegrationOptions,
): MeetingPlatformHostIntegrationPlan;

export function buildMeetingPlatformHostIntegrationScaffold(
  options?: MeetingPlatformHostIntegrationOptions,
): MeetingPlatformHostIntegrationScaffold;

export function buildMeetingPlatformHostIntegrationScaffoldAcceptanceReport(
  scaffoldOrOptions?: MeetingPlatformHostIntegrationScaffold | MeetingPlatformHostIntegrationOptions,
  options?: MeetingPlatformHostIntegrationOptions,
): MeetingPlatformHostIntegrationScaffoldAcceptanceReport;

export function assertMeetingPlatformHostIntegrationScaffold(
  scaffoldOrOptions?: MeetingPlatformHostIntegrationScaffold | MeetingPlatformHostIntegrationOptions,
  options?: MeetingPlatformHostIntegrationOptions,
): MeetingPlatformHostIntegrationScaffoldAcceptanceReport;
