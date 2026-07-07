export type MeetingAppExtensionPlatform =
  | 'google_meet'
  | 'microsoft_teams'
  | 'zoom'
  | 'lark'
  | 'webex';

export type MeetingAppExtensionMessageType =
  | 'meeting_timeline.client_call'
  | 'meeting_timeline.extension_attached'
  | 'meeting_timeline.extension_status'
  | 'meeting_timeline.observe_candidates'
  | 'meeting_timeline.preflight_current_window'
  | 'meeting_timeline.preflight_candidates';

export type MeetingAppExtensionClientCallMethod =
  | 'startMeeting'
  | 'endMeeting'
  | 'insertMark'
  | 'insertMarks'
  | 'importTranscript';

export interface MeetingAppExtensionProfile {
  schema: 'meeting_app_extension_profile';
  version: number;
  platform: MeetingAppExtensionPlatform;
  display_name: string;
  matches: string[];
  host_permissions: string[];
}

export interface MeetingAppExtensionOptions {
  platform?: MeetingAppExtensionPlatform | string;
  platforms?: Array<MeetingAppExtensionPlatform | string> | MeetingAppExtensionPlatform | string;
  platform_keys?: Array<MeetingAppExtensionPlatform | string> | MeetingAppExtensionPlatform | string;
  platformKeys?: Array<MeetingAppExtensionPlatform | string> | MeetingAppExtensionPlatform | string;
  js?: string[] | string;
  contentScriptJs?: string[] | string;
  content_script_js?: string[] | string;
  runAt?: string;
  run_at?: string;
  allFrames?: boolean;
  all_frames?: boolean;
  matchAboutBlank?: boolean;
  match_about_blank?: boolean;
  extraMatches?: string[] | string;
  extra_matches?: string[] | string;
  extraHostPermissions?: string[] | string;
  extra_host_permissions?: string[] | string;
  permissions?: string[] | string;
  extension_permissions?: string[] | string;
  manifestVersion?: number;
  manifest_version?: number;
  name?: string;
  version?: string;
  description?: string;
  webAccessibleResources?: unknown[];
  web_accessible_resources?: unknown[];
  manifestOverrides?: Record<string, unknown>;
  manifest_overrides?: Record<string, unknown>;
  includeManifest?: boolean;
  include_manifest?: boolean;
  baseUrl?: string;
  base_url?: string;
  outputScript?: string;
  output_script?: string;
  backgroundScript?: string;
  background_script?: string;
  liveCaptureScript?: string;
  live_capture_script?: string;
  contentScriptEntry?: string;
  content_script_entry?: string;
  liveCaptureEntry?: string;
  live_capture_entry?: string;
  backgroundEntry?: string;
  background_entry?: string;
  includeLiveCapture?: boolean;
  include_live_capture?: boolean;
  liveCaptureGlobal?: string;
  live_capture_global?: string;
  buildTarget?: string[] | string;
  build_target?: string[] | string;
  messagePrefix?: string;
  message_prefix?: string;
  packageName?: string;
  package_name?: string;
  packageVersion?: string;
  package_version?: string;
  privatePackage?: boolean;
  private_package?: boolean;
  sdkDependencyVersion?: string;
  sdk_dependency_version?: string;
  esbuildVersion?: string;
  esbuild_version?: string;
  scripts?: Record<string, string>;
  dependencies?: Record<string, string>;
  devDependencies?: Record<string, string>;
  dev_dependencies?: Record<string, string>;
  [key: string]: unknown;
}

export interface MeetingAppExtensionMatchPatterns {
  type: 'meeting_app_extension_match_patterns';
  schema: 'meeting_app_extension_profile';
  version: number;
  platforms: MeetingAppExtensionPlatform[];
  profiles: MeetingAppExtensionProfile[];
  matches: string[];
  host_permissions: string[];
  content_scripts: Array<Record<string, unknown>>;
}

export interface MeetingAppExtensionInstallPlan extends Omit<MeetingAppExtensionMatchPatterns, 'type'> {
  type: 'meeting_app_extension_install_plan';
  manifest?: Record<string, unknown>;
  content_script_adapter: string;
  meeting_app_content_script_adapter?: string;
  platform_integration_runtime_adapter?: string;
  browser_runtime_adapter: string;
  snapshot_recorder_adapter: string;
  launch_gate_adapter: string;
  runtime_contract: Record<string, unknown>;
  next_steps: string[];
  constraints: string[];
}

export interface MeetingAppExtensionScaffoldFile {
  path: string;
  role: string;
  mime: string;
  content: string;
}

export interface MeetingAppExtensionScaffold {
  type: 'meeting_app_extension_scaffold';
  schema: 'meeting_app_extension_profile';
  version: number;
  install_plan: MeetingAppExtensionInstallPlan;
  manifest: Record<string, unknown>;
  files: MeetingAppExtensionScaffoldFile[];
  bundle: Record<string, unknown>;
  validation: Record<string, unknown>;
}

export interface MeetingAppExtensionScaffoldAcceptanceReport {
  type: 'meeting_app_extension_scaffold_acceptance_report';
  schema: 'meeting_app_extension_profile';
  version: number;
  accepted: boolean;
  platforms: MeetingAppExtensionPlatform[];
  accepted_platform_count: number;
  platform_count: number;
  coverage_by_platform: Partial<Record<MeetingAppExtensionPlatform, Record<string, unknown>>>;
  files: string[];
  manifest: Record<string, unknown>;
  bundle: Record<string, unknown>;
  issues: Array<Record<string, unknown>>;
}

export interface MeetingAppExtensionMessageOptions {
  type?: MeetingAppExtensionMessageType | string;
  messageType?: MeetingAppExtensionMessageType | string;
  message_type?: MeetingAppExtensionMessageType | string;
  platform?: MeetingAppExtensionPlatform | string;
  platform_key?: MeetingAppExtensionPlatform | string;
  platformKey?: MeetingAppExtensionPlatform | string;
  captured_at_ms?: number;
  capturedAtMs?: number;
  url?: string;
  href?: string;
  meeting_url?: string;
  meetingUrl?: string;
  request_id?: string;
  requestId?: string;
  [key: string]: unknown;
}

export interface MeetingAppExtensionClientCallMessageInput extends MeetingAppExtensionMessageOptions {
  method?: MeetingAppExtensionClientCallMethod | string;
  action?: MeetingAppExtensionClientCallMethod | string;
  input?: Record<string, unknown>;
}

export const MEETING_APP_EXTENSION_SCHEMA: 'meeting_app_extension_profile';
export const MEETING_APP_EXTENSION_PLATFORM_KEYS: readonly MeetingAppExtensionPlatform[];
export const MEETING_APP_EXTENSION_PROFILES: Readonly<Record<MeetingAppExtensionPlatform, Readonly<MeetingAppExtensionProfile>>>;
export const MEETING_APP_EXTENSION_MESSAGE_TYPES: Readonly<{
  client_call: 'meeting_timeline.client_call';
  extension_attached: 'meeting_timeline.extension_attached';
  extension_status: 'meeting_timeline.extension_status';
  observe_candidates: 'meeting_timeline.observe_candidates';
  preflight_current_window: 'meeting_timeline.preflight_current_window';
  preflight_candidates: 'meeting_timeline.preflight_candidates';
}>;
export const MEETING_APP_EXTENSION_STATUS_STORAGE_KEY: 'meeting_timeline_extension_status';
export const MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS: Readonly<Record<MeetingAppExtensionClientCallMethod, string>>;

export function normalizeMeetingAppExtensionPlatform(platform: MeetingAppExtensionPlatform | string): MeetingAppExtensionPlatform;

export function normalizeMeetingAppExtensionMessageType(messageType: MeetingAppExtensionMessageType | string): MeetingAppExtensionMessageType;

export function meetingAppExtensionTimelineEndpoint(method: MeetingAppExtensionClientCallMethod | string): string;

export function buildMeetingAppExtensionAttachedMessage(
  input?: MeetingAppExtensionMessageOptions | MeetingAppExtensionPlatform | string,
  options?: MeetingAppExtensionMessageOptions,
): Record<string, unknown>;

export function buildMeetingAppExtensionStatusMessage(
  input?: MeetingAppExtensionMessageOptions,
  options?: MeetingAppExtensionMessageOptions,
): Record<string, unknown>;

export function buildMeetingAppExtensionObserveCandidatesMessage(
  input?: MeetingAppExtensionMessageOptions,
  options?: MeetingAppExtensionMessageOptions,
): Record<string, unknown>;

export function buildMeetingAppExtensionCurrentWindowPreflightMessage(
  input?: MeetingAppExtensionMessageOptions & {
    input?: Record<string, unknown>;
    options?: Record<string, unknown>;
    preflightOptions?: Record<string, unknown>;
    preflight_options?: Record<string, unknown>;
  },
  options?: MeetingAppExtensionMessageOptions,
): Record<string, unknown>;

export function buildMeetingAppExtensionPreflightCandidatesMessage(
  input?: MeetingAppExtensionMessageOptions & {
    query?: Record<string, unknown>;
    tabs?: Record<string, unknown>[];
    browser_tabs?: Record<string, unknown>[];
    browserTabs?: Record<string, unknown>[];
    input?: Record<string, unknown>;
    options?: Record<string, unknown>;
    preflightOptions?: Record<string, unknown>;
    preflight_options?: Record<string, unknown>;
  },
  options?: MeetingAppExtensionMessageOptions,
): Record<string, unknown>;

export function buildMeetingAppExtensionClientCallMessage(
  methodOrInput: MeetingAppExtensionClientCallMethod | string | MeetingAppExtensionClientCallMessageInput,
  input?: Record<string, unknown>,
  options?: MeetingAppExtensionMessageOptions,
): Record<string, unknown>;

export function meetingAppExtensionProfile(
  platformOrInput: MeetingAppExtensionPlatform | string | MeetingAppExtensionOptions,
  options?: MeetingAppExtensionOptions,
): MeetingAppExtensionProfile;

export function buildMeetingAppExtensionMatchPatterns(
  platformsOrOptions?: MeetingAppExtensionOptions | Array<MeetingAppExtensionPlatform | string> | MeetingAppExtensionPlatform | string,
  options?: MeetingAppExtensionOptions,
): MeetingAppExtensionMatchPatterns;

export function buildMeetingAppContentScriptManifest(options?: MeetingAppExtensionOptions): Record<string, unknown>;

export function buildMeetingAppExtensionInstallPlan(options?: MeetingAppExtensionOptions): MeetingAppExtensionInstallPlan;

export function buildMeetingAppExtensionContentScriptSource(options?: MeetingAppExtensionOptions): string;

export function buildMeetingAppExtensionLiveCaptureSource(options?: MeetingAppExtensionOptions): string;

export function buildMeetingAppExtensionBackgroundSource(options?: MeetingAppExtensionOptions): string;

export function buildMeetingAppExtensionPackageJson(options?: MeetingAppExtensionOptions): Record<string, unknown>;

export function buildMeetingAppExtensionBuildSource(options?: MeetingAppExtensionOptions): string;

export function buildMeetingAppExtensionReadme(options?: MeetingAppExtensionOptions): string;

export function buildMeetingAppExtensionScaffold(options?: MeetingAppExtensionOptions): MeetingAppExtensionScaffold;

export function buildMeetingAppExtensionScaffoldAcceptanceReport(
  scaffoldOrOptions?: MeetingAppExtensionScaffold | MeetingAppExtensionOptions,
  options?: MeetingAppExtensionOptions,
): MeetingAppExtensionScaffoldAcceptanceReport;

export function assertMeetingAppExtensionScaffold(
  scaffoldOrOptions?: MeetingAppExtensionScaffold | MeetingAppExtensionOptions,
  options?: MeetingAppExtensionOptions,
): MeetingAppExtensionScaffoldAcceptanceReport;

export default buildMeetingAppExtensionInstallPlan;
