export type MeetingAppExtensionPlatform =
  | 'google_meet'
  | 'microsoft_teams'
  | 'zoom'
  | 'lark'
  | 'webex';

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
  messagePrefix?: string;
  message_prefix?: string;
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

export const MEETING_APP_EXTENSION_SCHEMA: 'meeting_app_extension_profile';
export const MEETING_APP_EXTENSION_PLATFORM_KEYS: readonly MeetingAppExtensionPlatform[];
export const MEETING_APP_EXTENSION_PROFILES: Readonly<Record<MeetingAppExtensionPlatform, Readonly<MeetingAppExtensionProfile>>>;

export function normalizeMeetingAppExtensionPlatform(platform: MeetingAppExtensionPlatform | string): MeetingAppExtensionPlatform;

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

export function buildMeetingAppExtensionBackgroundSource(options?: MeetingAppExtensionOptions): string;

export function buildMeetingAppExtensionReadme(options?: MeetingAppExtensionOptions): string;

export function buildMeetingAppExtensionScaffold(options?: MeetingAppExtensionOptions): MeetingAppExtensionScaffold;

export default buildMeetingAppExtensionInstallPlan;
