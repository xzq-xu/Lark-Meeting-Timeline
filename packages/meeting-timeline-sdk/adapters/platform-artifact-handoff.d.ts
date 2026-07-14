import type { ArtifactImportPlan } from './artifact-plan.mjs';
import type { MeetingPlatformProviderConnectionOptions } from './platform-provider-connection.mjs';

export const MEETING_PLATFORM_ARTIFACT_HANDOFF_PLAN_SCHEMA: 'meeting_platform_artifact_handoff_plan';
export const MEETING_PLATFORM_ARTIFACT_HANDOFF_MATRIX_SCHEMA: 'meeting_platform_artifact_handoff_matrix';
export const MEETING_PLATFORM_ARTIFACT_HANDOFF_SCHEMA: 'meeting_platform_artifact_handoff';
export const MEETING_PLATFORM_ARTIFACT_HANDOFF_SCHEMA_VERSION: number;

export interface MeetingPlatformArtifactHandoffOptions extends MeetingPlatformProviderConnectionOptions {
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  artifactKinds?: Iterable<string> | string[];
  artifact_kinds?: Iterable<string> | string[];
  importEndpoint?: string;
  import_endpoint?: string;
  includeIgnored?: boolean;
  include_ignored?: boolean;
  includeForeignPlatform?: boolean;
  include_foreign_platform?: boolean;
  fetchOptions?: Record<string, unknown>;
  fetch_options?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingPlatformArtifactHandoffInput {
  signals?: Array<Record<string, unknown>>;
  artifactSignals?: Array<Record<string, unknown>>;
  artifact_signals?: Array<Record<string, unknown>>;
  events?: Array<Record<string, unknown>>;
  rawSignals?: Array<Record<string, unknown>>;
  raw_signals?: Array<Record<string, unknown>>;
  [key: string]: unknown;
}

export interface MeetingPlatformArtifactSupportRow {
  artifact_kind: string;
  supported: boolean;
  action: string;
  event_types: string[];
  fetch_strategy?: string;
  normalizer?: string | null;
  content_hint?: string;
  token_env: string;
  post_meeting_only: boolean;
  realtime_blocking: boolean;
}

export interface MeetingPlatformArtifactHandoffPlan {
  type: 'meeting_platform_artifact_handoff_plan';
  schema: 'meeting_platform_artifact_handoff_plan';
  schema_version: number;
  platform: string;
  display_name?: string;
  status: string;
  post_meeting_only: boolean;
  provider_events_block_realtime: boolean;
  transcript_blocks_realtime: boolean;
  import_endpoint: string;
  artifacts: MeetingPlatformArtifactSupportRow[];
  supported_artifact_kinds: string[];
  transcript_supported: boolean;
  recording_supported: boolean;
  smart_notes_supported: boolean;
  source_event_types: string[];
  sdk_modules: Record<string, string>;
  handoff_contract: Record<string, unknown>;
  official_docs: Array<Record<string, unknown>>;
  next_actions: string[];
}

export interface MeetingPlatformArtifactHandoffMatrix {
  type: 'meeting_platform_artifact_handoff_matrix';
  schema: 'meeting_platform_artifact_handoff_matrix';
  schema_version: number;
  platform_count: number;
  transcript_supported_count: number;
  recording_supported_count: number;
  smart_notes_supported_count: number;
  realtime_blocking_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  plans: MeetingPlatformArtifactHandoffPlan[];
  next_actions: string[];
}

export interface MeetingPlatformArtifactHandoff {
  type: 'meeting_platform_artifact_handoff';
  schema: 'meeting_platform_artifact_handoff';
  schema_version: number;
  platform: string;
  status: string;
  post_meeting_only: boolean;
  provider_events_block_realtime: boolean;
  transcript_blocks_realtime: boolean;
  input_signal_count: number;
  import_plan_count: number;
  fetch_request_count: number;
  transcript_import_count: number;
  recording_count: number;
  warning_count: number;
  plan: MeetingPlatformArtifactHandoffPlan;
  import_plans: ArtifactImportPlan[];
  rows: Array<Record<string, unknown>>;
  fetch_requests: Array<Record<string, unknown>>;
  next_actions: string[];
}

export function buildMeetingPlatformArtifactHandoffPlan(
  platform: string,
  options?: MeetingPlatformArtifactHandoffOptions,
): MeetingPlatformArtifactHandoffPlan;

export function buildMeetingPlatformArtifactHandoffMatrix(
  options?: MeetingPlatformArtifactHandoffOptions,
): MeetingPlatformArtifactHandoffMatrix;

export function buildMeetingPlatformArtifactHandoff(
  platform: string,
  input?: MeetingPlatformArtifactHandoffInput | Array<Record<string, unknown>>,
  options?: MeetingPlatformArtifactHandoffOptions,
): MeetingPlatformArtifactHandoff;
