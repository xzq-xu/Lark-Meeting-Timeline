import type { MeetingAppTimelineConnectorPackage } from '../index.mjs';

export const MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA: 'meeting_app_timeline_connector_package';
export const MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_ACCEPTANCE_SCHEMA: 'meeting_app_timeline_connector_package_acceptance_report';
export const MEETING_APP_TIMELINE_CONNECTOR_HANDOFF_SCHEMA: 'meeting_app_timeline_connector_handoff';
export const MEETING_APP_TIMELINE_CONNECTOR_HOST_INSTALL_CHECKLIST_SCHEMA: 'meeting_app_timeline_connector_host_install_checklist';
export const MEETING_APP_TIMELINE_CONNECTOR_HOST_INSTALL_CHECKLIST_ACCEPTANCE_SCHEMA: 'meeting_app_timeline_connector_host_install_checklist_acceptance_report';
export const MEETING_APP_TIMELINE_CONNECTOR_RUNTIME_CLIENT_SCHEMA: 'meeting_app_timeline_connector_runtime_client';
export const MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA_VERSION: 1;

export interface MeetingAppTimelineConnectorPackageIssue {
  code: string;
  message: string;
  [key: string]: unknown;
}

export interface MeetingAppTimelineConnectorPackageSurfaceReport {
  surface: string;
  accepted: boolean;
  handoff_count: number;
  ready_count: number;
  realtime_annotation_ready_count: number;
  speaker_track_ready_count: number;
  participant_track_ready_count: number;
  observer_accepted_count: number;
  scheduler_sdk_ready_count: number;
  track_enabled_count: number;
  install_targets: string[];
  start_modes: string[];
  runtime_factories: string[];
  observer_factories: string[];
  issues: MeetingAppTimelineConnectorPackageIssue[];
}

export interface MeetingAppTimelineConnectorPackageAcceptanceReport {
  type: 'meeting_app_timeline_connector_package_acceptance_report';
  schema: 'meeting_app_timeline_connector_package_acceptance_report';
  schema_version: 1;
  accepted: boolean;
  target: string;
  package_id?: string;
  platform_count: number;
  surface_count: number;
  platforms: string[];
  surfaces: string[];
  require_tracks: boolean;
  missing_runtime_actions: Array<{ platform: string; action: string }>;
  surface_reports: MeetingAppTimelineConnectorPackageSurfaceReport[];
  extension: {
    required: boolean;
    accepted: boolean;
    file_count?: number;
    manifest_version?: number;
    issues: MeetingAppTimelineConnectorPackageIssue[];
  };
  issue_count: number;
  issues: MeetingAppTimelineConnectorPackageIssue[];
  next_actions: string[];
}

export interface MeetingAppTimelineConnectorHandoff {
  type: 'meeting_app_timeline_connector_handoff';
  schema: 'meeting_app_timeline_connector_handoff';
  schema_version: 1;
  package_id?: string;
  accepted: boolean;
  target: string;
  base_url?: string;
  platforms: string[];
  surfaces: string[];
  platform_count: number;
  surface_count: number;
  runtime_event_endpoint?: string;
  adapter_blueprints?: {
    ready_count?: number;
    platform_count?: number;
    sdk_method?: string;
    command?: string;
    rows?: Array<Record<string, unknown>>;
  };
  startup_plans?: {
    realtime_startup_ready_count?: number;
    platform_count?: number;
    sdk_method?: string;
    matrix_sdk_method?: string;
    rows?: Array<Record<string, unknown>>;
  };
  timestamp_field?: string;
  provider_events_block_realtime?: boolean;
  transcript_blocks_realtime?: boolean;
  host_entrypoints?: Array<Record<string, unknown>>;
  package_entrypoints?: Array<Record<string, unknown>>;
  ci_gates?: string[];
  rollout_checklist?: string[];
  surface_matrix: Array<Record<string, unknown>>;
  extension?: Record<string, unknown>;
  acceptance: MeetingAppTimelineConnectorPackageAcceptanceReport;
  next_actions: string[];
}

export interface MeetingAppTimelineConnectorHostInstallChecklist {
  type: 'meeting_app_timeline_connector_host_install_checklist';
  schema: 'meeting_app_timeline_connector_host_install_checklist';
  schema_version: 1;
  package_id?: string;
  accepted: boolean;
  target: string;
  base_url?: string;
  runtime_event_endpoint?: string;
  platform_count: number;
  surface_count: number;
  ready_count: number;
  timestamp_field?: string;
  contracts: Record<string, unknown>;
  files_to_read_first: string[];
  rows: Array<Record<string, unknown>>;
  acceptance: MeetingAppTimelineConnectorPackageAcceptanceReport;
  next_actions: string[];
}

export interface MeetingAppTimelineConnectorHostInstallChecklistAcceptanceReport {
  type: 'meeting_app_timeline_connector_host_install_checklist_acceptance_report';
  schema: 'meeting_app_timeline_connector_host_install_checklist_acceptance_report';
  schema_version: 1;
  accepted: boolean;
  target?: string;
  package_id?: string;
  platform_count: number;
  row_count: number;
  ready_count: number;
  timestamp_field?: string;
  runtime_event_endpoint?: string;
  checklist_accepted: boolean;
  issue_count: number;
  issues: MeetingAppTimelineConnectorPackageIssue[];
  rows: Array<Record<string, unknown>>;
  next_actions: string[];
}

export interface MeetingAppTimelineConnectorRuntimeClient {
  type: 'meeting_app_timeline_connector_runtime_client';
  schema: 'meeting_app_timeline_connector_runtime_client';
  schema_version: 1;
  package_id?: string;
  endpoint: string;
  platforms: string[];
  surfaces: string[];
  acceptance: MeetingAppTimelineConnectorPackageAcceptanceReport;
  runtime_event_client: import('./platform-runtime-event.mjs').MeetingPlatformRuntimeEventClient;
  supported_action_count: number;
  supported_actions: string[];
  supported_actions_by_platform: Record<string, string[]>;
  action_rows: Array<Record<string, unknown>>;
  supports(action: string, platform?: string): boolean;
  assertSupported(action: string, platform?: string, options?: Record<string, unknown>): Record<string, unknown>;
  buildEvent(input?: Record<string, unknown>, options?: Record<string, unknown>): import('./platform-runtime-event.mjs').MeetingPlatformRuntimeEvent;
  send(input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  observeMeetingApp(platform: string, snapshot?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  observePlatformCandidates(input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  ingestProvider(platform: string, payload?: unknown, options?: Record<string, unknown>): Promise<unknown>;
  insertAnnotation(platform: string, annotationInput?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  insertMark(platform: string, annotationInput?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  speakerTrack(platform: string, input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  participantTrack(platform: string, input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  timelineView(platform: string, input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  adapterRoute(platform: string, options?: Record<string, unknown>): Promise<unknown>;
  adapterRoutes(options?: Record<string, unknown>): Promise<unknown>;
  adapterBlueprint(platform: string, options?: Record<string, unknown>): Promise<unknown>;
  adapterBlueprints(options?: Record<string, unknown>): Promise<unknown>;
  runManifest(options?: Record<string, unknown>): Promise<unknown>;
  runHandoffReadiness(options?: Record<string, unknown>): Promise<unknown>;
}

export function buildMeetingAppTimelineConnectorPackageAcceptanceReport(
  pkg?: MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorPackageAcceptanceReport;

export function assertMeetingAppTimelineConnectorPackage<T extends MeetingAppTimelineConnectorPackage | Record<string, unknown>>(
  pkg?: T,
  options?: Record<string, unknown>,
): T;

export function stripMeetingAppTimelineConnectorPackageFileContents<T extends MeetingAppTimelineConnectorPackage | Record<string, unknown>>(
  pkg?: T,
): T;

export function buildMeetingAppTimelineConnectorHandoff(
  pkg?: MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorHandoff;

export function buildMeetingAppTimelineConnectorHostInstallChecklist(
  pkg?: MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorHostInstallChecklist;

export function buildMeetingAppTimelineConnectorHostInstallChecklistAcceptanceReport(
  checklistOrPackage?: MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorHostInstallChecklistAcceptanceReport;

export function assertMeetingAppTimelineConnectorHostInstallChecklist<T extends MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>>(
  checklistOrPackage?: T,
  options?: Record<string, unknown>,
): T;

export function createMeetingAppTimelineConnectorRuntimeClient(
  pkg?: MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorRuntimeClient;
