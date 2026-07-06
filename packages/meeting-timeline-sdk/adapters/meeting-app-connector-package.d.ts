import type { MeetingAppTimelineConnectorPackage } from '../index.mjs';

export const MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA: 'meeting_app_timeline_connector_package';
export const MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_ACCEPTANCE_SCHEMA: 'meeting_app_timeline_connector_package_acceptance_report';
export const MEETING_APP_TIMELINE_CONNECTOR_HANDOFF_SCHEMA: 'meeting_app_timeline_connector_handoff';
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
