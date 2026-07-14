import type { MeetingAppFixtureOptions, MeetingAppFixturePlatform } from './meeting-app-fixtures.mjs';
import type { MeetingAppSnapshot } from './meeting-apps.mjs';

export const MEETING_APP_FIXTURE_TRACK_READINESS_SCHEMA: 'meeting_app_fixture_track_readiness';
export const MEETING_APP_FIXTURE_TRACK_READINESS_REPORT_SCHEMA: 'meeting_app_fixture_track_readiness_report';
export const MEETING_APP_FIXTURE_TRACK_READINESS_SCHEMA_VERSION: number;

export interface MeetingAppFixtureTrackReadinessCoverage {
  participant_roster_snapshot: boolean;
  participant_roster_count: number;
  speaker_track_mark: boolean;
  participant_track_mark: boolean;
  speaker_track_payload_text_free: boolean;
  participant_track_payload_text_free: boolean;
  [key: string]: boolean | number;
}

export interface MeetingAppFixtureTrackReadiness {
  type: 'meeting_app_fixture_track_readiness';
  schema: typeof MEETING_APP_FIXTURE_TRACK_READINESS_SCHEMA;
  schema_version: number;
  platform: MeetingAppFixturePlatform;
  ready: boolean;
  required: string[];
  missing: string[];
  snapshots: {
    active: MeetingAppSnapshot;
    participant_changed: MeetingAppSnapshot;
  };
  normalized: MeetingAppSnapshot | null;
  speaker_track: Record<string, unknown>;
  participant_track: Record<string, unknown>;
  coverage: MeetingAppFixtureTrackReadinessCoverage;
}

export interface MeetingAppFixtureTrackReadinessReport {
  type: 'meeting_app_fixture_track_readiness_report';
  schema: typeof MEETING_APP_FIXTURE_TRACK_READINESS_REPORT_SCHEMA;
  schema_version: number;
  accepted: boolean;
  required: string[];
  missing: string[];
  platform_count: number;
  accepted_count: number;
  coverage_by_platform: Record<MeetingAppFixturePlatform, MeetingAppFixtureTrackReadinessCoverage>;
  rows: MeetingAppFixtureTrackReadiness[];
}

export function diagnoseMeetingAppFixtureTrackReadiness(
  platform: string,
  options?: MeetingAppFixtureOptions,
): MeetingAppFixtureTrackReadiness;

export function buildMeetingAppFixtureTrackReadinessReport(
  options?: MeetingAppFixtureOptions,
): MeetingAppFixtureTrackReadinessReport;
