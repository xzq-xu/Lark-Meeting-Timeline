import type { MeetingAppObserverOptions, MeetingAppSnapshot } from './meeting-apps.mjs';
import type { BrowserMeetingObserverResult } from './browser-meeting.mjs';

export type MeetingAppFixturePlatform = 'google_meet' | 'microsoft_teams' | 'zoom' | 'lark' | 'webex';

export interface MeetingAppFixtureOptions extends MeetingAppObserverOptions {
  observedAtMs?: number | string | Date;
  observed_at_ms?: number | string | Date;
  endObservedAtMs?: number | string | Date;
  end_observed_at_ms?: number | string | Date;
  endOffsetMs?: number;
  end_offset_ms?: number;
  startMs?: number | string | Date;
  start_ms?: number | string | Date;
  state?: 'active' | 'joined' | 'in_meeting' | 'in-meeting' | 'prejoin' | 'pre_join' | 'pre-join' | 'ended' | 'left' | 'inactive' | string;
  fixtureState?: string;
  fixture_state?: string;
  source?: string;
  url?: string;
  meeting_url?: string;
  title?: string;
  speakerId?: string;
  speaker_id?: string;
  speakerName?: string;
  speaker_name?: string;
  mutedId?: string;
  muted_id?: string;
  mutedName?: string;
  muted_name?: string;
  browserName?: string;
  browser_name?: string;
  documentVisible?: boolean;
  document_visible?: boolean;
  platforms?: MeetingAppFixturePlatform[];
  platform_keys?: MeetingAppFixturePlatform[];
  [key: string]: unknown;
}

export interface MeetingAppFixtureDiagnosis {
  platform: MeetingAppFixturePlatform;
  snapshot: MeetingAppSnapshot;
  normalized: MeetingAppSnapshot | null;
  signal_types: string[];
  signals: Record<string, unknown>[];
  coverage: {
    platform_detected: boolean;
    meeting_id: boolean;
    in_meeting: boolean;
    active_speaker: boolean;
    meeting_started: boolean;
    speaker_started: boolean;
    [key: string]: boolean;
  };
  observation: BrowserMeetingObserverResult;
}

export interface MeetingAppFixtureLifecycleDiagnosis {
  platform: MeetingAppFixturePlatform;
  snapshots: {
    active: MeetingAppSnapshot;
    ended: MeetingAppSnapshot;
  };
  normalized_ended: MeetingAppSnapshot | null;
  signal_types: string[];
  signals: Record<string, unknown>[];
  coverage: {
    meeting_started: boolean;
    speaker_started: boolean;
    meeting_ended: boolean;
    ended_in_meeting_false: boolean;
    [key: string]: boolean;
  };
  observations: {
    active: BrowserMeetingObserverResult;
    ended: BrowserMeetingObserverResult;
  };
}

export const MEETING_APP_FIXTURE_PLATFORMS: readonly MeetingAppFixturePlatform[];

export function buildMeetingAppFixtureSnapshot(
  platform: string,
  options?: MeetingAppFixtureOptions,
): MeetingAppSnapshot;

export function buildAllMeetingAppFixtureSnapshots(
  options?: MeetingAppFixtureOptions,
): Record<MeetingAppFixturePlatform, MeetingAppSnapshot>;

export function diagnoseMeetingAppFixture(
  platform: string,
  options?: MeetingAppFixtureOptions,
): MeetingAppFixtureDiagnosis;

export function diagnoseMeetingAppFixtureLifecycle(
  platform: string,
  options?: MeetingAppFixtureOptions,
): MeetingAppFixtureLifecycleDiagnosis;

export function buildMeetingAppFixtureAcceptanceReport(
  options?: MeetingAppFixtureOptions,
): {
  type: string;
  accepted: boolean;
  required: string[];
  missing: string[];
  platform_count: number;
  accepted_count: number;
  coverage_by_platform: Record<MeetingAppFixturePlatform, MeetingAppFixtureDiagnosis['coverage']>;
  reports: MeetingAppFixtureDiagnosis[];
  lifecycle_reports: MeetingAppFixtureLifecycleDiagnosis[];
};
