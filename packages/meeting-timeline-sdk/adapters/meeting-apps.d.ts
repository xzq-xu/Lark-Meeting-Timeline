import type { MeetingTimelineClient } from '../index.mjs';
import type {
  BrowserMeetingObserverOptions,
  BrowserMeetingObserverResult,
  BrowserMeetingObserverState,
  BrowserMeetingTimelineObserverResult,
} from './browser-meeting.mjs';

export type MeetingAppPlatform = 'google_meet' | 'microsoft_teams' | 'zoom' | 'lark' | 'webex';

export const MEETING_APP_ADAPTER_FIT_SCHEMA: 'meeting_app_adapter_fit_report';
export const MEETING_APP_ADAPTER_FIT_MATRIX_SCHEMA: 'meeting_app_adapter_fit_matrix';
export const MEETING_APP_ADAPTER_FIT_SCHEMA_VERSION: 1;

export interface MeetingAppPreset {
  platform: MeetingAppPlatform;
  displayName?: string;
  joinedHints?: (RegExp | string)[];
  preJoinHints?: (RegExp | string)[];
  participantPaths?: string[];
  activeSpeakerPaths?: string[];
  textPaths?: string[];
  textCollectionPaths?: string[];
  participantIdPaths?: string[];
  participantNamePaths?: string[];
  speakingPaths?: string[];
  audioLevelPaths?: string[];
  speakingHints?: (RegExp | string)[];
  mutedHints?: (RegExp | string)[];
  audioLevelThreshold?: number;
  [key: string]: unknown;
}

export interface MeetingAppSnapshot {
  platform?: string;
  provider?: string;
  meeting?: Record<string, unknown>;
  url?: string;
  meeting_url?: string;
  meetingUrl?: string;
  meeting_id?: string;
  meetingId?: string;
  external_meeting_id?: string;
  externalMeetingId?: string;
  href?: string;
  title?: string;
  topic?: string;
  name?: string;
  browser?: Record<string, unknown>;
  application?: Record<string, unknown>;
  app?: Record<string, unknown>;
  process?: Record<string, unknown>;
  window?: Record<string, unknown>;
  tab?: Record<string, unknown>;
  page?: Record<string, unknown>;
  dom?: Record<string, unknown>;
  accessibility?: Record<string, unknown>;
  ax?: Record<string, unknown>;
  activeSpeaker?: Record<string, unknown>;
  active_speaker?: Record<string, unknown>;
  speaker?: Record<string, unknown>;
  participants?: Record<string, unknown>[];
  tiles?: Record<string, unknown>[];
  videoTiles?: Record<string, unknown>[];
  buttons?: unknown[];
  controls?: unknown[];
  labels?: unknown[];
  texts?: unknown[];
  ariaLabels?: unknown[];
  in_meeting?: boolean;
  inMeeting?: boolean;
  active?: boolean;
  focused?: boolean;
  selected?: boolean;
  visible?: boolean;
  audible?: boolean;
  observed_at_ms?: number | string | Date;
  observedAtMs?: number | string | Date;
  timestamp_ms?: number | string | Date;
  timestampMs?: number | string | Date;
  timestamp?: number | string | Date;
  windows?: MeetingAppSnapshot[];
  tabs?: MeetingAppSnapshot[];
  pages?: MeetingAppSnapshot[];
  frames?: MeetingAppSnapshot[];
  applications?: MeetingAppSnapshot[];
  apps?: MeetingAppSnapshot[];
  processes?: MeetingAppSnapshot[];
  candidates?: MeetingAppSnapshot[];
  snapshots?: MeetingAppSnapshot[];
  items?: MeetingAppSnapshot[];
  [key: string]: unknown;
}

export interface MeetingAppObserverOptions extends BrowserMeetingObserverOptions {
  platform?: string;
  preset?: MeetingAppPreset;
}

export interface MeetingAppPresetDetection extends MeetingAppPreset {
  confidence?: string;
  reason?: string;
  detected?: Record<string, unknown>;
}

export interface MeetingAppAdapterFitRow {
  index: number;
  platform?: string;
  meeting_id?: string;
  meeting_url?: string;
  title?: string;
  in_meeting?: boolean;
  active?: boolean;
  visible?: boolean;
  audible?: boolean;
  participant_count?: number;
  active_speaker_id?: string;
  active_speaker_name?: string;
  has_active_speaker?: boolean;
}

export interface MeetingAppAdapterFitIssue {
  severity: 'error' | 'warning' | string;
  code: string;
  message: string;
  [key: string]: unknown;
}

export interface MeetingAppAdapterFitReport {
  type: 'meeting_app_adapter_fit_report';
  schema: 'meeting_app_adapter_fit_report';
  version: 1;
  platform?: string;
  expected_platform?: string;
  detected_platforms: string[];
  accepted: boolean;
  ready_for_realtime_axis: boolean;
  ready_for_speaker_track: boolean;
  ready_for_participant_track: boolean;
  recommended_surface: string;
  input_shapes: string[];
  candidate_count: number;
  meeting_identity_count: number;
  start_candidate_count: number;
  end_candidate_count: number;
  active_speaker_count: number;
  participant_count: number;
  coverage: Record<string, boolean>;
  rows: MeetingAppAdapterFitRow[];
  issues: MeetingAppAdapterFitIssue[];
  next_actions: string[];
}

export interface MeetingAppAdapterFitMatrix {
  type: 'meeting_app_adapter_fit_matrix';
  schema: 'meeting_app_adapter_fit_matrix';
  version: 1;
  platform_count: number;
  accepted_count: number;
  realtime_axis_ready_count: number;
  speaker_track_ready_count: number;
  participant_track_ready_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  reports: MeetingAppAdapterFitReport[];
  next_actions: string[];
}

export const MEETING_APP_PRESETS: Readonly<Record<MeetingAppPlatform, MeetingAppPreset>>;

export function detectMeetingAppPreset(
  input?: MeetingAppSnapshot,
  options?: MeetingAppObserverOptions,
): MeetingAppPresetDetection | null;

export function normalizeMeetingAppSnapshot(
  input?: MeetingAppSnapshot,
  options?: MeetingAppObserverOptions,
): MeetingAppSnapshot | null;

export function normalizeMeetingAppSnapshots(
  input?: MeetingAppSnapshot | MeetingAppSnapshot[],
  options?: MeetingAppObserverOptions,
): MeetingAppSnapshot[];

export function buildMeetingAppAdapterFitReport(
  input?: MeetingAppSnapshot | MeetingAppSnapshot[] | Record<string, unknown>,
  options?: MeetingAppObserverOptions,
): MeetingAppAdapterFitReport;

export function buildMeetingAppAdapterFitMatrix(
  input?: MeetingAppSnapshot | MeetingAppSnapshot[] | Record<string, unknown>,
  options?: MeetingAppObserverOptions,
): MeetingAppAdapterFitMatrix;

export function observeMeetingAppSample(
  state?: BrowserMeetingObserverState | null,
  input?: MeetingAppSnapshot,
  options?: MeetingAppObserverOptions,
): BrowserMeetingObserverResult;

export function createMeetingAppObserver(options?: MeetingAppObserverOptions): {
  observe(input?: MeetingAppSnapshot, options?: MeetingAppObserverOptions): BrowserMeetingObserverResult;
  getState(): BrowserMeetingObserverState;
  reset(nextState?: BrowserMeetingObserverState | null): BrowserMeetingObserverState;
};

export function createMeetingAppTimelineObserver(
  client: MeetingTimelineClient,
  options?: MeetingAppObserverOptions,
): {
  observe(input?: MeetingAppSnapshot, options?: MeetingAppObserverOptions): Promise<BrowserMeetingTimelineObserverResult>;
  getState(): BrowserMeetingObserverState;
  reset(nextState?: BrowserMeetingObserverState | null): BrowserMeetingObserverState;
};
