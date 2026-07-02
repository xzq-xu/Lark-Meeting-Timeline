import type { MeetingTimelineClient } from '../index.mjs';
import type {
  BrowserMeetingObserverOptions,
  BrowserMeetingObserverResult,
  BrowserMeetingObserverState,
  BrowserMeetingTimelineObserverResult,
} from './browser-meeting.mjs';

export type MeetingAppPlatform = 'google_meet' | 'microsoft_teams' | 'zoom' | 'lark' | 'webex';

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
