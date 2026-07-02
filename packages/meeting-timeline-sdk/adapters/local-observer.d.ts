import type { MeetingTimelineClient } from '../index.mjs';
import type { ApplyMeetingSignalOptions, ApplyMeetingSignalResult, NormalizedMeetingSignal } from './core.mjs';
import type { DetectedMeetingFromUrl } from './meeting-url.mjs';

export interface LocalObserverState {
  activeMeeting: DetectedMeetingFromUrl | null;
  activeSinceMs: number | null;
  lastObservedAtMs: number | null;
}

export interface LocalObserverSnapshot {
  url?: string;
  meeting_url?: string;
  meetingUrl?: string;
  join_url?: string;
  joinUrl?: string;
  title?: string;
  topic?: string;
  name?: string;
  platform?: string;
  provider?: string;
  meeting_id?: string;
  meetingId?: string;
  active?: boolean;
  in_meeting?: boolean;
  inMeeting?: boolean;
  closed?: boolean;
  visible?: boolean;
  observed_at_ms?: number | string | Date;
  observedAtMs?: number | string | Date;
  timestamp_ms?: number | string | Date;
  timestampMs?: number | string | Date;
  timestamp?: number | string | Date;
  window?: Record<string, unknown>;
  browser?: Record<string, unknown>;
  tab?: Record<string, unknown>;
  meeting?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface LocalObserverOptions {
  source?: string;
  initialState?: LocalObserverState | null;
  observedAtMs?: number | string | Date;
  receivedAtMs?: number | string | Date;
  [key: string]: unknown;
}

export interface LocalTimelineObserverOptions extends LocalObserverOptions {
  applyOptions?: ApplyMeetingSignalOptions;
}

export interface LocalObserverResult {
  state: LocalObserverState;
  signals: NormalizedMeetingSignal[];
  detectedMeeting?: DetectedMeetingFromUrl | null;
}

export interface LocalTimelineObserverResult extends LocalObserverResult {
  results: ApplyMeetingSignalResult[];
}

export function observeMeetingSnapshot(
  state?: LocalObserverState | null,
  snapshot?: LocalObserverSnapshot,
  options?: LocalObserverOptions,
): LocalObserverResult;

export function createLocalMeetingObserver(options?: LocalObserverOptions): {
  observe(snapshot?: LocalObserverSnapshot, options?: LocalObserverOptions): LocalObserverResult;
  getState(): LocalObserverState;
  reset(nextState?: LocalObserverState | null): LocalObserverState;
};

export function createLocalMeetingTimelineObserver(
  client: MeetingTimelineClient,
  options?: LocalTimelineObserverOptions,
): {
  observe(snapshot?: LocalObserverSnapshot, options?: LocalTimelineObserverOptions): Promise<LocalTimelineObserverResult>;
  getState(): LocalObserverState;
  reset(nextState?: LocalObserverState | null): LocalObserverState;
};
