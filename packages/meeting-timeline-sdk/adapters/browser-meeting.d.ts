import type { MeetingTimelineClient } from '../index.mjs';
import type { ActiveSpeakerObserverResult, ActiveSpeakerTimelineObserverResult, NormalizedActiveSpeakerSample } from './active-speaker.mjs';
import type {
  MeetingSessionDiscoveryOptions,
  NormalizedMeetingSessionCandidate,
} from './meeting-session-discovery.mjs';
import type {
  LocalObserverCandidateResult,
  LocalObserverState,
  LocalTimelineObserverCandidateResult,
} from './local-observer.mjs';
import type { ApplyMeetingSignalOptions, ApplyMeetingSignalResult, NormalizedMeetingSignal } from './core.mjs';

export interface BrowserMeetingSnapshot {
  browser?: Record<string, unknown>;
  window?: Record<string, unknown>;
  tab?: Record<string, unknown>;
  page?: Record<string, unknown>;
  dom?: Record<string, unknown>;
  frame?: Record<string, unknown>;
  meeting?: Record<string, unknown>;
  activeSpeaker?: Record<string, unknown>;
  active_speaker?: Record<string, unknown>;
  speaker?: Record<string, unknown>;
  participants?: Record<string, unknown>[];
  url?: string;
  href?: string;
  title?: string;
  platform?: string;
  provider?: string;
  meeting_id?: string;
  meetingId?: string;
  active?: boolean;
  focused?: boolean;
  visible?: boolean;
  audible?: boolean;
  in_meeting?: boolean;
  inMeeting?: boolean;
  observed_at_ms?: number | string | Date;
  observedAtMs?: number | string | Date;
  timestamp_ms?: number | string | Date;
  timestampMs?: number | string | Date;
  timestamp?: number | string | Date;
  windows?: BrowserMeetingSnapshot[];
  tabs?: BrowserMeetingSnapshot[];
  pages?: BrowserMeetingSnapshot[];
  frames?: BrowserMeetingSnapshot[];
  candidates?: BrowserMeetingSnapshot[];
  [key: string]: unknown;
}

export interface BrowserMeetingObserverState {
  sessionState: LocalObserverState;
  speakerState: Record<string, unknown>;
}

export interface BrowserMeetingObserverOptions extends MeetingSessionDiscoveryOptions {
  source?: string;
  sessionSource?: string;
  speakerSource?: string;
  sessionOptions?: Record<string, unknown>;
  speakerOptions?: Record<string, unknown>;
  sessionObserveOptions?: Record<string, unknown>;
  speakerObserveOptions?: Record<string, unknown>;
  sessionApplyOptions?: ApplyMeetingSignalOptions;
  speakerApplyOptions?: ApplyMeetingSignalOptions;
  applyOptions?: ApplyMeetingSignalOptions;
}

export interface BrowserMeetingObserverResult {
  state: BrowserMeetingObserverState;
  signals: NormalizedMeetingSignal[];
  session: LocalObserverCandidateResult;
  speaker: ActiveSpeakerObserverResult;
  normalizedCandidates: NormalizedMeetingSessionCandidate[];
}

export interface BrowserMeetingTimelineObserverResult extends Omit<BrowserMeetingObserverResult, 'session' | 'speaker'> {
  results: ApplyMeetingSignalResult[];
  session: LocalTimelineObserverCandidateResult;
  speaker: ActiveSpeakerTimelineObserverResult;
}

export function normalizeBrowserMeetingCandidate(
  input?: BrowserMeetingSnapshot,
  options?: BrowserMeetingObserverOptions,
): NormalizedMeetingSessionCandidate | null;

export function normalizeBrowserMeetingCandidates(
  input?: BrowserMeetingSnapshot | BrowserMeetingSnapshot[],
  options?: BrowserMeetingObserverOptions,
): NormalizedMeetingSessionCandidate[];

export function selectBrowserMeetingCandidate(
  input?: BrowserMeetingSnapshot | BrowserMeetingSnapshot[],
  options?: BrowserMeetingObserverOptions,
): Record<string, unknown> & { normalizedCandidates: NormalizedMeetingSessionCandidate[] };

export function normalizeBrowserActiveSpeakerSample(
  input?: BrowserMeetingSnapshot,
  options?: BrowserMeetingObserverOptions & { candidate?: NormalizedMeetingSessionCandidate | Record<string, unknown> | null },
): NormalizedActiveSpeakerSample | null;

export function observeBrowserMeetingSample(
  state?: BrowserMeetingObserverState | null,
  input?: BrowserMeetingSnapshot,
  options?: BrowserMeetingObserverOptions,
): BrowserMeetingObserverResult;

export function createBrowserMeetingObserver(options?: BrowserMeetingObserverOptions): {
  observe(input?: BrowserMeetingSnapshot, options?: BrowserMeetingObserverOptions): BrowserMeetingObserverResult;
  getState(): BrowserMeetingObserverState;
  reset(nextState?: BrowserMeetingObserverState | null): BrowserMeetingObserverState;
};

export function createBrowserMeetingTimelineObserver(
  client: MeetingTimelineClient,
  options?: BrowserMeetingObserverOptions,
): {
  observe(input?: BrowserMeetingSnapshot, options?: BrowserMeetingObserverOptions): Promise<BrowserMeetingTimelineObserverResult>;
  getState(): BrowserMeetingObserverState;
  reset(nextState?: BrowserMeetingObserverState | null): BrowserMeetingObserverState;
};
