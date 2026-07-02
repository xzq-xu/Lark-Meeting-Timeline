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

export interface NativeMeetingSnapshot {
  application?: Record<string, unknown>;
  app?: Record<string, unknown>;
  process?: Record<string, unknown>;
  window?: Record<string, unknown>;
  accessibility?: Record<string, unknown>;
  ax?: Record<string, unknown>;
  audio?: Record<string, unknown>;
  meeting?: Record<string, unknown>;
  activeSpeaker?: Record<string, unknown>;
  active_speaker?: Record<string, unknown>;
  speaker?: Record<string, unknown>;
  participants?: Record<string, unknown>[];
  url?: string;
  deep_link?: string;
  deepLink?: string;
  title?: string;
  platform?: string;
  provider?: string;
  meeting_id?: string;
  meetingId?: string;
  processName?: string;
  process_name?: string;
  bundleId?: string;
  bundle_id?: string;
  active?: boolean;
  focused?: boolean;
  visible?: boolean;
  in_meeting?: boolean;
  inMeeting?: boolean;
  observed_at_ms?: number | string | Date;
  observedAtMs?: number | string | Date;
  timestamp_ms?: number | string | Date;
  timestampMs?: number | string | Date;
  timestamp?: number | string | Date;
  applications?: NativeMeetingSnapshot[];
  apps?: NativeMeetingSnapshot[];
  processes?: NativeMeetingSnapshot[];
  windows?: NativeMeetingSnapshot[];
  candidates?: NativeMeetingSnapshot[];
  snapshots?: NativeMeetingSnapshot[];
  items?: NativeMeetingSnapshot[];
  [key: string]: unknown;
}

export interface NativeMeetingObserverState {
  sessionState: LocalObserverState;
  speakerState: Record<string, unknown>;
}

export interface NativeMeetingObserverOptions extends MeetingSessionDiscoveryOptions {
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

export interface NativeMeetingObserverResult {
  state: NativeMeetingObserverState;
  signals: NormalizedMeetingSignal[];
  session: LocalObserverCandidateResult;
  speaker: ActiveSpeakerObserverResult;
  normalizedCandidates: NormalizedMeetingSessionCandidate[];
}

export interface NativeMeetingTimelineObserverResult extends Omit<NativeMeetingObserverResult, 'session' | 'speaker'> {
  results: ApplyMeetingSignalResult[];
  session: LocalTimelineObserverCandidateResult;
  speaker: ActiveSpeakerTimelineObserverResult;
}

export function normalizeNativeMeetingCandidate(
  input?: NativeMeetingSnapshot,
  options?: NativeMeetingObserverOptions,
): NormalizedMeetingSessionCandidate | null;

export function normalizeNativeMeetingCandidates(
  input?: NativeMeetingSnapshot | NativeMeetingSnapshot[],
  options?: NativeMeetingObserverOptions,
): NormalizedMeetingSessionCandidate[];

export function selectNativeMeetingCandidate(
  input?: NativeMeetingSnapshot | NativeMeetingSnapshot[],
  options?: NativeMeetingObserverOptions,
): Record<string, unknown> & { normalizedCandidates: NormalizedMeetingSessionCandidate[] };

export function normalizeNativeActiveSpeakerSample(
  input?: NativeMeetingSnapshot,
  options?: NativeMeetingObserverOptions & { candidate?: NormalizedMeetingSessionCandidate | Record<string, unknown> | null },
): NormalizedActiveSpeakerSample | null;

export function observeNativeMeetingSample(
  state?: NativeMeetingObserverState | null,
  input?: NativeMeetingSnapshot,
  options?: NativeMeetingObserverOptions,
): NativeMeetingObserverResult;

export function createNativeMeetingObserver(options?: NativeMeetingObserverOptions): {
  observe(input?: NativeMeetingSnapshot, options?: NativeMeetingObserverOptions): NativeMeetingObserverResult;
  getState(): NativeMeetingObserverState;
  reset(nextState?: NativeMeetingObserverState | null): NativeMeetingObserverState;
};

export function createNativeMeetingTimelineObserver(
  client: MeetingTimelineClient,
  options?: NativeMeetingObserverOptions,
): {
  observe(input?: NativeMeetingSnapshot, options?: NativeMeetingObserverOptions): Promise<NativeMeetingTimelineObserverResult>;
  getState(): NativeMeetingObserverState;
  reset(nextState?: NativeMeetingObserverState | null): NativeMeetingObserverState;
};
