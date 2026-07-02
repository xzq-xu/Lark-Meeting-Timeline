import type { NormalizedMeetingSignal } from './core.mjs';

export interface MeetingSignalReconcilerStateEntry {
  key?: string;
  type?: string;
  source?: string;
  platform?: string;
  meeting_id?: string;
  meeting_url?: string;
  source_event_id?: string;
  occurred_at_ms?: number;
  [key: string]: unknown;
}

export interface MeetingSignalReconcilerState {
  seen: MeetingSignalReconcilerStateEntry[];
  active_meetings: MeetingSignalReconcilerStateEntry[];
  ended_meetings: MeetingSignalReconcilerStateEntry[];
  recent_speakers: MeetingSignalReconcilerStateEntry[];
}

export interface MeetingSignalReconcilerOptions {
  initialState?: Partial<MeetingSignalReconcilerState>;
  initial_state?: Partial<MeetingSignalReconcilerState>;
  duplicateWindowMs?: number;
  duplicate_window_ms?: number;
  speakerRepeatWindowMs?: number;
  speaker_repeat_window_ms?: number;
  maxSeen?: number;
  max_seen?: number;
  maxMeetings?: number;
  max_meetings?: number;
  maxSpeakers?: number;
  max_speakers?: number;
  [key: string]: unknown;
}

export interface MeetingSignalReconcilerDecision {
  action: 'apply' | 'skip';
  reason?: string;
  signal: NormalizedMeetingSignal;
  fingerprint?: string;
  existing?: MeetingSignalReconcilerStateEntry;
  [key: string]: unknown;
}

export interface MeetingSignalReconciliationResult {
  signals: NormalizedMeetingSignal[];
  skipped: MeetingSignalReconcilerDecision[];
  decisions: MeetingSignalReconcilerDecision[];
  state: MeetingSignalReconcilerState;
}

export function meetingSignalFingerprint(signal: NormalizedMeetingSignal | Record<string, unknown>): string;

export function reconcileMeetingSignals(
  signals?: Array<NormalizedMeetingSignal | Record<string, unknown>> | NormalizedMeetingSignal | Record<string, unknown>,
  state?: Partial<MeetingSignalReconcilerState>,
  options?: MeetingSignalReconcilerOptions,
): MeetingSignalReconciliationResult;

export function createMeetingSignalReconciler(options?: MeetingSignalReconcilerOptions): {
  reconcile(
    signals?: Array<NormalizedMeetingSignal | Record<string, unknown>> | NormalizedMeetingSignal | Record<string, unknown>,
    options?: MeetingSignalReconcilerOptions,
  ): MeetingSignalReconciliationResult;
  getState(): MeetingSignalReconcilerState;
  reset(nextState?: Partial<MeetingSignalReconcilerState>): MeetingSignalReconcilerState;
};
