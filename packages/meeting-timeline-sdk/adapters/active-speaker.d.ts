import type { MeetingTimelineClient } from '../index.mjs';
import type { ApplyMeetingSignalOptions, ApplyMeetingSignalResult, NormalizedMeetingSignal } from './core.mjs';

export interface ActiveSpeakerMeetingIdentity {
  platform: string;
  meeting_id: string;
  external_meeting_id?: string;
  meeting_url?: string;
  title?: string;
}

export interface ActiveSpeakerRawSample {
  meeting?: Record<string, unknown>;
  meeting_id?: string;
  meetingId?: string;
  external_meeting_id?: string;
  externalMeetingId?: string;
  platform?: string;
  provider?: string;
  detected_platform?: string;
  detectedPlatform?: string;
  url?: string;
  meeting_url?: string;
  meetingUrl?: string;
  join_url?: string;
  joinUrl?: string;
  title?: string;
  topic?: string;
  active_speaker?: Record<string, unknown>;
  activeSpeaker?: Record<string, unknown>;
  speaker?: Record<string, unknown>;
  participant?: Record<string, unknown>;
  user?: Record<string, unknown>;
  speaker_id?: string;
  speakerId?: string;
  speaker_name?: string;
  speakerName?: string;
  participant_id?: string;
  participantId?: string;
  participant_name?: string;
  participantName?: string;
  speaking?: boolean;
  is_speaking?: boolean;
  isSpeaking?: boolean;
  active?: boolean;
  observed_at_ms?: number | string | Date;
  observedAtMs?: number | string | Date;
  detected_at_ms?: number | string | Date;
  detectedAtMs?: number | string | Date;
  occurred_at_ms?: number | string | Date;
  occurredAtMs?: number | string | Date;
  timestamp_ms?: number | string | Date;
  timestampMs?: number | string | Date;
  timestamp?: number | string | Date;
  source_event_id?: string;
  sourceEventId?: string;
  event_id?: string;
  eventId?: string;
  id?: string;
  [key: string]: unknown;
}

export interface NormalizedActiveSpeakerSample {
  meeting: ActiveSpeakerMeetingIdentity;
  observed_at_ms: number;
  speaking: boolean;
  speaker_id?: string;
  speaker_name?: string;
  participant_id?: string;
  participant_name?: string;
  source_event_id?: string;
  raw?: ActiveSpeakerRawSample;
}

export interface ActiveSpeakerObserverState {
  activeSpeaker: NormalizedActiveSpeakerSample | null;
  activeSinceMs: number | null;
  candidateSpeaker: NormalizedActiveSpeakerSample | null;
  candidateSinceMs: number | null;
  silenceSinceMs: number | null;
  lastObservedAtMs: number | null;
}

export interface ActiveSpeakerObserverOptions {
  initialState?: ActiveSpeakerObserverState | null;
  minStableMs?: number;
  min_stable_ms?: number;
  switchStableMs?: number;
  switch_stable_ms?: number;
  endIdleMs?: number;
  end_idle_ms?: number;
  silenceEndMs?: number;
  silence_end_ms?: number;
  emitSpeakerEnd?: boolean;
  emit_speaker_end?: boolean;
  source?: string;
  observedAtMs?: number | string | Date;
  receivedAtMs?: number | string | Date;
  applyOptions?: ApplyMeetingSignalOptions;
  [key: string]: unknown;
}

export interface ActiveSpeakerObserverResult {
  state: ActiveSpeakerObserverState;
  signals: NormalizedMeetingSignal[];
  sample: NormalizedActiveSpeakerSample | null;
}

export interface ActiveSpeakerObserverManyResult {
  state: ActiveSpeakerObserverState;
  signals: NormalizedMeetingSignal[];
  observations: ActiveSpeakerObserverResult[];
}

export interface ActiveSpeakerTimelineObserverResult extends ActiveSpeakerObserverResult {
  results: ApplyMeetingSignalResult[];
}

export interface ActiveSpeakerTimelineObserverManyResult {
  state: ActiveSpeakerObserverState;
  signals: NormalizedMeetingSignal[];
  observations: ActiveSpeakerObserverResult[];
  results: ApplyMeetingSignalResult[];
}

export function normalizeActiveSpeakerSample(
  raw?: ActiveSpeakerRawSample,
  options?: ActiveSpeakerObserverOptions,
): NormalizedActiveSpeakerSample | null;

export function observeActiveSpeakerSample(
  state?: ActiveSpeakerObserverState | null,
  raw?: ActiveSpeakerRawSample,
  options?: ActiveSpeakerObserverOptions,
): ActiveSpeakerObserverResult;

export function observeActiveSpeakerSamples(
  state?: ActiveSpeakerObserverState | null,
  rows?: ActiveSpeakerRawSample[],
  options?: ActiveSpeakerObserverOptions,
): ActiveSpeakerObserverManyResult;

export function createActiveSpeakerObserver(options?: ActiveSpeakerObserverOptions): {
  observe(sample?: ActiveSpeakerRawSample, options?: ActiveSpeakerObserverOptions): ActiveSpeakerObserverResult;
  observeMany(samples?: ActiveSpeakerRawSample[], options?: ActiveSpeakerObserverOptions): ActiveSpeakerObserverManyResult;
  getState(): ActiveSpeakerObserverState;
  reset(nextState?: ActiveSpeakerObserverState | null): ActiveSpeakerObserverState;
};

export function createActiveSpeakerTimelineObserver(
  client: MeetingTimelineClient,
  options?: ActiveSpeakerObserverOptions,
): {
  observe(sample?: ActiveSpeakerRawSample, options?: ActiveSpeakerObserverOptions & { applyOptions?: ApplyMeetingSignalOptions }): Promise<ActiveSpeakerTimelineObserverResult>;
  observeMany(samples?: ActiveSpeakerRawSample[], options?: ActiveSpeakerObserverOptions & { applyOptions?: ApplyMeetingSignalOptions }): Promise<ActiveSpeakerTimelineObserverManyResult>;
  getState(): ActiveSpeakerObserverState;
  reset(nextState?: ActiveSpeakerObserverState | null): ActiveSpeakerObserverState;
};
