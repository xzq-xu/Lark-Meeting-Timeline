import type { MeetingTimelineClient } from '../index.mjs';
import type { MeetingPlatformEventAdapter } from './platform-registry.mjs';
import type { NormalizedMeetingSignal, ApplyMeetingSignalOptions, ApplyMeetingSignalResult } from './core.mjs';

export interface PlatformEventIngestInput {
  platform?: string;
  provider?: string;
  adapter?: string;
  payload?: unknown;
  body?: unknown;
  event?: unknown;
  raw?: unknown;
  receivedAtMs?: number | string | Date;
  received_at_ms?: number | string | Date;
  received_at?: number | string | Date;
  timestamp?: number | string | Date;
  ts?: number | string | Date;
  normalizerOptions?: Record<string, unknown>;
  normalizer_options?: Record<string, unknown>;
  applyOptions?: ApplyMeetingSignalOptions;
  apply_options?: ApplyMeetingSignalOptions;
  options?: PlatformEventIngestOptions;
}

export interface PlatformEventIngestOptions extends ApplyMeetingSignalOptions {
  receivedAtMs?: number | string | Date;
  received_at_ms?: number | string | Date;
  received_at?: number | string | Date;
  normalizerOptions?: Record<string, unknown>;
  normalizer_options?: Record<string, unknown>;
  applyOptions?: ApplyMeetingSignalOptions;
  apply_options?: ApplyMeetingSignalOptions;
}

export interface NormalizedPlatformEvent {
  adapter: MeetingPlatformEventAdapter;
  platform: string;
  source: string;
  signals: NormalizedMeetingSignal[];
}

export interface PlatformEventIngestResult extends NormalizedPlatformEvent {
  results: ApplyMeetingSignalResult[];
}

export function normalizePlatformEvent(
  platformOrInput: string | PlatformEventIngestInput,
  payload?: unknown,
  options?: PlatformEventIngestOptions,
): NormalizedPlatformEvent;

export function ingestPlatformEvent(
  client: MeetingTimelineClient,
  platformOrInput: string | PlatformEventIngestInput,
  payload?: unknown,
  options?: PlatformEventIngestOptions,
): Promise<PlatformEventIngestResult>;
