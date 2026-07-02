import type { NormalizedMeetingSignal } from './core.mjs';

export interface NormalizeGoogleMeetEventOptions {
  receivedAtMs?: number | string | Date;
}

export function normalizeGoogleMeetEvent(raw?: unknown, options?: NormalizeGoogleMeetEventOptions): NormalizedMeetingSignal[];
