import type { NormalizedMeetingSignal } from './core.mjs';

export interface NormalizeZoomEventOptions {
  receivedAtMs?: number | string | Date;
}

export function normalizeZoomEvent(raw?: unknown, options?: NormalizeZoomEventOptions): NormalizedMeetingSignal[];
