import type { NormalizedMeetingSignal } from './core.mjs';

export const LOCAL_DETECTOR_EVENT_TYPES: readonly string[];

export function normalizeLocalDetectorEvent(raw?: unknown, options?: Record<string, unknown>): NormalizedMeetingSignal[];
