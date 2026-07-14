import type { NormalizedMeetingSignal } from './core.mjs';

export const LARK_MEETING_EVENT_TYPES: readonly string[];

export function normalizeLarkEvent(raw?: unknown, options?: Record<string, unknown>): NormalizedMeetingSignal[];
