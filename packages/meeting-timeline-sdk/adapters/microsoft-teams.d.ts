import type { NormalizedMeetingSignal } from './core.mjs';

export interface NormalizeMicrosoftTeamsEventOptions {
  receivedAtMs?: number | string | Date;
}

export function normalizeMicrosoftTeamsEvent(raw?: unknown, options?: NormalizeMicrosoftTeamsEventOptions): NormalizedMeetingSignal[];
