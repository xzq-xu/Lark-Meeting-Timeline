import type { TranscriptImportInput, TranscriptSegmentInput } from '../index.mjs';

export interface TranscriptNormalizeOptions {
  source?: string;
  language?: string;
  language_code?: string;
  platform?: string;
  meeting?: Record<string, unknown>;
  [key: string]: unknown;
}

export function parseTimedTextTranscript(raw?: string, options?: TranscriptNormalizeOptions): TranscriptSegmentInput[];
export function normalizeGoogleMeetTranscriptEntries(raw?: unknown, options?: TranscriptNormalizeOptions): TranscriptSegmentInput[];
export function normalizeMicrosoftTeamsTranscript(raw?: unknown, options?: TranscriptNormalizeOptions): TranscriptSegmentInput[];
export function normalizeZoomTranscript(raw?: unknown, options?: TranscriptNormalizeOptions): TranscriptSegmentInput[];
export function buildPlatformTranscriptImportPayload(input?: TranscriptImportInput & {
  raw?: unknown;
  content?: unknown;
  platform?: string;
}): Record<string, unknown>;
