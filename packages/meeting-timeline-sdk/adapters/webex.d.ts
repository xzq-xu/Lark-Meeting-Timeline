export interface WebexNormalizeOptions {
  receivedAtMs?: number | string | Date;
  [key: string]: unknown;
}

export function normalizeWebexEvent(raw?: unknown, options?: WebexNormalizeOptions): Record<string, unknown>[];
