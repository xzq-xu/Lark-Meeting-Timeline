export class MeetingTimelineSdkError extends Error {
  constructor(message: string, details?: Record<string, unknown>);
  name: 'MeetingTimelineSdkError';
  details: Record<string, unknown>;
}

export function compactObject<T>(value: T): T;

export function normalizeAbsoluteMs(value: unknown, fieldName?: string): number | undefined;
