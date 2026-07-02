import type { NormalizedMeetingSignal } from './core.mjs';

export interface MeetingPlatformEventAdapter {
  key: string;
  aliases: readonly string[];
  source: string;
  normalize(raw?: unknown, options?: Record<string, unknown>): NormalizedMeetingSignal[];
}

export const MEETING_PLATFORM_EVENT_ADAPTERS: readonly MeetingPlatformEventAdapter[];

export function meetingPlatformEventAdapterFor(platform: string): MeetingPlatformEventAdapter | null;
