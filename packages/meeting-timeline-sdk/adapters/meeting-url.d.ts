export interface DetectedMeetingFromUrl {
  platform: string;
  meeting_id: string;
  external_meeting_id?: string;
  meeting_url?: string;
  title?: string;
  confidence?: 'high' | 'medium' | 'low' | string;
}

export function stableMeetingIdFromUrl(url?: string | URL): string | null;

export function detectMeetingFromUrl(input?: string | URL | Record<string, unknown>): DetectedMeetingFromUrl | null;

export function detectMeetingPlatformFromUrl(input?: string | URL | Record<string, unknown>): string | null;
