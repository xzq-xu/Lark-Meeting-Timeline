import type { MeetingAppObserverOptions, MeetingAppSnapshot } from './meeting-apps.mjs';

export const MEETING_APP_DOM_CAPTURE_SCHEMA: string;
export const MEETING_APP_DOM_CAPTURE_SCHEMA_VERSION: number;

export interface MeetingAppDomCaptureOptions extends MeetingAppObserverOptions {
  observedAtMs?: number | string | Date;
  observed_at_ms?: number | string | Date;
  source?: string;
  url?: string;
  title?: string;
  browserName?: string;
  browser_name?: string;
  maxControls?: number;
  max_controls?: number;
  maxParticipants?: number;
  max_participants?: number;
  maxTexts?: number;
  max_texts?: number;
  controlSelectors?: string[];
  control_selectors?: string[];
  participantSelectors?: string[];
  participant_selectors?: string[];
  textSelectors?: string[];
  text_selectors?: string[];
  [key: string]: unknown;
}

export interface MeetingAppDomCaptureInput {
  window?: {
    document?: Document | Record<string, unknown>;
    location?: Location | string | Record<string, unknown>;
    navigator?: Navigator | Record<string, unknown>;
    [key: string]: unknown;
  };
  document?: Document | Record<string, unknown>;
  defaultView?: Record<string, unknown>;
  location?: Location | string | Record<string, unknown>;
  url?: string;
  title?: string;
  observedAtMs?: number | string | Date;
  observed_at_ms?: number | string | Date;
  timestampMs?: number | string | Date;
  timestamp_ms?: number | string | Date;
  querySelectorAll?: (selector: string) => unknown[];
  [key: string]: unknown;
}

export interface MeetingAppDomCaptureSnapshot extends MeetingAppSnapshot {
  schema: string;
  schema_version: number;
  source?: string;
  observedAtMs: number;
  page?: Record<string, unknown> & {
    buttons?: Record<string, unknown>[];
    controls?: Record<string, unknown>[];
    tiles?: Record<string, unknown>[];
    participants?: Record<string, unknown>[];
    texts?: Record<string, unknown>[];
  };
  dom?: Record<string, unknown>;
  capture?: {
    control_count?: number;
    participant_count?: number;
    text_count?: number;
  };
}

export function captureMeetingAppDomSnapshot(
  input?: MeetingAppDomCaptureInput | Document,
  options?: MeetingAppDomCaptureOptions,
): MeetingAppDomCaptureSnapshot;

export function normalizeCapturedMeetingAppDomSnapshot(
  input?: MeetingAppDomCaptureInput | Document,
  options?: MeetingAppDomCaptureOptions,
): MeetingAppSnapshot | null;

export function normalizeCapturedMeetingAppDomSnapshots(
  inputs?: (MeetingAppDomCaptureInput | Document)[] | MeetingAppDomCaptureInput | Document,
  options?: MeetingAppDomCaptureOptions,
): MeetingAppSnapshot[];
