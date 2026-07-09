import type { MeetingAppObserverOptions, MeetingAppSnapshot } from './meeting-apps.mjs';

export const MEETING_APP_DOM_CAPTURE_SCHEMA: string;
export const MEETING_APP_DOM_CAPTURE_SCHEMA_VERSION: number;

export type MeetingAppDomCaptureProfilePlatform = 'google_meet' | 'microsoft_teams' | 'zoom' | 'lark' | 'webex';

export interface MeetingAppDomCaptureProfile {
  platform: MeetingAppDomCaptureProfilePlatform;
  displayName: string;
  controlSelectors: string[];
  participantSelectors: string[];
  textSelectors: string[];
}

export interface MeetingAppDomSemanticSignal {
  type: string;
  source?: string;
  confidence?: string;
  label?: string;
  text?: string;
  role?: string;
  id?: string;
  participant_id?: string;
  participant_name?: string;
  participant_count?: number;
  [key: string]: unknown;
}

export interface MeetingAppDomInteractionState {
  in_call?: boolean;
  pre_join?: boolean;
  can_join?: boolean;
  waiting_room?: boolean;
  can_leave?: boolean;
  microphone_control_available?: boolean;
  camera_control_available?: boolean;
  screen_share_available?: boolean;
  screen_share_active?: boolean;
  captions_available?: boolean;
  recording_observed?: boolean;
  participant_roster_observed?: boolean;
  chat_available?: boolean;
  ai_summary_available?: boolean;
  active_speaker_candidate?: {
    id?: string;
    name?: string;
    display_name?: string;
    speaking?: boolean;
    audioLevel?: number;
    [key: string]: unknown;
  };
  participant_count?: number;
  [key: string]: unknown;
}

export interface MeetingAppDomControlSignalSummary {
  signal_types?: string[];
  signal_count?: number;
  join_available?: boolean;
  waiting_room?: boolean;
  leave_available?: boolean;
  microphone_available?: boolean;
  camera_available?: boolean;
  screen_share_available?: boolean;
  screen_share_active?: boolean;
  captions_available?: boolean;
  recording_observed?: boolean;
  participants_available?: boolean;
  chat_available?: boolean;
  ai_summary_available?: boolean;
  active_speaker_observed?: boolean;
  participant_roster_observed?: boolean;
  [key: string]: unknown;
}

export interface MeetingAppDomCaptureOptions extends MeetingAppObserverOptions {
  observedAtMs?: number | string | Date;
  observed_at_ms?: number | string | Date;
  source?: string;
  url?: string;
  title?: string;
  platform?: string;
  provider?: string;
  captureProfile?: string | false | null;
  capture_profile?: string | false | null;
  includeShadowDom?: boolean;
  include_shadow_dom?: boolean;
  deepDom?: boolean;
  deep_dom?: boolean;
  maxShadowRoots?: number;
  max_shadow_roots?: number;
  maxShadowHosts?: number;
  max_shadow_hosts?: number;
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
  inMeeting?: boolean;
  interaction?: MeetingAppDomInteractionState;
  semanticSignals?: MeetingAppDomSemanticSignal[];
  semanticSignalTypes?: string[];
  controlSignalSummary?: MeetingAppDomControlSignalSummary;
  page?: Record<string, unknown> & {
    inMeeting?: boolean;
    interaction?: MeetingAppDomInteractionState;
    semanticSignals?: MeetingAppDomSemanticSignal[];
    semanticSignalTypes?: string[];
    controlSignalSummary?: MeetingAppDomControlSignalSummary;
    buttons?: Record<string, unknown>[];
    controls?: Record<string, unknown>[];
    tiles?: Record<string, unknown>[];
    participants?: Record<string, unknown>[];
    texts?: Record<string, unknown>[];
  };
  dom?: Record<string, unknown> & {
    inMeeting?: boolean;
    interaction?: MeetingAppDomInteractionState;
    semanticSignals?: MeetingAppDomSemanticSignal[];
    semanticSignalTypes?: string[];
    controlSignalSummary?: MeetingAppDomControlSignalSummary;
  };
  capture?: {
    profile?: MeetingAppDomCaptureProfilePlatform;
    profile_display_name?: string;
    shadow_root_count?: number;
    control_count?: number;
    participant_count?: number;
    text_count?: number;
    semantic_signal_count?: number;
    semantic_signal_types?: string[];
    control_signal_summary?: MeetingAppDomControlSignalSummary;
  };
}

export const MEETING_APP_DOM_CAPTURE_PROFILES: Readonly<Record<MeetingAppDomCaptureProfilePlatform, MeetingAppDomCaptureProfile>>;

export function meetingAppDomCaptureProfile(
  platformOrInput?: string | MeetingAppDomCaptureInput | Record<string, unknown>,
  options?: MeetingAppDomCaptureOptions,
): MeetingAppDomCaptureProfile | null;

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
