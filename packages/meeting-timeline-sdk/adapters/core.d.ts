import type { MeetingTimelineClient } from '../index.mjs';

export type MeetingSignalType =
  | 'meeting_started'
  | 'meeting_ended'
  | 'participant_joined'
  | 'participant_left'
  | 'speaker_started'
  | 'speaker_ended'
  | 'artifact_ready'
  | 'subscription_lifecycle';

export type MeetingSignalSource =
  | 'webhook'
  | 'long_connection'
  | 'polling'
  | 'local_detector'
  | 'manual'
  | string;

export interface NormalizedMeetingIdentity {
  platform: string;
  meeting_id: string;
  external_meeting_id?: string;
  meeting_url?: string;
  minute_token?: string;
  title?: string;
  organizer_id?: string;
  organizer_name?: string;
}

export interface NormalizedMeetingSignal {
  type: MeetingSignalType;
  meeting?: NormalizedMeetingIdentity;
  platform?: string;
  occurred_at_ms: number;
  source_event_id?: string;
  source?: MeetingSignalSource;
  participant_id?: string;
  participant_name?: string;
  speaker_id?: string;
  speaker_name?: string;
  artifact_kind?: 'transcript' | 'recording' | 'smart_notes' | string;
  artifact_id?: string;
  artifact_url?: string;
  lifecycle_type?: 'expiration_reminder' | 'expired' | 'suspended' | 'reauthorization_required' | 'subscription_removed' | 'missed' | string;
  subscription_id?: string;
  subscription_name?: string;
  expires_at_ms?: number;
  resource?: string;
  tenant_id?: string;
  client_state?: string;
  raw?: unknown;
}

export interface ApplyMeetingSignalOptions {
  defaults?: Record<string, unknown>;
  participantAsAnnotation?: boolean;
  participantAsMark?: boolean;
  speakerAsAnnotation?: boolean;
  speakerAsMark?: boolean;
  onParticipantSignal?: (signal: NormalizedMeetingSignal, client: MeetingTimelineClient) => unknown | Promise<unknown>;
  onSpeakerSignal?: (signal: NormalizedMeetingSignal, client: MeetingTimelineClient) => unknown | Promise<unknown>;
  onArtifactSignal?: (signal: NormalizedMeetingSignal, client: MeetingTimelineClient) => unknown | Promise<unknown>;
  onSubscriptionLifecycleSignal?: (signal: NormalizedMeetingSignal, client: MeetingTimelineClient) => unknown | Promise<unknown>;
}

export interface ApplyMeetingSignalResult {
  applied: boolean;
  action: string;
  reason?: string;
  signal: NormalizedMeetingSignal;
  response?: unknown;
}

export const MEETING_SIGNAL_TYPES: readonly MeetingSignalType[];

export function normalizeMeetingIdentity(input?: Record<string, unknown>, defaults?: Record<string, unknown>): NormalizedMeetingIdentity;
export function normalizeMeetingSignal(input?: Record<string, unknown>, defaults?: Record<string, unknown>): NormalizedMeetingSignal;
export function applyMeetingSignal(
  client: MeetingTimelineClient,
  signal: NormalizedMeetingSignal | Record<string, unknown>,
  options?: ApplyMeetingSignalOptions,
): Promise<ApplyMeetingSignalResult>;
export function applyMeetingSignals(
  client: MeetingTimelineClient,
  signals: Array<NormalizedMeetingSignal | Record<string, unknown>> | NormalizedMeetingSignal | Record<string, unknown>,
  options?: ApplyMeetingSignalOptions,
): Promise<ApplyMeetingSignalResult[]>;
