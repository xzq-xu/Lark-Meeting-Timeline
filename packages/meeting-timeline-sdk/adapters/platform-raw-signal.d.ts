import type {
  MeetingPlatformRuntimeEvent,
  MeetingPlatformRuntimeEventOptions,
} from './platform-runtime-event.mjs';

export const MEETING_PLATFORM_RAW_SIGNAL_SCHEMA: 'meeting_platform_raw_signal';
export const MEETING_PLATFORM_RAW_SIGNAL_BATCH_SCHEMA: 'meeting_platform_raw_signal_batch';
export const MEETING_PLATFORM_RAW_SIGNAL_EXAMPLE_BATCH_SCHEMA: 'meeting_platform_raw_signal_example_batch';
export const MEETING_PLATFORM_RAW_SIGNAL_SCHEMA_VERSION: 1;
export const MEETING_PLATFORM_RAW_SIGNAL_KINDS: readonly [
  'meeting_app_snapshot',
  'platform_candidates',
  'provider_event',
  'annotation',
  'speaker_track',
  'participant_track',
];

export type MeetingPlatformRawSignalKind = typeof MEETING_PLATFORM_RAW_SIGNAL_KINDS[number];

export interface MeetingPlatformRawSignalOptions extends MeetingPlatformRuntimeEventOptions {
  platform?: string;
  kind?: string;
  action?: string;
  source?: string;
  deriveSpeakerTrack?: boolean;
  derive_speaker_track?: boolean;
  sample_at_ms?: number | string | Date;
  sampleAtMs?: number | string | Date;
  platforms?: string[];
  platform_keys?: string[];
  [key: string]: unknown;
}

export interface MeetingPlatformRawSignalInput {
  kind?: string;
  action?: string;
  platform?: string;
  provider?: string;
  source?: string;
  event_id?: string;
  eventId?: string;
  meeting_id?: string;
  meetingId?: string;
  meeting_url?: string;
  meetingUrl?: string;
  url?: string;
  title?: string;
  observed_at_ms?: number | string | Date;
  observedAtMs?: number | string | Date;
  captured_at_ms?: number | string | Date;
  capturedAtMs?: number | string | Date;
  occurred_at_ms?: number | string | Date;
  occurredAtMs?: number | string | Date;
  current_meeting?: Record<string, unknown>;
  currentMeeting?: Record<string, unknown>;
  meeting?: Record<string, unknown>;
  snapshot?: Record<string, unknown>;
  provider_event?: unknown;
  providerEvent?: unknown;
  annotation?: Record<string, unknown>;
  mark?: Record<string, unknown>;
  speaker?: Record<string, unknown> | string;
  active_speaker?: Record<string, unknown> | string;
  activeSpeaker?: Record<string, unknown> | string;
  participant?: Record<string, unknown> | string;
  participants?: unknown[];
  signals?: unknown[];
  windows?: unknown[];
  tabs?: unknown[];
  applications?: unknown[];
  candidates?: unknown[];
  payload?: unknown;
  [key: string]: unknown;
}

export interface MeetingPlatformRawSignal {
  type: typeof MEETING_PLATFORM_RAW_SIGNAL_SCHEMA;
  schema: typeof MEETING_PLATFORM_RAW_SIGNAL_SCHEMA;
  schema_version: typeof MEETING_PLATFORM_RAW_SIGNAL_SCHEMA_VERSION;
  kind: MeetingPlatformRawSignalKind;
  platform?: string;
  source?: string;
  event_id?: string;
  observed_at_ms: number;
  current_meeting?: Record<string, unknown>;
  payload?: unknown;
  runtime_action?: string;
  runtime_event_count: number;
  runtime_events: MeetingPlatformRuntimeEvent[];
}

export interface MeetingPlatformRawSignalBatch {
  type: typeof MEETING_PLATFORM_RAW_SIGNAL_BATCH_SCHEMA | typeof MEETING_PLATFORM_RAW_SIGNAL_EXAMPLE_BATCH_SCHEMA;
  schema: typeof MEETING_PLATFORM_RAW_SIGNAL_BATCH_SCHEMA | typeof MEETING_PLATFORM_RAW_SIGNAL_EXAMPLE_BATCH_SCHEMA;
  schema_version: typeof MEETING_PLATFORM_RAW_SIGNAL_SCHEMA_VERSION;
  signal_count: number;
  runtime_event_count: number;
  platform_count?: number;
  platforms: string[];
  kinds: string[];
  rows: Record<string, unknown>[];
  signals: MeetingPlatformRawSignal[];
  runtime_events: MeetingPlatformRuntimeEvent[];
}

export function normalizeMeetingPlatformRawSignalKind(kind: string): MeetingPlatformRawSignalKind;
export function buildMeetingPlatformRawSignal(
  input?: string | URL | MeetingPlatformRawSignalInput,
  options?: MeetingPlatformRawSignalOptions,
): MeetingPlatformRawSignal;
export function buildMeetingPlatformRuntimeEventsFromRawSignal(
  input?: string | URL | MeetingPlatformRawSignalInput,
  options?: MeetingPlatformRawSignalOptions,
): MeetingPlatformRuntimeEvent[];
export function buildMeetingPlatformRawSignalBatch(
  input?: MeetingPlatformRawSignalInput[] | Record<string, unknown>,
  options?: MeetingPlatformRawSignalOptions,
): MeetingPlatformRawSignalBatch;
export function buildMeetingPlatformRawSignalExamples(
  platform: string,
  options?: MeetingPlatformRawSignalOptions,
): MeetingPlatformRawSignalInput[];
export function buildMeetingPlatformRawSignalExampleBatch(
  options?: MeetingPlatformRawSignalOptions,
): MeetingPlatformRawSignalBatch;
