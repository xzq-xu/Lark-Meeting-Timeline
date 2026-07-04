export const MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA: 'meeting_platform_runtime_event';
export const MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA_VERSION: 1;
export const MEETING_PLATFORM_RUNTIME_EVENT_ENDPOINT: '/api/meeting-platform/runtime-events';
export const MEETING_PLATFORM_RUNTIME_EVENT_ACTIONS: readonly string[];

export interface MeetingPlatformRuntimeEventOptions {
  action?: string;
  kind?: string;
  platform?: string;
  source?: string;
  payload?: unknown;
  event_id?: string;
  eventId?: string;
  correlation_id?: string;
  correlationId?: string;
  sent_at_ms?: number | string | Date;
  sentAtMs?: number | string | Date;
  sent_at?: number | string | Date;
  sentAt?: number | string | Date;
  now?: () => number | string | Date;
  clock?: () => number | string | Date;
  [key: string]: unknown;
}

export interface MeetingPlatformRuntimeEvent {
  type: typeof MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA;
  schema: typeof MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA;
  schema_version: typeof MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA_VERSION;
  action: string;
  platform?: string;
  source?: string;
  event_id?: string;
  correlation_id?: string;
  sent_at_ms: number;
  payload?: unknown;
  snapshot?: unknown;
  provider_event?: unknown;
  annotation?: unknown;
  current_meeting?: Record<string, unknown>;
  signals?: unknown[];
  meeting?: Record<string, unknown>;
  annotations?: unknown[];
  speakerTrack?: Record<string, unknown>;
  participantTrack?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingPlatformRuntimeEventClientOptions extends MeetingPlatformRuntimeEventOptions {
  baseUrl?: string;
  base_url?: string;
  endpoint?: string;
  path?: string;
  runtimeEventEndpoint?: string;
  runtime_event_endpoint?: string;
  fetch?: typeof fetch;
  fetchImpl?: typeof fetch;
  headers?: Record<string, string>;
}

export interface MeetingPlatformRuntimeEventSendOptions extends MeetingPlatformRuntimeEventOptions {
  endpoint?: string;
  fetchOptions?: RequestInit;
  fetch_options?: RequestInit;
  headers?: Record<string, string>;
  signal?: AbortSignal;
}

export interface MeetingPlatformRuntimeEventClient {
  type: 'meeting_platform_runtime_event_client';
  schema: typeof MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA;
  schema_version: typeof MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA_VERSION;
  endpoint: string;
  send(eventInput?: Record<string, unknown> | MeetingPlatformRuntimeEvent, sendOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  observeMeetingApp(platform: string, snapshot?: Record<string, unknown>, observeOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  ingestProvider(platform: string, payload?: unknown, ingestOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  insertAnnotation(platform: string, annotationInput?: Record<string, unknown>, markOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  insertMark(platform: string, annotationInput?: Record<string, unknown>, markOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  speakerTrack(platform: string, input?: Record<string, unknown>, trackOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  participantTrack(platform: string, input?: Record<string, unknown>, trackOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  timelineView(platform: string, input?: Record<string, unknown>, viewOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  runtimeBundles(bundleOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  registry(registryOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  manifest(manifestOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  readiness(readinessOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  handoffReadiness(readinessOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
}

export function normalizeMeetingPlatformRuntimeEventAction(action: string): string;
export function buildMeetingPlatformRuntimeEvent(
  input?: Record<string, unknown>,
  options?: MeetingPlatformRuntimeEventOptions,
): MeetingPlatformRuntimeEvent;
export function assertMeetingPlatformRuntimeEvent(
  event?: Record<string, unknown> | MeetingPlatformRuntimeEvent,
  options?: MeetingPlatformRuntimeEventOptions,
): MeetingPlatformRuntimeEvent;
export function buildMeetingPlatformObserveRuntimeEvent(
  platform: string,
  snapshot?: Record<string, unknown>,
  options?: MeetingPlatformRuntimeEventOptions,
): MeetingPlatformRuntimeEvent;
export function buildMeetingPlatformProviderRuntimeEvent(
  platform: string,
  payload?: unknown,
  options?: MeetingPlatformRuntimeEventOptions,
): MeetingPlatformRuntimeEvent;
export function buildMeetingPlatformAnnotationRuntimeEvent(
  platform: string,
  annotationInput?: Record<string, unknown>,
  options?: MeetingPlatformRuntimeEventOptions,
): MeetingPlatformRuntimeEvent;
export function buildMeetingPlatformSpeakerTrackRuntimeEvent(
  platform: string,
  input?: Record<string, unknown>,
  options?: MeetingPlatformRuntimeEventOptions,
): MeetingPlatformRuntimeEvent;
export function buildMeetingPlatformParticipantTrackRuntimeEvent(
  platform: string,
  input?: Record<string, unknown>,
  options?: MeetingPlatformRuntimeEventOptions,
): MeetingPlatformRuntimeEvent;
export function buildMeetingPlatformTimelineViewRuntimeEvent(
  platform: string,
  input?: Record<string, unknown>,
  options?: MeetingPlatformRuntimeEventOptions,
): MeetingPlatformRuntimeEvent;
export function meetingPlatformRuntimeEventEndpoint(options?: MeetingPlatformRuntimeEventClientOptions): string;
export function createMeetingPlatformRuntimeEventClient(
  options?: MeetingPlatformRuntimeEventClientOptions,
): MeetingPlatformRuntimeEventClient;
