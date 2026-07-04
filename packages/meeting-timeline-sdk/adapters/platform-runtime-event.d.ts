export const MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA: 'meeting_platform_runtime_event';
export const MEETING_PLATFORM_RUNTIME_EVENT_PLAN_SCHEMA: 'meeting_platform_runtime_event_plan';
export const MEETING_PLATFORM_RUNTIME_EVENT_PLAN_MATRIX_SCHEMA: 'meeting_platform_runtime_event_plan_matrix';
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
  windows?: unknown[];
  tabs?: unknown[];
  applications?: unknown[];
  candidates?: unknown[];
  environment?: unknown;
  [key: string]: unknown;
}

export interface MeetingPlatformRuntimeEventClientOptions extends MeetingPlatformRuntimeEventOptions {
  baseUrl?: string;
  base_url?: string;
  platforms?: string[];
  platform_keys?: string[];
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
  observePlatformCandidates(input?: Record<string, unknown>, observeOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  ingestProvider(platform: string, payload?: unknown, ingestOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  insertAnnotation(platform: string, annotationInput?: Record<string, unknown>, markOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  insertMark(platform: string, annotationInput?: Record<string, unknown>, markOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  speakerTrack(platform: string, input?: Record<string, unknown>, trackOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  participantTrack(platform: string, input?: Record<string, unknown>, trackOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  timelineView(platform: string, input?: Record<string, unknown>, viewOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  adapterRoute(platform: string, routeOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  adapterRoutes(routeOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  runtimeBundles(bundleOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  registry(registryOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  manifest(manifestOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  readiness(readinessOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  handoffReadiness(readinessOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  runManifest(manifestOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  runHandoffReadiness(readinessOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
}

export interface MeetingPlatformRuntimeEventPlan {
  type: 'meeting_platform_runtime_event_plan';
  schema: typeof MEETING_PLATFORM_RUNTIME_EVENT_PLAN_SCHEMA;
  schema_version: typeof MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA_VERSION;
  platform: string;
  endpoint: string;
  client_factory: string;
  event_schema: typeof MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA;
  supported_actions: readonly string[];
  provider_start_event_example?: string;
  provider_end_event_example?: string;
  realtime_contract: {
    primary_clock_field: string;
    provider_events_required_for_realtime: boolean;
    transcript_required_for_realtime: boolean;
    local_observer_required_for_reliable_start_end: boolean;
    annotation_should_use_device_capture_time: boolean;
  };
  imports: Record<string, string>;
  sequence: string[];
  actions: Record<string, unknown>[];
  examples: Record<string, MeetingPlatformRuntimeEvent>;
  next_actions: string[];
}

export interface MeetingPlatformRuntimeEventPlanMatrix {
  type: 'meeting_platform_runtime_event_plan_matrix';
  schema: typeof MEETING_PLATFORM_RUNTIME_EVENT_PLAN_MATRIX_SCHEMA;
  schema_version: typeof MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA_VERSION;
  platform_count: number;
  platforms: string[];
  realtime_provider_dependency_count: number;
  transcript_realtime_dependency_count: number;
  rows: Record<string, unknown>[];
  plans: MeetingPlatformRuntimeEventPlan[];
  next_actions: string[];
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
export function buildMeetingPlatformCandidateObservationRuntimeEvent(
  input?: Record<string, unknown>,
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
export function buildMeetingPlatformAdapterRouteRuntimeEvent(
  platform: string,
  input?: Record<string, unknown>,
  options?: MeetingPlatformRuntimeEventOptions,
): MeetingPlatformRuntimeEvent;
export function buildMeetingPlatformAdapterRoutesRuntimeEvent(
  input?: Record<string, unknown>,
  options?: MeetingPlatformRuntimeEventOptions,
): MeetingPlatformRuntimeEvent;
export function buildMeetingPlatformRunManifestRuntimeEvent(
  input?: Record<string, unknown>,
  options?: MeetingPlatformRuntimeEventOptions,
): MeetingPlatformRuntimeEvent;
export function buildMeetingPlatformRunHandoffReadinessRuntimeEvent(
  input?: Record<string, unknown>,
  options?: MeetingPlatformRuntimeEventOptions,
): MeetingPlatformRuntimeEvent;
export function buildMeetingPlatformRuntimeEventPlan(
  platform: string,
  options?: MeetingPlatformRuntimeEventClientOptions,
): MeetingPlatformRuntimeEventPlan;
export function buildMeetingPlatformRuntimeEventPlanMatrix(
  options?: MeetingPlatformRuntimeEventClientOptions,
): MeetingPlatformRuntimeEventPlanMatrix;
export function meetingPlatformRuntimeEventEndpoint(options?: MeetingPlatformRuntimeEventClientOptions): string;
export function createMeetingPlatformRuntimeEventClient(
  options?: MeetingPlatformRuntimeEventClientOptions,
): MeetingPlatformRuntimeEventClient;
