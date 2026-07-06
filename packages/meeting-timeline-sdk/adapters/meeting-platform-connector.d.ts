import type { NormalizedMeetingSignal } from './core.mjs';
import type {
  MeetingPlatformRuntimeEvent,
  MeetingPlatformRuntimeEventClient,
  MeetingPlatformRuntimeEventSendOptions,
} from './platform-runtime-event.mjs';

export const MEETING_PLATFORM_CONNECTOR_SCHEMA: 'meeting_platform_connector';
export const MEETING_PLATFORM_CONNECTOR_MATRIX_SCHEMA: 'meeting_platform_connector_matrix';
export const MEETING_PLATFORM_CONNECTOR_ACCEPTANCE_SCHEMA: 'meeting_platform_connector_acceptance';
export const MEETING_PLATFORM_CONNECTOR_RUNTIME_SCHEMA: 'meeting_platform_connector_runtime';
export const MEETING_PLATFORM_CONNECTOR_HUB_SCHEMA: 'meeting_platform_connector_hub';
export const MEETING_PLATFORM_CONNECTOR_RESOLUTION_SCHEMA: 'meeting_platform_connector_resolution';
export const MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION: 1;

export interface MeetingPlatformConnectorOptions {
  platform?: string;
  platform_key?: string;
  baseUrl?: string;
  base_url?: string;
  endpoint?: string;
  runtimeEventEndpoint?: string;
  runtime_event_endpoint?: string;
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  requireProviderReady?: boolean;
  require_provider_ready?: boolean;
  assertConnector?: boolean;
  assert_connector?: boolean;
  fetch?: typeof fetch;
  fetchImpl?: typeof fetch;
  headers?: Record<string, string>;
  env?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingPlatformConnector {
  type: 'meeting_platform_connector';
  schema: typeof MEETING_PLATFORM_CONNECTOR_SCHEMA;
  schema_version: typeof MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION;
  platform: string;
  display_name?: string;
  aliases: string[];
  base_url?: string;
  objective: string;
  event_adapter: Record<string, unknown>;
  provider: Record<string, unknown>;
  browser_observer: Record<string, unknown>;
  runtime_events: {
    endpoint?: string;
    event_schema?: string;
    client_factory?: string;
    supported_actions: readonly string[];
    action_count: number;
    realtime_contract?: Record<string, unknown>;
  };
  timeline_ingest: Record<string, unknown>;
  realtime_policy: Record<string, unknown>;
  adapter_route: Record<string, unknown>;
  host: Record<string, unknown>;
  sdk: Record<string, unknown>;
  readiness: Record<string, unknown>;
  commands: Record<string, string>;
  next_actions: string[];
}

export interface MeetingPlatformConnectorAcceptanceReport {
  type: 'meeting_platform_connector_acceptance';
  schema: typeof MEETING_PLATFORM_CONNECTOR_ACCEPTANCE_SCHEMA;
  schema_version: typeof MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION;
  accepted: boolean;
  platform: string;
  blocking_count: number;
  warning_count: number;
  issues: Array<Record<string, unknown>>;
  connector: MeetingPlatformConnector;
  next_actions: string[];
}

export interface MeetingPlatformConnectorMatrix {
  type: 'meeting_platform_connector_matrix';
  schema: typeof MEETING_PLATFORM_CONNECTOR_MATRIX_SCHEMA;
  schema_version: typeof MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION;
  platform_count: number;
  accepted_count: number;
  realtime_ready_count: number;
  provider_ready_count: number;
  candidate_observer_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  connectors: MeetingPlatformConnector[];
  acceptance_reports: MeetingPlatformConnectorAcceptanceReport[];
  registry_manifest: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingPlatformConnectorResolution {
  type: 'meeting_platform_connector_resolution';
  schema: typeof MEETING_PLATFORM_CONNECTOR_RESOLUTION_SCHEMA;
  schema_version: typeof MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION;
  detected: boolean;
  supported: boolean;
  platform?: string;
  reason?: string;
  current_platforms: string[];
  meeting?: Record<string, unknown>;
  browser?: Record<string, unknown>;
  candidate_count?: number;
  next_actions?: string[];
}

export interface MeetingPlatformConnectorHub {
  type: 'meeting_platform_connector_hub';
  schema: typeof MEETING_PLATFORM_CONNECTOR_HUB_SCHEMA;
  schema_version: typeof MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION;
  objective: string;
  accepted: boolean;
  blocking_count: number;
  warning_count: number;
  platform_count: number;
  accepted_count: number;
  realtime_ready_count: number;
  candidate_observer_count: number;
  platforms: string[];
  default_platform: string;
  runtime_event_endpoint?: string;
  routing: Record<string, unknown>;
  matrix: MeetingPlatformConnectorMatrix;
  connectors: MeetingPlatformConnector[];
  readiness: Record<string, unknown>;
  issues: Array<Record<string, unknown>>;
  next_actions: string[];
}

export interface MeetingPlatformConnectorRuntime {
  type: 'meeting_platform_connector_runtime';
  schema: typeof MEETING_PLATFORM_CONNECTOR_RUNTIME_SCHEMA;
  schema_version: typeof MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION;
  platform: string;
  endpoint: string;
  connector: MeetingPlatformConnector;
  runtime_event_client: MeetingPlatformRuntimeEventClient;
  supported_actions: readonly string[];
  supports(action: string): boolean;
  assertSupported(action: string): string;
  normalizeProviderEvent(raw?: unknown, normalizeOptions?: Record<string, unknown>): NormalizedMeetingSignal[];
  buildEvent(input?: Record<string, unknown>, eventOptions?: Record<string, unknown>): MeetingPlatformRuntimeEvent;
  send(input?: Record<string, unknown>, sendOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  observeMeetingApp(snapshot?: Record<string, unknown>, observeOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  observePlatformCandidates(input?: Record<string, unknown>, observeOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  ingestProvider(payload?: unknown, ingestOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  insertAnnotation(annotationInput?: Record<string, unknown>, markOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  insertMark(annotationInput?: Record<string, unknown>, markOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  speakerTrack(input?: Record<string, unknown>, trackOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  participantTrack(input?: Record<string, unknown>, trackOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  timelineView(input?: Record<string, unknown>, viewOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  adapterRoute(routeOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  adapterRoutes(routeOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  runtimeBundles(bundleOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  registry(registryOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  manifest(manifestOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  readiness(readinessOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  handoffReadiness(readinessOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  runManifest(manifestOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  runHandoffReadiness(readinessOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
}

export interface MeetingPlatformConnectorHubRuntime {
  type: 'meeting_platform_connector_hub_runtime';
  schema: typeof MEETING_PLATFORM_CONNECTOR_HUB_SCHEMA;
  schema_version: typeof MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION;
  hub: MeetingPlatformConnectorHub;
  matrix: MeetingPlatformConnectorMatrix;
  platforms: readonly string[];
  default_platform: string;
  endpoint?: string;
  connectors: MeetingPlatformConnector[];
  resolvePlatform(input?: unknown, resolveOptions?: MeetingPlatformConnectorOptions): MeetingPlatformConnectorResolution;
  connectorFor(input?: unknown, connectorOptions?: MeetingPlatformConnectorOptions): MeetingPlatformConnector;
  runtimeFor(input?: unknown, runtimeOptions?: MeetingPlatformConnectorOptions): MeetingPlatformConnectorRuntime;
  supports(input: unknown, action: string): boolean;
  normalizeProviderEvent(input?: unknown, raw?: unknown, normalizeOptions?: MeetingPlatformConnectorOptions): NormalizedMeetingSignal[];
  buildEvent(input?: Record<string, unknown>, eventOptions?: MeetingPlatformConnectorOptions): MeetingPlatformRuntimeEvent;
  send(input?: Record<string, unknown>, sendOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  observeMeetingApp(snapshot?: Record<string, unknown>, observeOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  observePlatformCandidates(input?: Record<string, unknown>, observeOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  ingestProvider(input?: unknown, payload?: unknown, ingestOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  insertAnnotation(input?: Record<string, unknown>, annotation?: Record<string, unknown>, markOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  insertMark(input?: Record<string, unknown>, annotation?: Record<string, unknown>, markOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  speakerTrack(input?: Record<string, unknown>, track?: Record<string, unknown>, trackOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  participantTrack(input?: Record<string, unknown>, track?: Record<string, unknown>, trackOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  timelineView(input?: Record<string, unknown>, view?: Record<string, unknown>, viewOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  adapterRoute(input?: unknown, routeOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  adapterRoutes(routeOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  runtimeBundles(bundleOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  registry(registryOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  manifest(manifestOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  readiness(readinessOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  handoffReadiness(readinessOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  runManifest(manifestOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
  runHandoffReadiness(readinessOptions?: MeetingPlatformRuntimeEventSendOptions): Promise<unknown>;
}

export function buildMeetingPlatformConnector(
  platformOrOptions?: string | MeetingPlatformConnectorOptions,
  options?: MeetingPlatformConnectorOptions,
): MeetingPlatformConnector;

export function buildMeetingPlatformConnectorAcceptanceReport(
  connectorOrOptions?: MeetingPlatformConnector | string | MeetingPlatformConnectorOptions,
  options?: MeetingPlatformConnectorOptions,
): MeetingPlatformConnectorAcceptanceReport;

export function assertMeetingPlatformConnector(
  connectorOrOptions?: MeetingPlatformConnector | string | MeetingPlatformConnectorOptions,
  options?: MeetingPlatformConnectorOptions,
): MeetingPlatformConnectorAcceptanceReport;

export function buildMeetingPlatformConnectorMatrix(
  options?: MeetingPlatformConnectorOptions,
): MeetingPlatformConnectorMatrix;

export function resolveMeetingPlatformConnectorInput(
  input?: unknown,
  options?: MeetingPlatformConnectorOptions,
): MeetingPlatformConnectorResolution;

export function buildMeetingPlatformConnectorHub(
  options?: MeetingPlatformConnectorOptions,
): MeetingPlatformConnectorHub;

export function createMeetingPlatformConnectorRuntime(
  platformOrConnector: string | MeetingPlatformConnector,
  options?: MeetingPlatformConnectorOptions,
): MeetingPlatformConnectorRuntime;

export function createMeetingPlatformConnectorHub(
  options?: MeetingPlatformConnectorOptions,
): MeetingPlatformConnectorHubRuntime;

export function buildDefaultMeetingPlatformConnectorMatrix(
  options?: MeetingPlatformConnectorOptions,
): MeetingPlatformConnectorMatrix;

export function buildDefaultMeetingPlatformConnectorHub(
  options?: MeetingPlatformConnectorOptions,
): MeetingPlatformConnectorHub;
