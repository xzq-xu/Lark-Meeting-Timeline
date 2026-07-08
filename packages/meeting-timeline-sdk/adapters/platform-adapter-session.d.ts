import type { MeetingPlatformAdapterLaunchPlan } from './platform-adapter-launch-plan.mjs';

export const MEETING_PLATFORM_ADAPTER_SESSION_SCHEMA: 'meeting_platform_adapter_session';
export const MEETING_PLATFORM_ADAPTER_SESSION_EVENT_SCHEMA: 'meeting_platform_adapter_session_event';
export const MEETING_PLATFORM_ADAPTER_SESSION_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdapterSessionClient {
  platformAdapterSelection?(platformOrPayload: string | Record<string, unknown>, payloadOrOptions?: Record<string, unknown>, options?: Record<string, unknown>): unknown | Promise<unknown>;
  adapterSelection?(platformOrPayload: string | Record<string, unknown>, payloadOrOptions?: Record<string, unknown>, options?: Record<string, unknown>): unknown | Promise<unknown>;
  platformRawSignalBatch?(payload: Record<string, unknown>, options?: Record<string, unknown>): unknown | Promise<unknown>;
  rawSignalBatch?(payload: Record<string, unknown>, options?: Record<string, unknown>): unknown | Promise<unknown>;
  validateRawSignal?(payload: Record<string, unknown>, options?: Record<string, unknown>): unknown | Promise<unknown>;
  observePlatformCandidates?(payload: Record<string, unknown>, options?: Record<string, unknown>): unknown | Promise<unknown>;
  observeCandidates?(candidates: Array<Record<string, unknown>>, options?: Record<string, unknown>): unknown | Promise<unknown>;
  observe?(payload: Record<string, unknown>, options?: Record<string, unknown>): unknown | Promise<unknown>;
  insertAnnotation?(platform: string, payload: Record<string, unknown>, options?: Record<string, unknown>): unknown | Promise<unknown>;
  insertMark?(payload: Record<string, unknown>, options?: Record<string, unknown>): unknown | Promise<unknown>;
  speakerTrack?(platform: string, payload: Record<string, unknown>, options?: Record<string, unknown>): unknown | Promise<unknown>;
  participantTrack?(platform: string, payload: Record<string, unknown>, options?: Record<string, unknown>): unknown | Promise<unknown>;
  ingestProvider?(platform: string, payload: Record<string, unknown>, options?: Record<string, unknown>): unknown | Promise<unknown>;
}

export interface MeetingPlatformAdapterSessionOptions {
  client?: MeetingPlatformAdapterSessionClient;
  sdk?: MeetingPlatformAdapterSessionClient;
  launchPlan?: MeetingPlatformAdapterLaunchPlan;
  launch_plan?: MeetingPlatformAdapterLaunchPlan;
  sessionId?: string;
  session_id?: string;
  source?: string;
  capturedAtMs?: number;
  captured_at_ms?: number;
  clock?: () => number;
  now?: () => number;
  platformArgument?: boolean;
  platform_argument?: boolean;
  allowUnobservedAxis?: boolean;
  allow_unobserved_axis?: boolean;
  readAdapterSelection?: boolean;
  read_adapter_selection?: boolean;
  allowUnreadAdapterSelection?: boolean;
  allow_unread_adapter_selection?: boolean;
  validateRawSignal?: boolean;
  validate_raw_signal?: boolean;
  allowUnvalidatedRawSignal?: boolean;
  allow_unvalidated_raw_signal?: boolean;
  [key: string]: unknown;
}

export interface MeetingPlatformAdapterSessionEvent {
  type: 'meeting_platform_adapter_session_event';
  schema: 'meeting_platform_adapter_session_event';
  schema_version: 1;
  session_id: string;
  action: string;
  platform?: string;
  selected_surface?: string;
  payload: Record<string, unknown>;
  result?: unknown;
  captured_at_ms?: number;
}

export interface MeetingPlatformAdapterSession {
  type: 'meeting_platform_adapter_session';
  schema: 'meeting_platform_adapter_session';
  schema_version: 1;
  id: string;
  platform?: string;
  selected_surface?: string;
  adapter_selection?: Record<string, unknown>;
  launch_plan: MeetingPlatformAdapterLaunchPlan;
  axis_contract: Record<string, unknown>;
  runtime_actions: Array<Record<string, unknown>>;
  getState(): Record<string, unknown>;
  readAdapterSelection(input?: Record<string, unknown>, options?: MeetingPlatformAdapterSessionOptions): Promise<MeetingPlatformAdapterSessionEvent>;
  validateRawSignal(input?: Record<string, unknown>, options?: MeetingPlatformAdapterSessionOptions): Promise<MeetingPlatformAdapterSessionEvent>;
  observeAxis(input?: Record<string, unknown>, options?: MeetingPlatformAdapterSessionOptions): Promise<MeetingPlatformAdapterSessionEvent>;
  insertAnnotation(input?: Record<string, unknown>, options?: MeetingPlatformAdapterSessionOptions): Promise<MeetingPlatformAdapterSessionEvent>;
  insertMark(input?: Record<string, unknown>, options?: MeetingPlatformAdapterSessionOptions): Promise<MeetingPlatformAdapterSessionEvent>;
  speakerTrack(input?: Record<string, unknown>, options?: MeetingPlatformAdapterSessionOptions): Promise<MeetingPlatformAdapterSessionEvent>;
  participantTrack(input?: Record<string, unknown>, options?: MeetingPlatformAdapterSessionOptions): Promise<MeetingPlatformAdapterSessionEvent>;
  providerReconcile(input?: Record<string, unknown>, options?: MeetingPlatformAdapterSessionOptions): Promise<MeetingPlatformAdapterSessionEvent>;
}

export interface MeetingPlatformAdapterSessionHandoff {
  type: 'meeting_platform_adapter_session_handoff';
  schema: 'meeting_platform_adapter_session_handoff';
  schema_version: 1;
  platform?: string;
  selected_surface?: string;
  session_factory: 'createMeetingPlatformAdapterSession';
  required_client_methods: string[];
  optional_client_methods: string[];
  launch_plan_schema: string;
  runtime_actions: Array<Record<string, unknown>>;
  axis_contract: Record<string, unknown>;
  next_actions: string[];
}

export function createMeetingPlatformAdapterSession(
  launchPlanOrInput?: MeetingPlatformAdapterLaunchPlan | Record<string, unknown>,
  clientOrOptions?: MeetingPlatformAdapterSessionClient | MeetingPlatformAdapterSessionOptions,
  options?: MeetingPlatformAdapterSessionOptions,
): MeetingPlatformAdapterSession;

export function buildMeetingPlatformAdapterSessionHandoff(
  launchPlanOrInput?: MeetingPlatformAdapterLaunchPlan | Record<string, unknown>,
  options?: MeetingPlatformAdapterSessionOptions,
): MeetingPlatformAdapterSessionHandoff;
