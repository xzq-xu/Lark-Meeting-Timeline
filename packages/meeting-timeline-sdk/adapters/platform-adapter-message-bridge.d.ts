import type { MeetingPlatformAdapterInstallManifest } from './platform-adapter-install-manifest.mjs';
import type {
  MeetingPlatformAdapterCandidateLaunchPlan,
  MeetingPlatformAdapterLaunchPlan,
} from './platform-adapter-launch-plan.mjs';
import type {
  MeetingPlatformAdapterRunner,
  MeetingPlatformAdapterRunnerOptions,
} from './platform-adapter-runner.mjs';
import type { MeetingPlatformAdapterSessionClient } from './platform-adapter-session.mjs';

export const MEETING_PLATFORM_ADAPTER_MESSAGE_BRIDGE_SCHEMA: 'meeting_platform_adapter_message_bridge';
export const MEETING_PLATFORM_ADAPTER_MESSAGE_BRIDGE_EVENT_SCHEMA: 'meeting_platform_adapter_message_bridge_event';
export const MEETING_PLATFORM_ADAPTER_MESSAGE_BRIDGE_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdapterMessageBridgeOptions extends MeetingPlatformAdapterRunnerOptions {
  bridgeId?: string;
  bridge_id?: string;
  autoOpen?: boolean;
  auto_open?: boolean;
  autoOpenOnInsert?: boolean;
  auto_open_on_insert?: boolean;
  [key: string]: unknown;
}

export interface MeetingPlatformAdapterMessageBridgeEvent {
  type: 'meeting_platform_adapter_message_bridge_event';
  schema: 'meeting_platform_adapter_message_bridge_event';
  schema_version: 1;
  handled: boolean;
  bridge_id: string;
  action: string;
  message_type?: string;
  request_id?: string;
  platform?: string;
  selected_surface?: string;
  result?: unknown;
  runner_state?: Record<string, unknown>;
}

export interface MeetingPlatformAdapterMessageBridge {
  type: 'meeting_platform_adapter_message_bridge';
  schema: 'meeting_platform_adapter_message_bridge';
  schema_version: 1;
  id: string;
  runner: MeetingPlatformAdapterRunner;
  getState(): Record<string, unknown>;
  handleMessage(message?: Record<string, unknown>, options?: MeetingPlatformAdapterMessageBridgeOptions): Promise<MeetingPlatformAdapterMessageBridgeEvent>;
  dispatchMessage(message?: Record<string, unknown>, options?: MeetingPlatformAdapterMessageBridgeOptions): Promise<MeetingPlatformAdapterMessageBridgeEvent>;
  open(input?: Record<string, unknown> | MeetingPlatformAdapterLaunchPlan | MeetingPlatformAdapterCandidateLaunchPlan | string, options?: MeetingPlatformAdapterMessageBridgeOptions): Promise<MeetingPlatformAdapterMessageBridgeEvent>;
  insertAnnotation(input?: Record<string, unknown>, options?: MeetingPlatformAdapterMessageBridgeOptions): Promise<MeetingPlatformAdapterMessageBridgeEvent>;
  insertMark(input?: Record<string, unknown>, options?: MeetingPlatformAdapterMessageBridgeOptions): Promise<MeetingPlatformAdapterMessageBridgeEvent>;
}

export interface MeetingPlatformAdapterMessageBridgeHandoff {
  type: 'meeting_platform_adapter_message_bridge_handoff';
  schema: 'meeting_platform_adapter_message_bridge_handoff';
  schema_version: 1;
  bridge_factory: 'createMeetingPlatformAdapterMessageBridge';
  required_runner_factory: 'createMeetingPlatformAdapterRunner';
  install_manifest_schema?: string;
  message_types: string[];
  runtime_sequence: string[];
  supported_surfaces: string[];
  auto_open_on_insert: boolean;
  next_actions: string[];
}

export function createMeetingPlatformAdapterMessageBridge(
  manifestOrRunner?: MeetingPlatformAdapterInstallManifest | MeetingPlatformAdapterLaunchPlan | MeetingPlatformAdapterRunner | Record<string, unknown>,
  clientOrOptions?: MeetingPlatformAdapterSessionClient | MeetingPlatformAdapterMessageBridgeOptions,
  options?: MeetingPlatformAdapterMessageBridgeOptions,
): MeetingPlatformAdapterMessageBridge;

export function buildMeetingPlatformAdapterMessageBridgeHandoff(
  manifestOrInput?: MeetingPlatformAdapterInstallManifest | MeetingPlatformAdapterLaunchPlan | Record<string, unknown>,
  options?: MeetingPlatformAdapterMessageBridgeOptions,
): MeetingPlatformAdapterMessageBridgeHandoff;

export function assertMeetingPlatformAdapterMessageBridge(
  bridge: MeetingPlatformAdapterMessageBridge,
): MeetingPlatformAdapterMessageBridge;
