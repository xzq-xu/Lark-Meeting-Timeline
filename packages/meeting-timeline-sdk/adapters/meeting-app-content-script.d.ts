import type { MeetingTimelineClient, MeetingTimelineClientOptions } from '../index.mjs';
import type {
  MeetingAppBrowserRuntimeMessage,
  MeetingAppBrowserRuntimeOptions,
} from './meeting-app-browser-runtime.mjs';

export interface MeetingAppContentScriptBridgeOptions extends MeetingAppBrowserRuntimeOptions {
  runtime?: Record<string, unknown>;
  extensionRuntime?: Record<string, unknown>;
  extension_runtime?: Record<string, unknown>;
  chromeRuntime?: Record<string, unknown>;
  chrome_runtime?: Record<string, unknown>;
  browserRuntime?: Record<string, unknown>;
  browser_runtime?: Record<string, unknown>;
  chrome?: Record<string, unknown>;
  browser?: Record<string, unknown>;
  extensionMessaging?: boolean;
  extension_messaging?: boolean;
  windowMessaging?: boolean;
  window_messaging?: boolean;
  postWindowResponses?: boolean;
  post_window_responses?: boolean;
  allowedOrigins?: string[];
  allowed_origins?: string[];
  messagePrefixes?: string[];
  message_prefixes?: string[];
  directMessageTypes?: string[];
  direct_message_types?: string[];
  messageFilter?: (message: MeetingAppBrowserRuntimeMessage, options?: MeetingAppContentScriptBridgeOptions) => boolean;
  message_filter?: (message: MeetingAppBrowserRuntimeMessage, options?: MeetingAppContentScriptBridgeOptions) => boolean;
  messageOptions?: Record<string, unknown>;
  message_options?: Record<string, unknown>;
  startOptions?: Record<string, unknown>;
  start_options?: Record<string, unknown>;
  startRuntime?: boolean;
  start_runtime?: boolean;
}

export interface MeetingAppContentScriptBridge {
  runtime: Record<string, unknown>;
  dispatchMessage(message?: MeetingAppBrowserRuntimeMessage, messageOptions?: Record<string, unknown>): Promise<Record<string, unknown>>;
  start(options?: MeetingAppContentScriptBridgeOptions): Record<string, unknown>;
  stop(): Record<string, unknown>;
  dispose(): Record<string, unknown>;
  installExtensionMessaging(options?: MeetingAppContentScriptBridgeOptions): Record<string, unknown>;
  installWindowMessaging(options?: MeetingAppContentScriptBridgeOptions): Record<string, unknown>;
  removeMessaging(kind?: 'extension' | 'window' | string): Record<string, unknown>;
  getState(): Record<string, unknown>;
}

export function createMeetingAppContentScriptBridge(
  clientOrOptions: MeetingTimelineClient | MeetingTimelineClientOptions | MeetingAppContentScriptBridgeOptions,
  options?: MeetingAppContentScriptBridgeOptions,
): MeetingAppContentScriptBridge;

export function installMeetingAppContentScriptBridge(
  clientOrOptions: MeetingTimelineClient | MeetingTimelineClientOptions | MeetingAppContentScriptBridgeOptions,
  options?: MeetingAppContentScriptBridgeOptions,
): MeetingAppContentScriptBridge;

export default createMeetingAppContentScriptBridge;
