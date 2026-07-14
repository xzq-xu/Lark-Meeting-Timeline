export interface DesktopMeetingHostOptions {
  baseUrl?: string;
  base_url?: string;
  platforms?: string[] | string;
  platform?: string[] | string;
  intervalMs?: number;
  interval_ms?: number;
  startStableSamples?: number;
  endMissingSamples?: number;
  speakerStableSamples?: number;
  dryRun?: boolean;
  scanner?: (options?: Record<string, unknown>) => Promise<Record<string, unknown>> | Record<string, unknown>;
  fetch?: typeof globalThis.fetch;
  onEvent?: (event: Record<string, unknown>) => unknown | Promise<unknown>;
}

export declare const DESKTOP_MEETING_SCAN_SCHEMA: string;
export declare const DESKTOP_MEETING_HOST_SCHEMA: string;
export declare const DESKTOP_MEETING_HOST_VERSION: number;
export declare const DESKTOP_MEETING_PROFILES: Readonly<Record<string, Readonly<Record<string, unknown>>>>;

export declare function scanDesktopMeetingApps(options?: Record<string, unknown>): Promise<Record<string, unknown>>;
export declare function normalizeDesktopMeetingScan(scan?: Record<string, unknown>, options?: Record<string, unknown>): Record<string, unknown>;
export declare function buildDesktopMeetingAdapterReadiness(scanOrResult?: Record<string, unknown>, options?: Record<string, unknown>): Record<string, unknown>;
export declare function createDesktopMeetingAdapterHost(options?: DesktopMeetingHostOptions): {
  type: string;
  platforms: string[];
  scan(options?: Record<string, unknown>): Promise<Record<string, unknown>> | Record<string, unknown>;
  tick(options?: Record<string, unknown>): Promise<Record<string, unknown>>;
  start(): Promise<Record<string, unknown>>;
  stop(options?: Record<string, unknown>): Promise<Record<string, unknown>>;
  getState(): Record<string, unknown>;
  readiness(scan?: Record<string, unknown>): Record<string, unknown>;
};

export default createDesktopMeetingAdapterHost;
