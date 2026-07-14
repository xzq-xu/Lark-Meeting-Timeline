export const PLATFORM_FIXTURE_SIGNAL_TYPES: readonly string[];
export const DEFAULT_PLATFORM_FIXTURE_SIGNAL_TYPES: Readonly<Record<string, readonly string[]>>;

export interface PlatformFixtureOptions {
  baseUrl?: string;
  base_url?: string;
  env?: Record<string, unknown>;
  platforms?: string[];
  signalTypes?: string[];
  signal_types?: string[];
  startMs?: number;
  start_ms?: number;
  durationMs?: number;
  duration_ms?: number;
  participantOffsetMs?: number;
  participant_offset_ms?: number;
  participantLeftOffsetMs?: number;
  participant_left_offset_ms?: number;
  speakerOffsetMs?: number;
  speaker_offset_ms?: number;
  artifactOffsetMs?: number;
  artifact_offset_ms?: number;
  lifecycleOffsetMs?: number;
  lifecycle_offset_ms?: number;
  title?: string;
  participantId?: string;
  participant_id?: string;
  participantName?: string;
  participant_name?: string;
  speakerId?: string;
  speaker_id?: string;
  speakerName?: string;
  speaker_name?: string;
  detectedPlatform?: string;
  detected_platform?: string;
  [key: string]: unknown;
}

export function buildPlatformFixtureEvent(platform: string, signalType: string, options?: PlatformFixtureOptions): Record<string, unknown>;
export function buildPlatformFixtureSamples(platform: string, options?: PlatformFixtureOptions): Record<string, unknown>[];
export function buildAllPlatformFixtureSamples(options?: PlatformFixtureOptions): Record<string, Record<string, unknown>[]>;
export function buildPlatformFixtureEnv(options?: PlatformFixtureOptions): Record<string, string | unknown>;
export function buildPlatformFixtureAcceptanceInput(options?: PlatformFixtureOptions): {
  baseUrl: string;
  env: Record<string, string | unknown>;
  samples: Record<string, Record<string, unknown>[]>;
};
