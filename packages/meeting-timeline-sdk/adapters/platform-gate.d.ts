export const PLATFORM_LAUNCH_GATE_EVIDENCE_LEVELS: readonly string[];

export interface PlatformLaunchGateOptions {
  baseUrl?: string;
  base_url?: string;
  env?: Record<string, unknown>;
  records?: unknown[];
  captureRecords?: unknown[];
  capture_records?: unknown[];
  samples?: unknown[] | Record<string, unknown[]>;
  sampleEvents?: unknown[] | Record<string, unknown[]>;
  sample_events?: unknown[] | Record<string, unknown[]>;
  requiredCoverage?: string[];
  required_coverage?: string[];
  requireEndEvent?: boolean;
  require_end_event?: boolean;
  requireParticipants?: boolean;
  require_participants?: boolean;
  requireSpeakerActivity?: boolean;
  require_speaker_activity?: boolean;
  requireArtifact?: boolean;
  require_artifact?: boolean;
  requireSubscriptionLifecycle?: boolean;
  require_subscription_lifecycle?: boolean;
  allowFixtureEvidence?: boolean;
  allow_fixture_evidence?: boolean;
  allowFixtureProduction?: boolean;
  allow_fixture_production?: boolean;
  requireProductionReady?: boolean;
  require_production_ready?: boolean;
  [key: string]: unknown;
}

export function buildPlatformLaunchGate(platform: string, options?: PlatformLaunchGateOptions): Record<string, unknown>;
export function buildAllPlatformLaunchGates(options?: PlatformLaunchGateOptions): Record<string, unknown>[];
export function buildMeetingPlatformLaunchGateSummary(options?: PlatformLaunchGateOptions): Record<string, unknown>;
export function assertPlatformLaunchGate(platform: string, options?: PlatformLaunchGateOptions): Record<string, unknown>;
export function assertAllPlatformLaunchGates(options?: PlatformLaunchGateOptions): Record<string, unknown>;
export function buildFixtureLaunchGateInput(options?: PlatformLaunchGateOptions): Record<string, unknown>;
