import type { PlatformEventDiagnosticResult } from './platform-ingest.mjs';
import type { PlatformSetupOptions } from './platform-setup.mjs';

export const PLATFORM_ACCEPTANCE_COVERAGE_KEYS: readonly string[];

export interface PlatformAcceptanceSample {
  label?: string;
  name?: string;
  id?: string;
  platform?: string;
  provider?: string;
  adapter?: string;
  body?: unknown;
  payload?: unknown;
  raw_event?: unknown;
  rawEvent?: unknown;
  raw?: unknown;
  receivedAtMs?: number | string | Date;
  received_at_ms?: number | string | Date;
  received_at?: number | string | Date;
  timestamp?: number | string | Date;
  ts?: number | string | Date;
  options?: Record<string, unknown>;
  [key: string]: unknown;
}

export type PlatformAcceptanceSamples =
  | PlatformAcceptanceSample[]
  | Record<string, PlatformAcceptanceSample | PlatformAcceptanceSample[]>;

export interface PlatformAcceptanceOptions extends PlatformSetupOptions {
  samples?: PlatformAcceptanceSamples;
  sampleEvents?: PlatformAcceptanceSamples;
  sample_events?: PlatformAcceptanceSamples;
  requiredCoverage?: string[];
  required_coverage?: string[];
  requireEndEvent?: boolean;
  require_end_event?: boolean;
  diagnosticOptions?: Record<string, unknown>;
  diagnostic_options?: Record<string, unknown>;
}

export interface PlatformAcceptanceIssue {
  severity: 'info' | 'warning' | 'error';
  code: string;
  message: string;
  sample?: string;
  [key: string]: unknown;
}

export interface PlatformAcceptanceSampleReport {
  label: string;
  diagnostic: PlatformEventDiagnosticResult;
}

export interface PlatformAcceptanceReport {
  platform: string;
  display_name?: string;
  status: 'accepted' | 'blocked' | 'pending_samples' | 'samples_failed' | 'missing_required_coverage';
  accepted: boolean;
  recommended_mode?: string;
  source_priority?: string[];
  readiness?: Record<string, unknown>;
  subscription_maintenance?: Record<string, unknown>;
  required_coverage: string[];
  missing_required_coverage: string[];
  coverage: Record<string, boolean>;
  signal_types: string[];
  sample_count: number;
  actionable_sample_count: number;
  samples: PlatformAcceptanceSampleReport[];
  issues: PlatformAcceptanceIssue[];
  next_actions: string[];
}

export interface MeetingPlatformAcceptanceSummary {
  ok: boolean;
  accepted_count: number;
  blocked_count: number;
  pending_samples_count: number;
  missing_required_coverage_count: number;
  reports: PlatformAcceptanceReport[];
}

export function buildPlatformAcceptanceReport(
  platform: string,
  options?: PlatformAcceptanceOptions,
): PlatformAcceptanceReport;

export function buildAllPlatformAcceptanceReports(
  options?: PlatformAcceptanceOptions,
): PlatformAcceptanceReport[];

export function buildMeetingPlatformAcceptanceSummary(
  options?: PlatformAcceptanceOptions,
): MeetingPlatformAcceptanceSummary;
