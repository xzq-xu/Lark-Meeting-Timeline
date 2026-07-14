import type { NormalizedMeetingIdentity, NormalizedMeetingSignal } from './core.mjs';

export const PLATFORM_ARTIFACT_IMPORTERS: Readonly<Record<string, Readonly<Record<string, {
  normalizer?: string | null;
  fetch_strategy: string;
  content_hint: string;
}>>>>;

export interface ArtifactImportPlan {
  status: 'requires_provider_fetch' | 'metadata_ready' | 'metadata_only' | 'unsupported_artifact' | 'ignored';
  action: 'fetch_and_import_transcript' | 'store_recording_artifact' | 'fetch_and_store_smart_notes' | 'store_artifact_metadata';
  platform?: string;
  meeting?: NormalizedMeetingIdentity;
  source_event_id?: string;
  occurred_at_ms?: number;
  artifact_kind: string;
  artifact_id?: string;
  artifact_url?: string;
  fetch?: {
    strategy?: string;
    resource?: string;
    url?: string;
    auth?: string;
  };
  transcript_import?: {
    endpoint: string;
    module: string;
    normalizer?: string | null;
    sdk_method: string;
  };
  content_hint?: string;
  issues: Array<{
    severity: 'info' | 'warning' | 'error';
    code: string;
    message: string;
    [key: string]: unknown;
  }>;
}

export interface ArtifactImportPlanOptions {
  importEndpoint?: string;
  import_endpoint?: string;
  includeIgnored?: boolean;
  include_ignored?: boolean;
}

export function buildArtifactImportPlan(
  signal?: NormalizedMeetingSignal | Record<string, unknown>,
  options?: ArtifactImportPlanOptions,
): ArtifactImportPlan;

export function buildArtifactImportPlans(
  input?: Array<NormalizedMeetingSignal | Record<string, unknown>> | {
    signals?: Array<NormalizedMeetingSignal | Record<string, unknown>>;
    rawSignals?: Array<NormalizedMeetingSignal | Record<string, unknown>>;
    raw_signals?: Array<NormalizedMeetingSignal | Record<string, unknown>>;
    diagnostic?: {
      signals?: Array<NormalizedMeetingSignal | Record<string, unknown>>;
      raw_signals?: Array<NormalizedMeetingSignal | Record<string, unknown>>;
    };
  } | NormalizedMeetingSignal | Record<string, unknown>,
  options?: ArtifactImportPlanOptions,
): ArtifactImportPlan[];
