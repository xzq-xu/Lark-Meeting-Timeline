import type { MeetingPlatformAdapterRoute } from './platform-adapter-route.mjs';
import type { MeetingPlatformAdaptationStrategy } from './platform-strategy.mjs';

export const MEETING_PLATFORM_ADAPTER_SELECTION_SCHEMA: 'meeting_platform_adapter_selection';
export const MEETING_PLATFORM_ADAPTER_SELECTION_MATRIX_SCHEMA: 'meeting_platform_adapter_selection_matrix';
export const MEETING_PLATFORM_ADAPTER_SELECTION_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdapterSelectionSource {
  id: string;
  priority: number;
  selected: boolean;
  role?: string;
  surface?: string;
  source?: string;
  required_for_realtime: boolean;
  required_for_production?: boolean;
  blocks_realtime_if_missing: boolean;
  timestamp_field?: string;
  evidence_input?: string;
  start_events?: string[];
  end_events?: string[];
  participant_events?: string[];
  lifecycle_events?: string[];
  [key: string]: unknown;
}

export interface MeetingPlatformAdapterSelection {
  type: 'meeting_platform_adapter_selection';
  schema: 'meeting_platform_adapter_selection';
  schema_version: 1;
  platform: string;
  display_name?: string;
  recommended_mode?: string;
  objective: string;
  selection: {
    axis_source?: string;
    axis_surface?: string;
    axis_source_role?: string;
    annotation_source: string;
    timestamp_field: string;
    provider_reconcile_source?: string;
    provider_reconcile_required_for_production: boolean;
    speaker_track_source?: string;
    post_meeting_artifact_source?: string;
  };
  runtime_policy: {
    provider_events_block_realtime: boolean;
    transcript_blocks_realtime: boolean;
    annotations_use_absolute_captured_at_ms: boolean;
    per_meeting_annotation_isolation_required: boolean;
    source_priority?: string[];
    route_order?: string[];
    runtime_actions: string[];
    startup_order: string[];
  };
  sources: MeetingPlatformAdapterSelectionSource[];
  current_evidence: {
    has_meeting_app_record_set: boolean;
    has_provider_records: boolean;
    has_local_detector_records: boolean;
    realtime_axis_evidence: boolean;
    production_evidence: boolean;
  };
  readiness: {
    selection_ready: boolean;
    route_ready: boolean;
    pilot_evidence_ready: boolean;
    production_evidence_ready: boolean;
    provider_reconcile_required_for_production: boolean;
    missing: string[];
    issues: Record<string, unknown>[];
  };
  gates: Record<string, unknown>;
  risk_profile?: Record<string, unknown>;
  references?: Record<string, unknown>;
  next_actions: string[];
  route: MeetingPlatformAdapterRoute;
  strategy?: MeetingPlatformAdaptationStrategy;
}

export interface MeetingPlatformAdapterSelectionMatrix {
  type: 'meeting_platform_adapter_selection_matrix';
  schema: 'meeting_platform_adapter_selection_matrix';
  schema_version: 1;
  platform_count: number;
  selection_ready_count: number;
  pilot_evidence_ready_count: number;
  production_evidence_ready_count: number;
  local_axis_selected_count: number;
  provider_reconcile_count: number;
  post_meeting_artifact_count: number;
  platforms: string[];
  rows: Record<string, unknown>[];
  selections: MeetingPlatformAdapterSelection[];
}

export function buildMeetingPlatformAdapterSelection(
  platform: string,
  input?: Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingPlatformAdapterSelection;

export function buildMeetingPlatformAdapterSelectionMatrix(
  input?: Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingPlatformAdapterSelectionMatrix;
