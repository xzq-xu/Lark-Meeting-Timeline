import type {
  MeetingPlatformAdapterDecisionInput,
} from './platform-adapter-decision.mjs';
import type {
  MeetingPlatformAdapterStartupPlan,
  MeetingPlatformAdapterStartupPlanOptions,
} from './platform-adapter-startup.mjs';
import type {
  MeetingAppDomAdaptationDiagnosis,
} from './meeting-app-profile.mjs';
import type {
  MeetingAppDomCaptureInput,
  MeetingAppDomCaptureSnapshot,
} from './meeting-app-capture.mjs';

export const MEETING_PLATFORM_ADAPTER_PREFLIGHT_SCHEMA: 'meeting_platform_adapter_preflight';
export const MEETING_PLATFORM_ADAPTER_PREFLIGHT_MATRIX_SCHEMA: 'meeting_platform_adapter_preflight_matrix';
export const MEETING_PLATFORM_ADAPTER_CANDIDATE_PREFLIGHT_SCHEMA: 'meeting_platform_adapter_candidate_preflight';
export const MEETING_PLATFORM_ADAPTER_PREFLIGHT_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdapterPreflightOptions extends MeetingPlatformAdapterStartupPlanOptions {
  preflightId?: string;
  preflight_id?: string;
  requireSpeakerTrack?: boolean;
  require_speaker_track?: boolean;
  requireCompleteLifecycle?: boolean;
  require_complete_lifecycle?: boolean;
  requireMeetingEnd?: boolean;
  require_meeting_end?: boolean;
  includeCapturedSnapshot?: boolean;
  include_captured_snapshot?: boolean;
  captureOptions?: Record<string, unknown>;
  capture_options?: Record<string, unknown>;
  captureSource?: string;
  capture_source?: string;
  [key: string]: unknown;
}

export interface MeetingPlatformAdapterPreflightReadiness {
  static_startup_ready: boolean;
  live_evidence_ready: boolean;
  meeting_start_ready: boolean;
  meeting_end_ready: boolean;
  speaker_track_ready: boolean;
  speaker_track_required: boolean;
  complete_lifecycle_required: boolean;
  provider_reconcile_required_for_realtime: boolean;
  transcript_blocks_realtime: boolean;
  realtime_annotation_ready: boolean;
  production_lifecycle_ready: boolean;
}

export interface MeetingPlatformAdapterPreflight {
  type: 'meeting_platform_adapter_preflight';
  schema: 'meeting_platform_adapter_preflight';
  schema_version: 1;
  id?: string;
  accepted: boolean;
  status: string;
  platform?: string;
  display_name?: string;
  input?: Record<string, unknown>;
  startup?: MeetingPlatformAdapterStartupPlan;
  dom_diagnosis?: MeetingAppDomAdaptationDiagnosis | Record<string, unknown>;
  native_diagnosis?: Record<string, unknown>;
  readiness: MeetingPlatformAdapterPreflightReadiness;
  summary?: Record<string, unknown>;
  current_window?: Record<string, unknown>;
  capture?: Record<string, unknown>;
  captured_snapshot?: MeetingAppDomCaptureSnapshot;
  issues?: Array<Record<string, unknown>>;
  next_actions: string[];
}

export interface MeetingPlatformAdapterPreflightMatrix {
  type: 'meeting_platform_adapter_preflight_matrix';
  schema: 'meeting_platform_adapter_preflight_matrix';
  schema_version: 1;
  platform_count: number;
  accepted_count: number;
  realtime_ready_count: number;
  live_evidence_ready_count: number;
  meeting_start_ready_count: number;
  meeting_end_ready_count: number;
  speaker_track_ready_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  preflights: MeetingPlatformAdapterPreflight[];
  next_actions: string[];
}

export interface MeetingPlatformAdapterCandidatePreflight {
  type: 'meeting_platform_adapter_candidate_preflight';
  schema: 'meeting_platform_adapter_candidate_preflight';
  schema_version: 1;
  selection_strategy: string;
  accepted: boolean;
  status: string;
  candidate_count: number;
  supported_candidate_count: number;
  accepted_count: number;
  realtime_ready_count: number;
  live_evidence_ready_count: number;
  meeting_start_ready_count: number;
  meeting_end_ready_count: number;
  speaker_track_ready_count: number;
  selected_candidate_index?: number;
  selected_candidate_score?: number;
  selected_platform?: string;
  selected_status?: string;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  preflights: MeetingPlatformAdapterPreflight[];
  next_actions: string[];
}

export function buildMeetingPlatformAdapterPreflight(
  input?: string | URL | MeetingPlatformAdapterDecisionInput,
  options?: MeetingPlatformAdapterPreflightOptions,
): MeetingPlatformAdapterPreflight;

export function buildMeetingPlatformAdapterCurrentWindowPreflight(
  input?: MeetingAppDomCaptureInput | Document | MeetingPlatformAdapterDecisionInput,
  options?: MeetingPlatformAdapterPreflightOptions,
): MeetingPlatformAdapterPreflight;

export function buildMeetingPlatformAdapterPreflightMatrix(
  input?: MeetingPlatformAdapterDecisionInput,
  options?: MeetingPlatformAdapterPreflightOptions,
): MeetingPlatformAdapterPreflightMatrix;

export function buildMeetingPlatformAdapterCandidatePreflight(
  input?: MeetingPlatformAdapterDecisionInput | MeetingPlatformAdapterDecisionInput[] | Record<string, unknown>,
  options?: MeetingPlatformAdapterPreflightOptions,
): MeetingPlatformAdapterCandidatePreflight;

export function assertMeetingPlatformAdapterCurrentWindowPreflight(
  input?: MeetingAppDomCaptureInput | Document | MeetingPlatformAdapterDecisionInput,
  options?: MeetingPlatformAdapterPreflightOptions,
): MeetingPlatformAdapterPreflight;

export function assertMeetingPlatformAdapterPreflight(
  input?: string | URL | MeetingPlatformAdapterDecisionInput,
  options?: MeetingPlatformAdapterPreflightOptions,
): MeetingPlatformAdapterPreflight;

export function assertMeetingPlatformAdapterCandidatePreflight(
  input?: MeetingPlatformAdapterDecisionInput | MeetingPlatformAdapterDecisionInput[] | Record<string, unknown>,
  options?: MeetingPlatformAdapterPreflightOptions,
): MeetingPlatformAdapterCandidatePreflight;

export function assertMeetingPlatformAdapterPreflightMatrix(
  input?: MeetingPlatformAdapterDecisionInput,
  options?: MeetingPlatformAdapterPreflightOptions,
): MeetingPlatformAdapterPreflightMatrix;
