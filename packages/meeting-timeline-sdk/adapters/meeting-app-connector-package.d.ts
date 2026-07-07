import type { MeetingAppTimelineConnectorPackage } from '../index.mjs';

export const MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA: 'meeting_app_timeline_connector_package';
export const MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_ACCEPTANCE_SCHEMA: 'meeting_app_timeline_connector_package_acceptance_report';
export const MEETING_APP_TIMELINE_CONNECTOR_HANDOFF_SCHEMA: 'meeting_app_timeline_connector_handoff';
export const MEETING_APP_TIMELINE_CONNECTOR_HOST_INSTALL_CHECKLIST_SCHEMA: 'meeting_app_timeline_connector_host_install_checklist';
export const MEETING_APP_TIMELINE_CONNECTOR_HOST_INSTALL_CHECKLIST_ACCEPTANCE_SCHEMA: 'meeting_app_timeline_connector_host_install_checklist_acceptance_report';
export const MEETING_APP_TIMELINE_CONNECTOR_ADOPTION_INDEX_SCHEMA: 'meeting_app_timeline_connector_adoption_index';
export const MEETING_APP_TIMELINE_CONNECTOR_FIELD_INTAKE_INDEX_SCHEMA: 'meeting_app_timeline_connector_field_intake_index';
export const MEETING_APP_TIMELINE_CONNECTOR_BRIDGE_HANDOFF_SCHEMA: 'meeting_app_timeline_connector_bridge_handoff';
export const MEETING_APP_TIMELINE_CONNECTOR_BRIDGE_HANDOFF_ACCEPTANCE_SCHEMA: 'meeting_app_timeline_connector_bridge_handoff_acceptance_report';
export const MEETING_APP_TIMELINE_CONNECTOR_BRIDGE_SMOKE_REPORT_SCHEMA: 'meeting_app_timeline_connector_bridge_smoke_report';
export const MEETING_APP_TIMELINE_CONNECTOR_SMOKE_PLAN_SCHEMA: 'meeting_app_timeline_connector_smoke_plan';
export const MEETING_APP_TIMELINE_CONNECTOR_SMOKE_PLAN_ACCEPTANCE_SCHEMA: 'meeting_app_timeline_connector_smoke_plan_acceptance_report';
export const MEETING_APP_TIMELINE_CONNECTOR_SMOKE_RUN_REPORT_SCHEMA: 'meeting_app_timeline_connector_smoke_run_report';
export const MEETING_APP_TIMELINE_CONNECTOR_RELEASE_GATE_SCHEMA: 'meeting_app_timeline_connector_release_gate';
export const MEETING_APP_TIMELINE_CONNECTOR_PLATFORM_ROADMAP_SCHEMA: 'meeting_app_timeline_connector_platform_roadmap';
export const MEETING_APP_TIMELINE_CONNECTOR_ADAPTER_MATRIX_SCHEMA: 'meeting_app_timeline_connector_adapter_matrix';
export const MEETING_APP_TIMELINE_CONNECTOR_ADAPTER_MATRIX_ACCEPTANCE_SCHEMA: 'meeting_app_timeline_connector_adapter_matrix_acceptance_report';
export const MEETING_APP_TIMELINE_CONNECTOR_RUNTIME_CLIENT_SCHEMA: 'meeting_app_timeline_connector_runtime_client';
export const MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA_VERSION: 1;

export interface MeetingAppTimelineConnectorPackageIssue {
  code: string;
  message: string;
  [key: string]: unknown;
}

export interface MeetingAppTimelineConnectorPackageSurfaceReport {
  surface: string;
  accepted: boolean;
  handoff_count: number;
  ready_count: number;
  realtime_annotation_ready_count: number;
  speaker_track_ready_count: number;
  participant_track_ready_count: number;
  observer_accepted_count: number;
  scheduler_sdk_ready_count: number;
  track_enabled_count: number;
  install_targets: string[];
  start_modes: string[];
  runtime_factories: string[];
  observer_factories: string[];
  issues: MeetingAppTimelineConnectorPackageIssue[];
}

export interface MeetingAppTimelineConnectorPackageAcceptanceReport {
  type: 'meeting_app_timeline_connector_package_acceptance_report';
  schema: 'meeting_app_timeline_connector_package_acceptance_report';
  schema_version: 1;
  accepted: boolean;
  target: string;
  package_id?: string;
  platform_count: number;
  surface_count: number;
  platforms: string[];
  surfaces: string[];
  require_tracks: boolean;
  missing_runtime_actions: Array<{ platform: string; action: string }>;
  surface_reports: MeetingAppTimelineConnectorPackageSurfaceReport[];
  extension: {
    required: boolean;
    accepted: boolean;
    file_count?: number;
    manifest_version?: number;
    issues: MeetingAppTimelineConnectorPackageIssue[];
  };
  issue_count: number;
  issues: MeetingAppTimelineConnectorPackageIssue[];
  next_actions: string[];
}

export interface MeetingAppTimelineConnectorHandoff {
  type: 'meeting_app_timeline_connector_handoff';
  schema: 'meeting_app_timeline_connector_handoff';
  schema_version: 1;
  package_id?: string;
  accepted: boolean;
  target: string;
  base_url?: string;
  platforms: string[];
  surfaces: string[];
  platform_count: number;
  surface_count: number;
  runtime_event_endpoint?: string;
  provider_replay?: {
    accepted?: boolean;
    accepted_count?: number;
    platform_count?: number;
    runtime_event_count?: number;
    command?: string;
    bin?: string;
    sdk_method?: string;
    rows?: Array<Record<string, unknown>>;
    provider_events_block_realtime?: boolean;
  };
  adapter_blueprints?: {
    ready_count?: number;
    platform_count?: number;
    sdk_method?: string;
    command?: string;
    rows?: Array<Record<string, unknown>>;
  };
  startup_plans?: {
    realtime_startup_ready_count?: number;
    platform_count?: number;
    sdk_method?: string;
    matrix_sdk_method?: string;
    rows?: Array<Record<string, unknown>>;
  };
  timestamp_field?: string;
  provider_events_block_realtime?: boolean;
  transcript_blocks_realtime?: boolean;
  host_entrypoints?: Array<Record<string, unknown>>;
  package_entrypoints?: Array<Record<string, unknown>>;
  ci_gates?: string[];
  rollout_checklist?: string[];
  surface_matrix: Array<Record<string, unknown>>;
  extension?: Record<string, unknown>;
  acceptance: MeetingAppTimelineConnectorPackageAcceptanceReport;
  next_actions: string[];
}

export interface MeetingAppTimelineConnectorHostInstallChecklist {
  type: 'meeting_app_timeline_connector_host_install_checklist';
  schema: 'meeting_app_timeline_connector_host_install_checklist';
  schema_version: 1;
  package_id?: string;
  accepted: boolean;
  target: string;
  base_url?: string;
  runtime_event_endpoint?: string;
  provider_replay?: {
    accepted?: boolean;
    accepted_count?: number;
    platform_count?: number;
    runtime_event_count?: number;
    command?: string;
    bin?: string;
    sdk_method?: string;
    rows?: Array<Record<string, unknown>>;
    provider_events_block_realtime?: boolean;
  };
  platform_count: number;
  surface_count: number;
  ready_count: number;
  timestamp_field?: string;
  contracts: Record<string, unknown>;
  files_to_read_first: string[];
  rows: Array<Record<string, unknown>>;
  acceptance: MeetingAppTimelineConnectorPackageAcceptanceReport;
  next_actions: string[];
}

export interface MeetingAppTimelineConnectorHostInstallChecklistAcceptanceReport {
  type: 'meeting_app_timeline_connector_host_install_checklist_acceptance_report';
  schema: 'meeting_app_timeline_connector_host_install_checklist_acceptance_report';
  schema_version: 1;
  accepted: boolean;
  target?: string;
  package_id?: string;
  platform_count: number;
  row_count: number;
  ready_count: number;
  timestamp_field?: string;
  runtime_event_endpoint?: string;
  checklist_accepted: boolean;
  issue_count: number;
  issues: MeetingAppTimelineConnectorPackageIssue[];
  rows: Array<Record<string, unknown>>;
  next_actions: string[];
}

export interface MeetingAppTimelineConnectorFieldIntakeIndex {
  type: 'meeting_app_timeline_connector_field_intake_index';
  schema: 'meeting_app_timeline_connector_field_intake_index';
  schema_version: 1;
  accepted: boolean;
  target?: string;
  package_id?: string;
  base_url?: string;
  platform_count: number;
  row_count: number;
  production_ready_count: number;
  realtime_ready_count: number;
  provider_blocked_count: number;
  local_capture_needed_count: number;
  provider_capture_needed_count: number;
  field_matrix_schema?: string;
  rows: Array<Record<string, unknown> & {
    platform: string;
    status: string;
    production_ready: boolean;
    ready_for_realtime_annotations: boolean;
    connector_selected_surface?: string;
    connector_install_target?: string;
    provider_endpoint?: string;
    field_evidence_input?: string;
    evidence_package?: string;
    required_local_snapshots: string[];
    required_provider_coverage: string[];
    missing_env: string[];
    commands?: Record<string, string>;
    operator_steps: Array<Record<string, unknown>>;
    next_actions: string[];
  }>;
  issue_count: number;
  issues: string[];
  next_actions: string[];
}

export interface MeetingAppTimelineConnectorAdoptionIndex {
  type: 'meeting_app_timeline_connector_adoption_index';
  schema: 'meeting_app_timeline_connector_adoption_index';
  schema_version: 1;
  accepted: boolean;
  target?: string;
  package_id?: string;
  runtime_event_endpoint?: string;
  timestamp_field?: string;
  platform_count: number;
  row_count: number;
  realtime_ready_count: number;
  bridge_ready_count: number;
  production_evidence_ready_count: number;
  production_evidence_pending_count: number;
  source_schemas: Record<string, unknown>;
  files_to_read_first: string[];
  rows: Array<Record<string, unknown> & {
    platform: string;
    status: string;
    selected_surface?: string;
    install_target?: string;
    runtime_preset?: string;
    realtime_ready: boolean;
    bridge_ready: boolean;
    smoke_plan_ready?: boolean;
    can_start_axis_before_provider?: boolean;
    can_insert_annotation_on_current_axis?: boolean;
    production_evidence_accepted: boolean;
    production_evidence_required: string[];
    runtime_actions: string[];
    missing: string[];
    next_actions: string[];
  }>;
  issue_count: number;
  issues: string[];
  next_actions: string[];
}

export interface MeetingAppTimelineConnectorBridgeHandoff {
  type: 'meeting_app_timeline_connector_bridge_handoff';
  schema: 'meeting_app_timeline_connector_bridge_handoff';
  schema_version: 1;
  accepted: boolean;
  target?: string;
  package_id?: string;
  base_url?: string;
  runtime_event_endpoint?: string;
  platform_count: number;
  row_count: number;
  platforms: string[];
  module?: string;
  factories?: Record<string, string>;
  kit_methods?: string[];
  message_contract?: Record<string, unknown>;
  host_requirements?: Record<string, unknown>;
  startup_order: string[];
  sample_messages: Array<Record<string, unknown>>;
  rows: Array<Record<string, unknown>>;
  issue_count: number;
  issues: string[];
  source_schemas: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingAppTimelineConnectorBridgeHandoffAcceptanceReport {
  type: 'meeting_app_timeline_connector_bridge_handoff_acceptance_report';
  schema: 'meeting_app_timeline_connector_bridge_handoff_acceptance_report';
  schema_version: 1;
  accepted: boolean;
  target?: string;
  package_id?: string;
  platform_count: number;
  row_count: number;
  runtime_event_endpoint?: string;
  timestamp_field?: string;
  module?: string;
  install_content_script_bridge_factory?: string;
  create_hub_factory?: string;
  required_message_types: string[];
  required_output_runtime_actions: string[];
  bridge_handoff_accepted: boolean;
  issue_count: number;
  issues: MeetingAppTimelineConnectorPackageIssue[];
  rows: Array<Record<string, unknown>>;
  next_actions: string[];
}

export interface MeetingAppTimelineConnectorBridgeSmokeReport {
  type: 'meeting_app_timeline_connector_bridge_smoke_report';
  schema: 'meeting_app_timeline_connector_bridge_smoke_report';
  schema_version: 1;
  accepted: boolean;
  bridge_handoff_accepted: boolean;
  bridge_handoff_acceptance_accepted: boolean;
  dry_run: boolean;
  target?: string;
  package_id?: string;
  platform: string;
  platform_count?: number;
  runtime_event_endpoint?: string;
  timestamp_field?: string;
  step_count: number;
  accepted_step_count: number;
  runtime_event_count: number;
  runtime_event_actions: string[];
  observe_before_insert: boolean;
  steps: Array<Record<string, unknown>>;
  calls: Array<Record<string, unknown>>;
  issue_count: number;
  issues: string[];
  next_actions: string[];
}

export interface MeetingAppTimelineConnectorSmokePlanStep {
  id: string;
  order: number;
  action: string;
  client_method: string;
  required: boolean;
  input: Record<string, unknown>;
  expected: Record<string, unknown>;
}

export interface MeetingAppTimelineConnectorSmokePlan {
  type: 'meeting_app_timeline_connector_smoke_plan';
  schema: 'meeting_app_timeline_connector_smoke_plan';
  schema_version: 1;
  package_id?: string;
  accepted: boolean;
  target?: string;
  runtime_event_endpoint?: string;
  platform_count: number;
  row_count: number;
  timestamp_field: 'captured_at_ms';
  source_checklist_schema?: string;
  rows: Array<Record<string, unknown> & {
    platform: string;
    selected_surface?: string;
    install_target?: string;
    realtime_startup_ready?: boolean;
    adapter_blueprint_ready?: boolean;
    required_action_count: number;
    optional_action_count: number;
    steps: MeetingAppTimelineConnectorSmokePlanStep[];
  }>;
  next_actions: string[];
}

export interface MeetingAppTimelineConnectorSmokePlanAcceptanceReport {
  type: 'meeting_app_timeline_connector_smoke_plan_acceptance_report';
  schema: 'meeting_app_timeline_connector_smoke_plan_acceptance_report';
  schema_version: 1;
  accepted: boolean;
  target?: string;
  package_id?: string;
  platform_count: number;
  row_count: number;
  required_step_count: number;
  optional_step_count: number;
  timestamp_field?: string;
  runtime_event_endpoint?: string;
  smoke_plan_accepted: boolean;
  issue_count: number;
  issues: MeetingAppTimelineConnectorPackageIssue[];
  rows: Array<Record<string, unknown>>;
  next_actions: string[];
}

export interface MeetingAppTimelineConnectorSmokeRunStepResult {
  id: string;
  order: number;
  action: string;
  client_method?: string;
  required: boolean;
  platform?: string;
  captured_at_ms?: number;
  accepted: boolean;
  result?: Record<string, unknown>;
  error?: Record<string, unknown>;
}

export interface MeetingAppTimelineConnectorSmokeRunReport {
  type: 'meeting_app_timeline_connector_smoke_run_report';
  schema: 'meeting_app_timeline_connector_smoke_run_report';
  schema_version: 1;
  accepted: boolean;
  dry_run: boolean;
  plan_accepted: boolean;
  target?: string;
  package_id?: string;
  platform_count: number;
  row_count: number;
  timestamp_field?: string;
  runtime_event_endpoint?: string;
  required_step_count: number;
  optional_step_count: number;
  executed_step_count: number;
  failed_step_count: number;
  call_count: number;
  calls: Array<Record<string, unknown>>;
  issue_count: number;
  issues: string[];
  rows: Array<Record<string, unknown> & {
    platform: string;
    selected_surface?: string;
    install_target?: string;
    accepted: boolean;
    required_step_count: number;
    optional_step_count: number;
    executed_step_count: number;
    failed_step_count: number;
    observe_before_insert?: boolean;
    captured_at_ms_preserved?: boolean;
    steps: MeetingAppTimelineConnectorSmokeRunStepResult[];
    issues: string[];
  }>;
  next_actions: string[];
}

export interface MeetingAppTimelineConnectorReleaseGate {
  type: 'meeting_app_timeline_connector_release_gate';
  schema: 'meeting_app_timeline_connector_release_gate';
  schema_version: 1;
  accepted: boolean;
  target: 'pilot' | 'production';
  package_id?: string;
  runtime_event_endpoint?: string;
  timestamp_field?: string;
  platform_count: number;
  row_count: number;
  pilot_ready_count: number;
  production_ready_count: number;
  realtime_ready_count: number;
  bridge_ready_count: number;
  field_intake_ready_count: number;
  require_smoke_reports: boolean;
  source_schemas: Record<string, unknown>;
  files_to_read_first: string[];
  required_gates: Array<Record<string, unknown> & {
    id: string;
    accepted: boolean;
    supplied: boolean;
    required: boolean;
    severity: string;
    status: string;
    source_schema?: string;
    issue_count: number;
    issues: string[];
  }>;
  rows: Array<Record<string, unknown> & {
    platform: string;
    pilot_ready: boolean;
    production_ready: boolean;
    realtime_ready: boolean;
    bridge_ready: boolean;
    field_intake_ready: boolean;
    smoke_run_ready: boolean;
    production_evidence_accepted: boolean;
    selected_surface?: string;
    install_target?: string;
    status: string;
    missing: string[];
    next_actions: string[];
  }>;
  issue_count: number;
  issues: MeetingAppTimelineConnectorPackageIssue[];
  next_actions: string[];
}

export interface MeetingAppTimelineConnectorPlatformRoadmap {
  type: 'meeting_app_timeline_connector_platform_roadmap';
  schema: 'meeting_app_timeline_connector_platform_roadmap';
  schema_version: 1;
  accepted: boolean;
  target?: string;
  package_id?: string;
  platform_count: number;
  row_count: number;
  pilot_ready_count: number;
  production_ready_count: number;
  recommended_first_platform?: string;
  recommended_first_surface?: string;
  source_schemas: Record<string, unknown>;
  files_to_read_first: string[];
  rows: Array<Record<string, unknown> & {
    platform: string;
    display_name?: string;
    order: number;
    priority_tier?: string;
    release_status: string;
    recommended_first_surface?: string;
    selected_surface?: string;
    install_target?: string;
    surface_order?: string[];
    pilot_ready: boolean;
    production_ready: boolean;
    release_missing: string[];
    provider_path?: string;
    provider_permission_risk?: string;
    next_action?: string;
    surface_coverage: Record<string, boolean>;
    runtime_contract: Record<string, unknown>;
    validation_sequence: string[];
    sdk_facade_methods: Record<string, string>;
    reasons: string[];
    next_actions: string[];
  }>;
  issue_count: number;
  issues: string[];
  next_actions: string[];
}

export interface MeetingAppTimelineConnectorAdapterMatrix {
  type: 'meeting_app_timeline_connector_adapter_matrix';
  schema: 'meeting_app_timeline_connector_adapter_matrix';
  schema_version: 1;
  accepted: boolean;
  target?: string;
  package_id?: string;
  platform_count: number;
  row_count: number;
  host_wiring_ready_count: number;
  pilot_ready_count: number;
  production_ready_count: number;
  recommended_first_platform?: string;
  recommended_first_surface?: string;
  runtime_event_endpoint?: string;
  timestamp_field?: string;
  provider_events_block_realtime: boolean;
  transcript_blocks_realtime: boolean;
  source_schemas: Record<string, unknown>;
  files_to_read_first: string[];
  runtime_invariants: Record<string, unknown>;
  rows: Array<Record<string, unknown> & {
    platform: string;
    display_name?: string;
    status: string;
    selected_surface?: string;
    adapter_mode?: string;
    install_target?: string;
    install_step?: string;
    recommended_first_surface?: string;
    surface_order?: string[];
    can_start_axis_before_provider: boolean;
    can_insert_annotation_on_current_axis: boolean;
    provider_reconcile_blocks_realtime: boolean;
    transcript_blocks_realtime: boolean;
    provider_replay: Record<string, unknown> & {
      accepted?: boolean;
      record_count?: number;
      runtime_event_count?: number;
      signal_count?: number;
      signal_types?: string[];
      coverage?: Record<string, boolean>;
      required_coverage?: string[];
      provider_events_block_realtime?: boolean;
      file?: string;
      command?: string;
      sdk_method?: string;
    };
    runtime_event_endpoint?: string;
    timestamp_field?: string;
    input_sources: Array<Record<string, unknown>>;
    runtime_sequence: Array<Record<string, unknown>>;
    bridge_contract: Record<string, unknown>;
    sdk_facade_methods: Record<string, string>;
    evidence_contract: Record<string, unknown>;
    validation_files: string[];
    release_missing: string[];
    missing: string[];
    next_actions: string[];
  }>;
  issue_count: number;
  issues: string[];
  next_actions: string[];
}

export interface MeetingAppTimelineConnectorAdapterMatrixAcceptanceReport {
  type: 'meeting_app_timeline_connector_adapter_matrix_acceptance_report';
  schema: 'meeting_app_timeline_connector_adapter_matrix_acceptance_report';
  schema_version: 1;
  accepted: boolean;
  target?: string;
  package_id?: string;
  platform_count: number;
  row_count: number;
  host_wiring_ready_count?: number;
  pilot_ready_count?: number;
  production_ready_count?: number;
  timestamp_field?: string;
  runtime_event_endpoint?: string;
  issue_count: number;
  issues: MeetingAppTimelineConnectorPackageIssue[];
  rows: Array<Record<string, unknown>>;
  next_actions: string[];
}

export interface MeetingAppTimelineConnectorRuntimeClient {
  type: 'meeting_app_timeline_connector_runtime_client';
  schema: 'meeting_app_timeline_connector_runtime_client';
  schema_version: 1;
  package_id?: string;
  endpoint: string;
  platforms: string[];
  surfaces: string[];
  acceptance: MeetingAppTimelineConnectorPackageAcceptanceReport;
  runtime_event_client: import('./platform-runtime-event.mjs').MeetingPlatformRuntimeEventClient;
  supported_action_count: number;
  supported_actions: string[];
  supported_actions_by_platform: Record<string, string[]>;
  action_rows: Array<Record<string, unknown>>;
  supports(action: string, platform?: string): boolean;
  assertSupported(action: string, platform?: string, options?: Record<string, unknown>): Record<string, unknown>;
  buildEvent(input?: Record<string, unknown>, options?: Record<string, unknown>): import('./platform-runtime-event.mjs').MeetingPlatformRuntimeEvent;
  send(input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  observeMeetingApp(platform: string, snapshot?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  observePlatformCandidates(input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  ingestProvider(platform: string, payload?: unknown, options?: Record<string, unknown>): Promise<unknown>;
  insertAnnotation(platform: string, annotationInput?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  insertMark(platform: string, annotationInput?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  speakerTrack(platform: string, input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  participantTrack(platform: string, input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  timelineView(platform: string, input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  adapterRoute(platform: string, options?: Record<string, unknown>): Promise<unknown>;
  adapterRoutes(options?: Record<string, unknown>): Promise<unknown>;
  adapterBlueprint(platform: string, options?: Record<string, unknown>): Promise<unknown>;
  adapterBlueprints(options?: Record<string, unknown>): Promise<unknown>;
  runManifest(options?: Record<string, unknown>): Promise<unknown>;
  runHandoffReadiness(options?: Record<string, unknown>): Promise<unknown>;
}

export function buildMeetingAppTimelineConnectorPackageAcceptanceReport(
  pkg?: MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorPackageAcceptanceReport;

export function assertMeetingAppTimelineConnectorPackage<T extends MeetingAppTimelineConnectorPackage | Record<string, unknown>>(
  pkg?: T,
  options?: Record<string, unknown>,
): T;

export function stripMeetingAppTimelineConnectorPackageFileContents<T extends MeetingAppTimelineConnectorPackage | Record<string, unknown>>(
  pkg?: T,
): T;

export function buildMeetingAppTimelineConnectorHandoff(
  pkg?: MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorHandoff;

export function buildMeetingAppTimelineConnectorHostInstallChecklist(
  pkg?: MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorHostInstallChecklist;

export function buildMeetingAppTimelineConnectorHostInstallChecklistAcceptanceReport(
  checklistOrPackage?: MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorHostInstallChecklistAcceptanceReport;

export function assertMeetingAppTimelineConnectorHostInstallChecklist<T extends MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>>(
  checklistOrPackage?: T,
  options?: Record<string, unknown>,
): T;

export function buildMeetingAppTimelineConnectorFieldIntakeIndex(
  checklistOrPackage?: MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorFieldIntakeIndex;

export function assertMeetingAppTimelineConnectorFieldIntakeIndex(
  checklistOrPackage?: MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorFieldIntakeIndex;

export function buildMeetingAppTimelineConnectorAdoptionIndex(
  checklistOrPackage?: MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorAdoptionIndex;

export function assertMeetingAppTimelineConnectorAdoptionIndex(
  checklistOrPackage?: MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorAdoptionIndex;

export function buildMeetingAppTimelineConnectorBridgeHandoff(
  pkgOrChecklist?: MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorBridgeHandoff;

export function buildMeetingAppTimelineConnectorBridgeHandoffAcceptanceReport(
  handoffOrPackageOrChecklist?: MeetingAppTimelineConnectorBridgeHandoff | MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorBridgeHandoffAcceptanceReport;

export function assertMeetingAppTimelineConnectorBridgeHandoff<T extends MeetingAppTimelineConnectorBridgeHandoff | MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>>(
  handoffOrPackageOrChecklist?: T,
  options?: Record<string, unknown>,
): T;

export function runMeetingAppTimelineConnectorBridgeSmoke(
  handoffOrPackageOrChecklist?: MeetingAppTimelineConnectorBridgeHandoff | MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): Promise<MeetingAppTimelineConnectorBridgeSmokeReport>;

export function assertMeetingAppTimelineConnectorBridgeSmoke(
  handoffOrPackageOrChecklist?: MeetingAppTimelineConnectorBridgeHandoff | MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): Promise<MeetingAppTimelineConnectorBridgeSmokeReport>;

export function buildMeetingAppTimelineConnectorSmokePlan(
  checklistOrPackage?: MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorSmokePlan;

export function buildMeetingAppTimelineConnectorSmokePlanAcceptanceReport(
  planOrChecklistOrPackage?: MeetingAppTimelineConnectorSmokePlan | MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorSmokePlanAcceptanceReport;

export function assertMeetingAppTimelineConnectorSmokePlan<T extends MeetingAppTimelineConnectorSmokePlan | MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>>(
  planOrChecklistOrPackage?: T,
  options?: Record<string, unknown>,
): T;

export function runMeetingAppTimelineConnectorSmokePlan(
  planOrChecklistOrPackage?: MeetingAppTimelineConnectorSmokePlan | MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): Promise<MeetingAppTimelineConnectorSmokeRunReport>;

export function assertMeetingAppTimelineConnectorSmokeRun(
  planOrChecklistOrPackage?: MeetingAppTimelineConnectorSmokePlan | MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): Promise<MeetingAppTimelineConnectorSmokeRunReport>;

export function buildMeetingAppTimelineConnectorReleaseGate(
  checklistOrPackage?: MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorReleaseGate;

export function assertMeetingAppTimelineConnectorReleaseGate(
  checklistOrPackage?: MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorReleaseGate;

export function buildMeetingAppTimelineConnectorPlatformRoadmap(
  checklistOrPackage?: MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorPlatformRoadmap;

export function assertMeetingAppTimelineConnectorPlatformRoadmap(
  checklistOrPackage?: MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorPlatformRoadmap;

export function buildMeetingAppTimelineConnectorAdapterMatrix(
  checklistOrPackage?: MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorAdapterMatrix;

export function buildMeetingAppTimelineConnectorAdapterMatrixAcceptanceReport(
  matrixOrChecklistOrPackage?: MeetingAppTimelineConnectorAdapterMatrix | MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorAdapterMatrixAcceptanceReport;

export function assertMeetingAppTimelineConnectorAdapterMatrix(
  matrixOrChecklistOrPackage?: MeetingAppTimelineConnectorAdapterMatrix | MeetingAppTimelineConnectorHostInstallChecklist | MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorAdapterMatrix;

export function createMeetingAppTimelineConnectorRuntimeClient(
  pkg?: MeetingAppTimelineConnectorPackage | Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppTimelineConnectorRuntimeClient;
