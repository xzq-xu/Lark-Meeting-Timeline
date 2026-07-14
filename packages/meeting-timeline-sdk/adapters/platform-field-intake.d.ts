import type { MeetingPlatformFieldCaptureOptions } from './platform-field-capture.mjs';

export const MEETING_PLATFORM_FIELD_INTAKE_PLAN_SCHEMA: 'meeting_platform_field_intake_plan';
export const MEETING_PLATFORM_FIELD_INTAKE_MATRIX_SCHEMA: 'meeting_platform_field_intake_matrix';
export const MEETING_PLATFORM_FIELD_INTAKE_SCHEMA_VERSION: 1;

export interface MeetingPlatformFieldIntakeOptions extends MeetingPlatformFieldCaptureOptions {
  includeManifest?: boolean;
  include_manifest?: boolean;
  includeCollectorConfig?: boolean;
  include_collector_config?: boolean;
  includeRealIntakePlan?: boolean;
  include_real_intake_plan?: boolean;
}

export interface MeetingPlatformFieldIntakePlan {
  type: 'meeting_platform_field_intake_plan';
  schema: 'meeting_platform_field_intake_plan';
  schema_version: 1;
  platform: string;
  display_name?: string;
  status: string;
  production_ready: boolean;
  ready_for_realtime_annotations: boolean;
  objective: string;
  base_url: string;
  provider_endpoint?: string;
  evidence_dirs: Record<string, string>;
  files?: Record<string, string>;
  required_inputs?: string[];
  forbidden_inputs?: string[];
  provider_connection: Record<string, unknown>;
  local_observer?: Record<string, unknown>;
  provider_observer: Record<string, unknown>;
  acceptance: Record<string, unknown>;
  commands: Record<string, string>;
  operator_steps: Array<Record<string, unknown>>;
  sdk_methods: string[];
  handoff: Record<string, unknown>;
  field_capture_plan: Record<string, unknown>;
  field_capture_manifest?: Record<string, unknown>;
  collector_config?: Record<string, unknown>;
  real_intake_plan?: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingPlatformFieldIntakeMatrix {
  type: 'meeting_platform_field_intake_matrix';
  schema: 'meeting_platform_field_intake_matrix';
  schema_version: 1;
  platform_count: number;
  production_ready_count: number;
  realtime_ready_count: number;
  provider_blocked_count: number;
  local_capture_needed_count: number;
  provider_capture_needed_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  plans: MeetingPlatformFieldIntakePlan[];
  next_actions: string[];
}

export function buildMeetingPlatformFieldIntakePlan(
  platform: string,
  options?: MeetingPlatformFieldIntakeOptions,
): MeetingPlatformFieldIntakePlan;

export function buildMeetingPlatformFieldIntakeMatrix(
  options?: MeetingPlatformFieldIntakeOptions,
): MeetingPlatformFieldIntakeMatrix;
