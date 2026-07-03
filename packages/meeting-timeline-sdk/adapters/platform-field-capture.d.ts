import type { MeetingPlatformRuntimeProfileOptions } from './platform-runtime-profile.mjs';
import type {
  MeetingPlatformEvidencePackage,
  MeetingPlatformEvidencePackageOptions,
  MeetingPlatformEvidencePackageSummary,
  MeetingPlatformEvidencePackageVerification,
} from './platform-evidence-package.mjs';

export const MEETING_PLATFORM_FIELD_CAPTURE_PLAN_SCHEMA: 'meeting_platform_field_capture_plan';
export const MEETING_PLATFORM_FIELD_CAPTURE_MATRIX_SCHEMA: 'meeting_platform_field_capture_matrix';
export const MEETING_PLATFORM_FIELD_EVIDENCE_BUNDLE_SCHEMA: 'meeting_platform_field_evidence_bundle';
export const MEETING_PLATFORM_FIELD_EVIDENCE_MATRIX_SCHEMA: 'meeting_platform_field_evidence_matrix';
export const MEETING_PLATFORM_FIELD_CAPTURE_SCHEMA_VERSION: 1;

export interface MeetingPlatformFieldCaptureOptions extends MeetingPlatformRuntimeProfileOptions, MeetingPlatformEvidencePackageOptions {
  evidencePackage?: unknown;
  evidence_package?: unknown;
  platformEvidencePackage?: unknown;
  platform_evidence_package?: unknown;
  'package'?: unknown;
  input?: unknown;
  inputs?: unknown;
  evidence?: unknown;
  evidenceInput?: unknown;
  evidence_input?: unknown;
  fieldEvidence?: unknown;
  field_evidence?: unknown;
  fieldCaptureEvidence?: unknown;
  field_capture_evidence?: unknown;
  evidenceDir?: string;
  evidence_dir?: string;
}

export interface MeetingPlatformFieldCapturePlan {
  type: 'meeting_platform_field_capture_plan';
  schema: 'meeting_platform_field_capture_plan';
  schema_version: 1;
  platform: string;
  display_name?: string;
  status: string;
  production_ready: boolean;
  ready_for_realtime_annotations: boolean;
  objective: string;
  output_paths: Record<string, string>;
  runtime_contract: Record<string, unknown>;
  local_observer?: Record<string, unknown>;
  provider_events: Record<string, unknown>;
  checklist: Record<string, unknown>[];
  current_evidence?: Record<string, unknown>;
  missing_items: string[];
  validation: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingPlatformFieldCaptureMatrix {
  type: 'meeting_platform_field_capture_matrix';
  schema: 'meeting_platform_field_capture_matrix';
  schema_version: 1;
  platform_count: number;
  production_ready_count: number;
  realtime_ready_count: number;
  missing_item_count: number;
  platforms: string[];
  rows: Record<string, unknown>[];
  plans: MeetingPlatformFieldCapturePlan[];
  next_actions: string[];
}

export interface MeetingPlatformFieldEvidenceBundle {
  type: 'meeting_platform_field_evidence_bundle';
  schema: 'meeting_platform_field_evidence_bundle';
  schema_version: 1;
  platform: string;
  status: string;
  production_ready: boolean;
  ready_for_realtime_annotations: boolean;
  package_id: string;
  evidence_package: MeetingPlatformEvidencePackage;
  evidence_summary: MeetingPlatformEvidencePackageSummary;
  verification: MeetingPlatformEvidencePackageVerification;
  field_capture_plan: MeetingPlatformFieldCapturePlan;
  handoff: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingPlatformFieldEvidenceMatrix {
  type: 'meeting_platform_field_evidence_matrix';
  schema: 'meeting_platform_field_evidence_matrix';
  schema_version: 1;
  platform_count: number;
  production_ready_count: number;
  realtime_ready_count: number;
  verified_count: number;
  missing_item_count: number;
  platforms: string[];
  rows: Record<string, unknown>[];
  bundles: MeetingPlatformFieldEvidenceBundle[];
  next_actions: string[];
}

export function buildMeetingPlatformFieldCapturePlan(
  platform: string,
  options?: MeetingPlatformFieldCaptureOptions,
): MeetingPlatformFieldCapturePlan;

export function buildMeetingPlatformFieldCaptureMatrix(
  options?: MeetingPlatformFieldCaptureOptions,
): MeetingPlatformFieldCaptureMatrix;

export function buildMeetingPlatformFieldEvidenceBundle(
  platform: string,
  input?: MeetingPlatformEvidencePackage | MeetingPlatformEvidencePackageOptions,
  options?: MeetingPlatformFieldCaptureOptions,
): MeetingPlatformFieldEvidenceBundle;

export function buildMeetingPlatformFieldEvidenceMatrix(
  options?: MeetingPlatformFieldCaptureOptions,
): MeetingPlatformFieldEvidenceMatrix;
