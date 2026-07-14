import type { MeetingPlatformAdapterPortfolioOptions } from './platform-adapter-portfolio.mjs';
import type { MeetingPlatformHandoffReadinessOptions } from './platform-handoff-readiness.mjs';

export const MEETING_PLATFORM_ADAPTER_ACCEPTANCE_CHECKLIST_SCHEMA: 'meeting_platform_adapter_acceptance_checklist';
export const MEETING_PLATFORM_ADAPTER_ACCEPTANCE_CHECKLIST_MATRIX_SCHEMA: 'meeting_platform_adapter_acceptance_checklist_matrix';
export const MEETING_PLATFORM_ADAPTER_ACCEPTANCE_CHECKLIST_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdapterAcceptanceChecklistOptions
  extends MeetingPlatformAdapterPortfolioOptions, MeetingPlatformHandoffReadinessOptions {
  target?: 'static' | 'pilot' | 'production' | string;
  acceptanceTarget?: 'static' | 'pilot' | 'production' | string;
  acceptance_target?: 'static' | 'pilot' | 'production' | string;
}

export interface MeetingPlatformAdapterAcceptanceChecklistItem {
  id: string;
  group: string;
  label: string;
  required_from: string;
  required: boolean;
  passed: boolean;
  status: 'pass' | 'fail' | 'skip' | string;
  blocking: boolean;
  evidence?: unknown;
  expected?: unknown;
  actual?: unknown;
  next_action?: string;
}

export interface MeetingPlatformAdapterAcceptanceChecklist {
  type: 'meeting_platform_adapter_acceptance_checklist';
  schema: 'meeting_platform_adapter_acceptance_checklist';
  schema_version: 1;
  platform: string;
  display_name: string;
  target: string;
  accepted: boolean;
  adapter_status: string;
  recommended_first_surface: string;
  runtime_event_contract: Record<string, unknown>;
  implementation_sequence: Array<Record<string, unknown>>;
  sdk_entrypoints: Record<string, string>;
  evidence_collection_plan: Record<string, unknown>;
  pilot_measurement_contract: Record<string, unknown>;
  checklist: MeetingPlatformAdapterAcceptanceChecklistItem[];
  summary: Record<string, unknown>;
  portfolio_item: Record<string, unknown>;
  handoff_readiness?: Record<string, unknown>;
  commands: Record<string, string>;
  next_actions: string[];
}

export interface MeetingPlatformAdapterAcceptanceChecklistMatrix {
  type: 'meeting_platform_adapter_acceptance_checklist_matrix';
  schema: 'meeting_platform_adapter_acceptance_checklist_matrix';
  schema_version: 1;
  target: string;
  platform_count: number;
  accepted_count: number;
  blocked_count: number;
  static_ready_count: number;
  pilot_ready_count: number;
  production_ready_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  checklists: MeetingPlatformAdapterAcceptanceChecklist[];
  next_actions: string[];
}

export function buildMeetingPlatformAdapterAcceptanceChecklist(
  platform: string,
  input?: MeetingPlatformAdapterAcceptanceChecklistOptions,
  options?: MeetingPlatformAdapterAcceptanceChecklistOptions,
): MeetingPlatformAdapterAcceptanceChecklist;

export function buildMeetingPlatformPilotMeasurementContract(
  platform: string,
  options?: MeetingPlatformAdapterAcceptanceChecklistOptions,
): Record<string, unknown>;

export function buildMeetingPlatformAdapterAcceptanceChecklistMatrix(
  input?: MeetingPlatformAdapterAcceptanceChecklistOptions,
  options?: MeetingPlatformAdapterAcceptanceChecklistOptions,
): MeetingPlatformAdapterAcceptanceChecklistMatrix;
