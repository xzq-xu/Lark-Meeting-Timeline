import type {
  MeetingPlatformAdapterDecisionInput,
  MeetingPlatformAdapterDecisionOptions,
} from './platform-adapter-decision.mjs';
import type {
  MeetingPlatformAdapterStartupPlan,
} from './platform-adapter-startup.mjs';
import type {
  MeetingPlatformRawSignalBatch,
} from './platform-raw-signal.mjs';

export const MEETING_PLATFORM_ADAPTER_RUNTIME_RECIPE_SCHEMA: 'meeting_platform_adapter_runtime_recipe';
export const MEETING_PLATFORM_ADAPTER_RUNTIME_RECIPE_MATRIX_SCHEMA: 'meeting_platform_adapter_runtime_recipe_matrix';
export const MEETING_PLATFORM_ADAPTER_RUNTIME_RECIPE_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdapterRuntimeRecipeOptions extends MeetingPlatformAdapterDecisionOptions {
  includeExamples?: boolean;
  include_examples?: boolean;
  filterActiveSpeakerSamples?: boolean;
  filter_active_speaker_samples?: boolean;
  [key: string]: unknown;
}

export interface MeetingPlatformAdapterRuntimeRecipe {
  type: 'meeting_platform_adapter_runtime_recipe';
  schema: 'meeting_platform_adapter_runtime_recipe';
  schema_version: 1;
  accepted: boolean;
  runtime_ready: boolean;
  status: string;
  platform?: string;
  display_name?: string;
  selected_surface?: string;
  host_profile?: string;
  install_target?: string;
  runtime_contract?: Record<string, unknown>;
  host_wiring?: Record<string, unknown>;
  sequence?: Array<Record<string, unknown>>;
  raw_signal_examples?: MeetingPlatformRawSignalBatch | Record<string, unknown>;
  startup_plan?: MeetingPlatformAdapterStartupPlan;
  plan?: MeetingPlatformAdapterStartupPlan;
  readiness?: Record<string, unknown>;
  issues?: Array<Record<string, unknown>>;
  next_actions: string[];
}

export interface MeetingPlatformAdapterRuntimeRecipeMatrix {
  type: 'meeting_platform_adapter_runtime_recipe_matrix';
  schema: 'meeting_platform_adapter_runtime_recipe_matrix';
  schema_version: 1;
  platform_count: number;
  accepted_count: number;
  runtime_ready_count: number;
  browser_surface_count: number;
  native_surface_count: number;
  provider_reconcile_surface_count: number;
  raw_signal_runtime_event_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  recipes: MeetingPlatformAdapterRuntimeRecipe[];
  next_actions: string[];
}

export function buildMeetingPlatformAdapterRuntimeRecipe(
  input?: string | URL | MeetingPlatformAdapterDecisionInput,
  options?: MeetingPlatformAdapterRuntimeRecipeOptions,
): MeetingPlatformAdapterRuntimeRecipe;

export function buildMeetingPlatformAdapterRuntimeRecipeMatrix(
  input?: MeetingPlatformAdapterDecisionInput,
  options?: MeetingPlatformAdapterRuntimeRecipeOptions,
): MeetingPlatformAdapterRuntimeRecipeMatrix;

export function assertMeetingPlatformAdapterRuntimeRecipe(
  input?: string | URL | MeetingPlatformAdapterDecisionInput,
  options?: MeetingPlatformAdapterRuntimeRecipeOptions,
): MeetingPlatformAdapterRuntimeRecipe;

export function assertMeetingPlatformAdapterRuntimeRecipeMatrix(
  input?: MeetingPlatformAdapterDecisionInput,
  options?: MeetingPlatformAdapterRuntimeRecipeOptions,
): MeetingPlatformAdapterRuntimeRecipeMatrix;
