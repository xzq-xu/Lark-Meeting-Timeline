import type { MeetingPlatformRolloutOptions } from './platform-rollout.mjs';
import type {
  MeetingPlatformAdapterContract,
  MeetingPlatformAdapterContractAcceptanceReport,
} from './platform-adapter-contract.mjs';
import type { MeetingPlatformRuntimeProfile } from './platform-runtime-profile.mjs';
import type { MeetingPlatformRuntimeEventPlan } from './platform-runtime-event.mjs';
import type { MeetingPlatformLiveAdapterReadiness } from './platform-live-adapter.mjs';
import type { MeetingPlatformHandoffReadiness } from './platform-handoff-readiness.mjs';
import type { MeetingPlatformAdaptationStrategy } from './platform-strategy.mjs';

export const MEETING_PLATFORM_ADAPTATION_PACKAGE_SCHEMA: 'meeting_platform_adaptation_package';
export const MEETING_PLATFORM_ADAPTATION_PACKAGE_MATRIX_SCHEMA: 'meeting_platform_adaptation_package_matrix';
export const MEETING_PLATFORM_ADAPTATION_PACKAGE_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdaptationPackageOptions extends MeetingPlatformRolloutOptions {
  platforms?: string[];
  platform_keys?: string[];
  packageId?: string;
  package_id?: string;
  baseUrl?: string;
  base_url?: string;
}

export interface MeetingPlatformAdaptationPackage {
  type: 'meeting_platform_adaptation_package';
  schema: 'meeting_platform_adaptation_package';
  schema_version: 1;
  id: string;
  platform: string;
  display_name?: string;
  objective: string;
  mode?: string;
  recommended_mode?: string;
  runtime_contract: Record<string, unknown>;
  local_observer: Record<string, unknown>;
  provider_observer?: Record<string, unknown>;
  adaptation_strategy: MeetingPlatformAdaptationStrategy;
  adaptation_playbook: MeetingPlatformAdaptationStrategy['adaptation_playbook'];
  annotation_pipeline: Record<string, unknown>;
  runtime_event_plan: MeetingPlatformRuntimeEventPlan;
  speaker_markers?: Record<string, unknown>;
  transcript: Record<string, unknown>;
  extension: {
    matches: string[];
    host_permissions: string[];
    content_scripts: Record<string, unknown>[];
    install_plan: Record<string, unknown>;
  };
  evidence: Record<string, unknown>;
  implementation: Record<string, unknown>;
  commands: Record<string, string>;
  readiness: {
    sdk_wiring_ready: boolean;
    realtime_insertion_contract_ready: boolean;
    handoff_ready: boolean;
    production_ready: boolean;
    ready_for_realtime_annotations: boolean;
    acceptance_target?: string;
    issue_count: number;
    first_issue?: string;
    live_readiness_status?: string;
    handoff_status?: string;
    missing_items: string[];
  };
  contract_acceptance: MeetingPlatformAdapterContractAcceptanceReport;
  live_adapter_readiness: MeetingPlatformLiveAdapterReadiness;
  handoff_readiness: MeetingPlatformHandoffReadiness;
  contract: MeetingPlatformAdapterContract;
  runtime_profile: MeetingPlatformRuntimeProfile;
  collector: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingPlatformAdaptationPackageMatrix {
  type: 'meeting_platform_adaptation_package_matrix';
  schema: 'meeting_platform_adaptation_package_matrix';
  schema_version: 1;
  platform_count: number;
  sdk_wiring_ready_count: number;
  browser_observer_count: number;
  candidate_observer_count: number;
  provider_observer_count: number;
  production_ready_count: number;
  realtime_ready_count: number;
  platforms: string[];
  rows: Record<string, unknown>[];
  packages: MeetingPlatformAdaptationPackage[];
  next_actions: string[];
}

export function buildMeetingPlatformAdaptationPackage(
  platform: string,
  options?: MeetingPlatformAdaptationPackageOptions,
): MeetingPlatformAdaptationPackage;

export function buildMeetingPlatformAdaptationPackageMatrix(
  options?: MeetingPlatformAdaptationPackageOptions,
): MeetingPlatformAdaptationPackageMatrix;
