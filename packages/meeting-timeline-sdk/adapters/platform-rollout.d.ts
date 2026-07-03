import type { MeetingAppLaunchGate, MeetingAppLaunchGateOptions } from './meeting-app-gate.mjs';
import type { PlatformLaunchGateOptions } from './platform-gate.mjs';

export type MeetingPlatformRolloutStatus =
  | 'production_ready'
  | 'realtime_ready_provider_pending'
  | 'provider_ready_collect_local_evidence'
  | 'provider_ready_local_blocked'
  | 'needs_live_dom_and_provider_evidence'
  | 'needs_live_dom_evidence'
  | 'needs_provider_event_evidence'
  | 'needs_detector_event_evidence'
  | 'blocked';

export interface MeetingPlatformRolloutOptions extends PlatformLaunchGateOptions, MeetingAppLaunchGateOptions {
  providerRecords?: unknown[];
  provider_records?: unknown[];
  providerCaptureRecords?: unknown[];
  provider_capture_records?: unknown[];
  providerSamples?: unknown[] | Record<string, unknown[]>;
  provider_samples?: unknown[] | Record<string, unknown[]>;
  meetingAppSnapshots?: Record<string, unknown>[] | Record<string, Record<string, unknown>[] | Record<string, unknown>>;
  meeting_app_snapshots?: Record<string, unknown>[] | Record<string, Record<string, unknown>[] | Record<string, unknown>>;
  meetingAppSnapshotRecords?: unknown[] | Record<string, unknown>;
  meeting_app_snapshot_records?: unknown[] | Record<string, unknown>;
  meetingAppRecords?: unknown[] | Record<string, unknown>;
  meeting_app_records?: unknown[] | Record<string, unknown>;
  meetingAppRecordSet?: Record<string, unknown>;
  meeting_app_record_set?: Record<string, unknown>;
  allowProviderFixtureEvidence?: boolean;
  allow_provider_fixture_evidence?: boolean;
  allowProviderFixtureProduction?: boolean;
  allow_provider_fixture_production?: boolean;
  allowMeetingAppFixtureEvidence?: boolean;
  allow_meeting_app_fixture_evidence?: boolean;
  allowMeetingAppFixtureProduction?: boolean;
  allow_meeting_app_fixture_production?: boolean;
}

export interface MeetingPlatformRolloutPlan {
  type: 'meeting_platform_rollout_plan';
  platform: string;
  display_name?: string;
  status: MeetingPlatformRolloutStatus;
  recommended_mode: string;
  production_ready: boolean;
  ready_for_realtime_annotations: boolean;
  realtime_axis_primary?: string;
  source_priority?: string[];
  local_observer?: {
    supported: boolean;
    status: string;
    runtime_ready: boolean;
    production_ready: boolean;
    evidence_level: string;
    evidence_count: number;
    missing_required_coverage: string[];
    issue_codes: string[];
    gate?: MeetingAppLaunchGate;
  };
  provider_events: {
    status: string;
    gate_status?: string;
    setup_ready: boolean;
    production_ready: boolean;
    evidence_level?: string;
    evidence_count?: number;
    missing_required_coverage: string[];
    issue_codes: string[];
    event_types: string[];
    endpoint?: string;
    gate?: Record<string, unknown>;
  };
  post_meeting_transcript?: Record<string, unknown>;
  required_timestamp_invariant?: string;
  next_actions: string[];
  integration_plan?: Record<string, unknown>;
}

export interface MeetingPlatformRolloutSummary {
  type: 'meeting_platform_rollout_summary';
  ok: boolean;
  production_ready: boolean;
  plan_count: number;
  production_ready_count: number;
  realtime_ready_count: number;
  blocked_count: number;
  pending_platforms: string[];
  next_actions: string[];
  plans: MeetingPlatformRolloutPlan[];
}

export const MEETING_PLATFORM_ROLLOUT_STATUSES: readonly MeetingPlatformRolloutStatus[];

export function buildMeetingPlatformRolloutPlan(
  platform: string,
  options?: MeetingPlatformRolloutOptions,
): MeetingPlatformRolloutPlan;

export function buildAllMeetingPlatformRolloutPlans(
  options?: MeetingPlatformRolloutOptions,
): MeetingPlatformRolloutPlan[];

export function buildMeetingPlatformRolloutSummary(
  options?: MeetingPlatformRolloutOptions,
): MeetingPlatformRolloutSummary;
