import type { MeetingAppBrowserRuntimePreset } from './meeting-app-browser-runtime.mjs';
import type { MeetingAppDomCaptureProfile } from './meeting-app-capture.mjs';
import type {
  MeetingAppFixtureLifecycleDiagnosis,
  MeetingAppFixtureOptions,
  MeetingAppFixturePlatform,
} from './meeting-app-fixtures.mjs';

export type MeetingAppLaunchGateEvidenceLevel = 'captured_dom' | 'fixture_dom' | 'none';

export interface MeetingAppLaunchGateIssue {
  severity: 'error' | 'warning' | string;
  code: string;
  message: string;
  [key: string]: unknown;
}

export interface MeetingAppLaunchGateOptions extends MeetingAppFixtureOptions {
  snapshots?: Record<string, unknown>[] | Record<string, Record<string, unknown>[] | Record<string, unknown>>;
  domSnapshots?: Record<string, unknown>[] | Record<string, Record<string, unknown>[] | Record<string, unknown>>;
  dom_snapshots?: Record<string, unknown>[] | Record<string, Record<string, unknown>[] | Record<string, unknown>>;
  requiredCoverage?: string[];
  required_coverage?: string[];
  requireMeetingEnd?: boolean;
  require_meeting_end?: boolean;
  allowFixtureEvidence?: boolean;
  allow_fixture_evidence?: boolean;
  allowFixtureProduction?: boolean;
  allow_fixture_production?: boolean;
  requireProductionReady?: boolean;
  require_production_ready?: boolean;
}

export interface MeetingAppLaunchGate {
  type: 'meeting_app_launch_gate';
  platform: MeetingAppFixturePlatform;
  status: 'passed' | 'warning' | 'failed' | string;
  passed: boolean;
  production_ready: boolean;
  evidence_level: MeetingAppLaunchGateEvidenceLevel;
  evidence_count: number;
  required_coverage: string[];
  missing_required_coverage: string[];
  coverage: Record<string, boolean>;
  runtime_ready: boolean;
  recommended_runtime?: Record<string, unknown>;
  blocking_issues: MeetingAppLaunchGateIssue[];
  warnings: MeetingAppLaunchGateIssue[];
  next_actions: string[];
  reports: {
    runtime_preset?: MeetingAppBrowserRuntimePreset | null;
    capture_profile?: MeetingAppDomCaptureProfile | null;
    fixture_lifecycle?: MeetingAppFixtureLifecycleDiagnosis | null;
    captured?: Record<string, unknown> | null;
  };
}

export interface MeetingAppLaunchGateSummary {
  type: 'meeting_app_launch_gate_summary';
  ok: boolean;
  production_ready: boolean;
  passed_count: number;
  production_ready_count: number;
  failed_count: number;
  warning_count: number;
  gates: MeetingAppLaunchGate[];
}

export const MEETING_APP_LAUNCH_GATE_EVIDENCE_LEVELS: readonly MeetingAppLaunchGateEvidenceLevel[];

export function buildMeetingAppLaunchGate(
  platform: string,
  options?: MeetingAppLaunchGateOptions,
): MeetingAppLaunchGate;

export function buildAllMeetingAppLaunchGates(
  options?: MeetingAppLaunchGateOptions,
): MeetingAppLaunchGate[];

export function buildMeetingAppLaunchGateSummary(
  options?: MeetingAppLaunchGateOptions,
): MeetingAppLaunchGateSummary;

export function assertMeetingAppLaunchGate(
  platform: string,
  options?: MeetingAppLaunchGateOptions,
): MeetingAppLaunchGate;

export function assertAllMeetingAppLaunchGates(
  options?: MeetingAppLaunchGateOptions,
): MeetingAppLaunchGateSummary;
