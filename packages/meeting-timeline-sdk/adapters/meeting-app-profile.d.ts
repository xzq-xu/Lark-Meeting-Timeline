import type { MeetingAppDomCaptureProfilePlatform } from './meeting-app-capture.mjs';
import type { MeetingAppLaunchGateOptions } from './meeting-app-gate.mjs';

export const MEETING_APP_INTEGRATION_PROFILE_SCHEMA: 'meeting_app_integration_profile';
export const MEETING_APP_INTEGRATION_PROFILE_SCHEMA_VERSION: number;
export const MEETING_APP_INTEGRATION_PROFILE_PLATFORMS: readonly MeetingAppDomCaptureProfilePlatform[];

export interface MeetingAppIntegrationProfileOptions extends MeetingAppLaunchGateOptions {
  platform?: MeetingAppDomCaptureProfilePlatform | string;
  provider?: MeetingAppDomCaptureProfilePlatform | string;
  key?: MeetingAppDomCaptureProfilePlatform | string;
  name?: MeetingAppDomCaptureProfilePlatform | string;
  platforms?: Array<MeetingAppDomCaptureProfilePlatform | string> | MeetingAppDomCaptureProfilePlatform | string;
  platform_keys?: Array<MeetingAppDomCaptureProfilePlatform | string> | MeetingAppDomCaptureProfilePlatform | string;
  platformKeys?: Array<MeetingAppDomCaptureProfilePlatform | string> | MeetingAppDomCaptureProfilePlatform | string;
  includeLaunchGate?: boolean;
  include_launch_gate?: boolean;
  launchGateOptions?: MeetingAppLaunchGateOptions;
  launch_gate_options?: MeetingAppLaunchGateOptions;
  [key: string]: unknown;
}

export interface MeetingAppIntegrationProfile {
  type: 'meeting_app_integration_profile';
  schema: 'meeting_app_integration_profile';
  version: number;
  platform: MeetingAppDomCaptureProfilePlatform;
  display_name: string;
  recommended_mode: string;
  extension: Record<string, unknown>;
  runtime: Record<string, unknown>;
  capture: Record<string, unknown>;
  event_model: Record<string, unknown>;
  provider_reconciliation: Record<string, unknown>;
  readiness?: Record<string, unknown>;
  launch_gate?: Record<string, unknown>;
  implementation_steps: string[];
  limitations: string[];
}

export interface MeetingAppIntegrationMatrix {
  type: 'meeting_app_integration_matrix';
  schema: 'meeting_app_integration_profile';
  version: number;
  platform_count: number;
  platforms: MeetingAppDomCaptureProfilePlatform[];
  rows: Array<Record<string, unknown>>;
}

export function buildMeetingAppIntegrationProfile(
  platformOrInput?: MeetingAppDomCaptureProfilePlatform | string | MeetingAppIntegrationProfileOptions,
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppIntegrationProfile;

export function buildAllMeetingAppIntegrationProfiles(
  options?: MeetingAppIntegrationProfileOptions,
): Partial<Record<MeetingAppDomCaptureProfilePlatform, MeetingAppIntegrationProfile>>;

export function buildMeetingAppIntegrationMatrix(
  options?: MeetingAppIntegrationProfileOptions,
): MeetingAppIntegrationMatrix;

export default buildMeetingAppIntegrationProfile;
