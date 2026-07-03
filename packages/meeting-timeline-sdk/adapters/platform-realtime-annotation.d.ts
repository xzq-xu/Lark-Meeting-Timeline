import type { TimelineMarkInput } from '../index.mjs';
import type {
  MeetingPlatformAnnotationIntake,
  MeetingPlatformAnnotationIntakeOptions,
  MeetingPlatformAnnotationIntakePlan,
} from './platform-annotation-intake.mjs';
import type {
  MeetingPlatformClockSyncOptions,
  MeetingPlatformClockSyncPlan,
  MeetingPlatformClockSyncReport,
} from './platform-clock-sync.mjs';
import type {
  MeetingPlatformSessionBinding,
  MeetingPlatformSessionBindingOptions,
  MeetingPlatformSessionBindingPlan,
} from './platform-session-binding.mjs';

export const MEETING_PLATFORM_REALTIME_ANNOTATION_PLAN_SCHEMA: 'meeting_platform_realtime_annotation_plan';
export const MEETING_PLATFORM_REALTIME_ANNOTATION_MATRIX_SCHEMA: 'meeting_platform_realtime_annotation_matrix';
export const MEETING_PLATFORM_REALTIME_ANNOTATION_SCHEMA: 'meeting_platform_realtime_annotation';
export const MEETING_PLATFORM_REALTIME_ANNOTATION_SCHEMA_VERSION: number;

export interface MeetingPlatformRealtimeAnnotationOptions extends Record<string, unknown> {
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  allowUnsyncedCapturedAtMs?: boolean;
  allow_unsynced_captured_at_ms?: boolean;
  clock?: MeetingPlatformClockSyncOptions;
  clock_options?: MeetingPlatformClockSyncOptions;
  binding?: MeetingPlatformSessionBindingOptions;
  binding_options?: MeetingPlatformSessionBindingOptions;
  intake?: MeetingPlatformAnnotationIntakeOptions;
  intake_options?: MeetingPlatformAnnotationIntakeOptions;
}

export interface MeetingPlatformRealtimeAnnotationInput {
  platform?: string;
  annotation?: TimelineMarkInput | Record<string, unknown>;
  mark?: TimelineMarkInput | Record<string, unknown>;
  item?: TimelineMarkInput | Record<string, unknown>;
  payload?: Record<string, unknown>;
  current_meeting?: Record<string, unknown>;
  currentMeeting?: Record<string, unknown>;
  current_axis?: Record<string, unknown>;
  currentAxis?: Record<string, unknown>;
  meeting?: Record<string, unknown>;
  local_observer?: Record<string, unknown>;
  localObserver?: Record<string, unknown>;
  local_observers?: Array<Record<string, unknown>>;
  localObservers?: Array<Record<string, unknown>>;
  provider_signal?: Record<string, unknown>;
  providerSignal?: Record<string, unknown>;
  provider_signals?: Array<Record<string, unknown>>;
  providerSignals?: Array<Record<string, unknown>>;
  provider_event?: Record<string, unknown>;
  providerEvent?: Record<string, unknown>;
  provider_events?: Array<Record<string, unknown>>;
  providerEvents?: Array<Record<string, unknown>>;
  signal?: Record<string, unknown>;
  signals?: Array<Record<string, unknown>>;
  samples?: Array<Record<string, unknown>>;
  clock_samples?: Array<Record<string, unknown>>;
  clockSamples?: Array<Record<string, unknown>>;
  time_sync_samples?: Array<Record<string, unknown>>;
  timeSyncSamples?: Array<Record<string, unknown>>;
  clock_sync?: Record<string, unknown>;
  clockSync?: Record<string, unknown>;
  time_sync?: Record<string, unknown>;
  timeSync?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingPlatformRealtimeAnnotationPlan {
  type: 'meeting_platform_realtime_annotation_plan';
  schema: 'meeting_platform_realtime_annotation_plan';
  schema_version: number;
  platform: string;
  display_name?: string;
  status: string;
  pipeline: string[];
  modules: Record<string, string>;
  clock_sync: MeetingPlatformClockSyncPlan;
  session_binding: MeetingPlatformSessionBindingPlan;
  annotation_intake: MeetingPlatformAnnotationIntakePlan;
  realtime_policy: Record<string, unknown>;
  output_contract: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingPlatformRealtimeAnnotationMatrix {
  type: 'meeting_platform_realtime_annotation_matrix';
  schema: 'meeting_platform_realtime_annotation_matrix';
  schema_version: number;
  platform_count: number;
  provider_blocking_count: number;
  transcript_blocking_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  dependencies: Record<string, Array<Record<string, unknown>>>;
  plans: MeetingPlatformRealtimeAnnotationPlan[];
  next_actions: string[];
}

export interface MeetingPlatformRealtimeAnnotation {
  type: 'meeting_platform_realtime_annotation';
  schema: 'meeting_platform_realtime_annotation';
  schema_version: number;
  platform: string;
  status: string;
  accepted_for_realtime: boolean;
  actions: string[];
  clock_sync: MeetingPlatformClockSyncReport;
  session_binding: MeetingPlatformSessionBinding;
  annotation_intake: MeetingPlatformAnnotationIntake;
  selected_meeting?: Record<string, unknown>;
  calibrated_annotation?: Record<string, unknown>;
  start_payload?: Record<string, unknown>;
  insert_payload?: Record<string, unknown>;
  pending_payload?: Record<string, unknown>;
  warnings: string[];
  next_actions: string[];
}

export function buildMeetingPlatformRealtimeAnnotationPlan(
  platform: string,
  options?: MeetingPlatformRealtimeAnnotationOptions,
): MeetingPlatformRealtimeAnnotationPlan;

export function buildMeetingPlatformRealtimeAnnotationMatrix(
  options?: MeetingPlatformRealtimeAnnotationOptions,
): MeetingPlatformRealtimeAnnotationMatrix;

export function buildMeetingPlatformRealtimeAnnotation(
  platform: string,
  input?: MeetingPlatformRealtimeAnnotationInput,
  options?: MeetingPlatformRealtimeAnnotationOptions,
): MeetingPlatformRealtimeAnnotation;
