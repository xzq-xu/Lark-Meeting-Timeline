import { compactObject } from '../index.mjs';
import {
  buildMeetingPlatformAnnotationIntake,
  buildMeetingPlatformAnnotationIntakeMatrix,
  buildMeetingPlatformAnnotationIntakePlan,
} from './platform-annotation-intake.mjs';
import {
  buildMeetingPlatformClockSyncMatrix,
  buildMeetingPlatformClockSyncPlan,
  buildMeetingPlatformClockSyncReport,
} from './platform-clock-sync.mjs';
import {
  buildMeetingPlatformSessionBinding,
  buildMeetingPlatformSessionBindingMatrix,
  buildMeetingPlatformSessionBindingPlan,
} from './platform-session-binding.mjs';
import {
  MEETING_PLATFORM_KEYS,
  buildPlatformIntegrationPlan,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_REALTIME_ANNOTATION_PLAN_SCHEMA = 'meeting_platform_realtime_annotation_plan';
export const MEETING_PLATFORM_REALTIME_ANNOTATION_MATRIX_SCHEMA = 'meeting_platform_realtime_annotation_matrix';
export const MEETING_PLATFORM_REALTIME_ANNOTATION_SCHEMA = 'meeting_platform_realtime_annotation';
export const MEETING_PLATFORM_REALTIME_ANNOTATION_SCHEMA_VERSION = 1;

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

function unique(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function boolOption(options, keys, fallback) {
  const value = firstNonEmpty(...keys.map((key) => options[key]));
  return value == null ? fallback : value !== false && value !== 'false';
}

function selectedPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_PLATFORM_KEYS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

function optionGroup(options = {}, name) {
  return {
    ...options,
    ...(options[name] ?? {}),
    ...(options[`${name}_options`] ?? {}),
  };
}

function annotationInput(input = {}) {
  return firstNonEmpty(input.annotation, input.mark, input.item, input.payload?.annotation);
}

function currentMeetingInput(input = {}) {
  return firstNonEmpty(input.current_meeting, input.currentMeeting, input.current_axis, input.currentAxis, input.meeting);
}

function pendingMeeting(platform) {
  return {
    platform,
    pending_binding: true,
  };
}

function meetingForIntake(platform, input = {}, binding = {}) {
  if (binding.status === 'bound_to_current_axis') {
    return currentMeetingInput(input) ?? binding.selected_meeting;
  }
  if (binding.should_start_axis) {
    return binding.selected_meeting ?? binding.start_payload;
  }
  if (binding.should_store_pending) return pendingMeeting(platform);
  if (binding.status === 'binding_conflict') return pendingMeeting(platform);
  if (binding.status === 'insufficient_identity') return undefined;
  return currentMeetingInput(input);
}

function clockCanProceed(clock = {}, options = {}) {
  if (clock.accepted_for_realtime) return true;
  return boolOption(options, ['allowUnsyncedCapturedAtMs', 'allow_unsynced_captured_at_ms'], false);
}

function finalStatus({ clock, binding, intake, options }) {
  if (!clockCanProceed(clock, options)) return 'needs_clock_sync';
  if (binding.status === 'binding_conflict') return 'binding_conflict';
  if (binding.status === 'insufficient_identity') return 'insufficient_identity';
  if (binding.should_start_axis && intake.status === 'ready_to_insert_current_axis') return 'start_axis_then_insert';
  if (binding.bind_to_current_axis && intake.status === 'ready_to_insert_current_axis') return 'ready_to_insert';
  if (intake.status === 'pending_real_meeting') return 'pending_real_meeting';
  if (intake.status === 'needs_device_captured_at') return 'needs_device_captured_at';
  if (intake.status === 'after_meeting_end') return 'after_meeting_end';
  if (intake.status === 'before_meeting_start') return 'before_meeting_start';
  if (intake.status === 'start_open_session_then_insert') return 'start_open_session_then_insert';
  if (intake.status === 'ready_to_insert_current_axis') return 'ready_to_insert';
  return 'not_ready';
}

function actionsFor(status) {
  return {
    ready_to_insert: ['insert_mark'],
    start_axis_then_insert: ['start_meeting_session', 'insert_mark'],
    start_open_session_then_insert: ['start_open_session', 'insert_mark'],
    pending_real_meeting: ['store_pending_mark', 'rebind_when_meeting_identity_arrives'],
    needs_clock_sync: ['run_clock_sync_before_realtime_insert'],
    needs_device_captured_at: ['resend_annotation_with_captured_at_ms'],
    binding_conflict: ['do_not_insert_until_meeting_identity_conflict_is_resolved'],
    insufficient_identity: ['collect_meeting_url_or_current_axis_before_insert'],
    after_meeting_end: ['store_for_audit_only'],
    before_meeting_start: ['verify_clock_sync_and_meeting_start_time'],
    not_ready: ['inspect_realtime_annotation_pipeline'],
  }[status] ?? ['inspect_realtime_annotation_pipeline'];
}

function acceptedForRealtime(status) {
  return ['ready_to_insert', 'start_axis_then_insert', 'start_open_session_then_insert', 'pending_real_meeting'].includes(status);
}

export function buildMeetingPlatformRealtimeAnnotationPlan(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const integration = buildPlatformIntegrationPlan(key, options);
  return {
    type: 'meeting_platform_realtime_annotation_plan',
    schema: MEETING_PLATFORM_REALTIME_ANNOTATION_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_REALTIME_ANNOTATION_SCHEMA_VERSION,
    platform: key,
    display_name: integration.display_name,
    status: 'realtime_annotation_pipeline_ready',
    pipeline: [
      'clock_sync',
      'session_binding',
      'annotation_intake',
    ],
    modules: {
      clock_sync: '@ai-annotation/meeting-timeline-sdk/adapters/platform-clock-sync',
      session_binding: '@ai-annotation/meeting-timeline-sdk/adapters/platform-session-binding',
      annotation_intake: '@ai-annotation/meeting-timeline-sdk/adapters/platform-annotation-intake',
    },
    clock_sync: buildMeetingPlatformClockSyncPlan(key, optionGroup(options, 'clock')),
    session_binding: buildMeetingPlatformSessionBindingPlan(key, optionGroup(options, 'binding')),
    annotation_intake: buildMeetingPlatformAnnotationIntakePlan(key, optionGroup(options, 'intake')),
    realtime_policy: {
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      clock_sync_required: true,
      session_identity_required_before_insert: true,
    },
    output_contract: {
      status_values: [
        'ready_to_insert',
        'start_axis_then_insert',
        'start_open_session_then_insert',
        'pending_real_meeting',
        'needs_clock_sync',
        'needs_device_captured_at',
        'binding_conflict',
        'insufficient_identity',
        'after_meeting_end',
      ],
      actions_field: 'actions',
      start_payload_field: 'start_payload',
      insert_payload_field: 'insert_payload',
    },
    next_actions: [
      'sync_device_clock',
      'bind_annotation_to_meeting_identity',
      'insert_or_pending_mark_by_pipeline_status',
    ],
  };
}

export function buildMeetingPlatformRealtimeAnnotationMatrix(options = {}) {
  const platforms = selectedPlatforms(options);
  const plans = platforms.map((platform) => buildMeetingPlatformRealtimeAnnotationPlan(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  const clockMatrix = buildMeetingPlatformClockSyncMatrix({ ...options, platforms });
  const bindingMatrix = buildMeetingPlatformSessionBindingMatrix({ ...options, platforms });
  const intakeMatrix = buildMeetingPlatformAnnotationIntakeMatrix({ ...options, platforms });
  return {
    type: 'meeting_platform_realtime_annotation_matrix',
    schema: MEETING_PLATFORM_REALTIME_ANNOTATION_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_REALTIME_ANNOTATION_SCHEMA_VERSION,
    platform_count: plans.length,
    provider_blocking_count: plans.filter((plan) => plan.realtime_policy.provider_events_block_realtime).length,
    transcript_blocking_count: plans.filter((plan) => plan.realtime_policy.transcript_blocks_realtime).length,
    platforms,
    rows: plans.map((plan) => ({
      platform: plan.platform,
      display_name: plan.display_name,
      status: plan.status,
      clock_sync_required: plan.realtime_policy.clock_sync_required,
      provider_events_block_realtime: plan.realtime_policy.provider_events_block_realtime,
      transcript_blocks_realtime: plan.realtime_policy.transcript_blocks_realtime,
    })),
    dependencies: {
      clock_sync: clockMatrix.rows,
      session_binding: bindingMatrix.rows,
      annotation_intake: intakeMatrix.rows,
    },
    plans,
    next_actions: unique(plans.flatMap((plan) => plan.next_actions ?? [])),
  };
}

export function buildMeetingPlatformRealtimeAnnotation(platform, input = {}, options = {}) {
  const key = normalizeMeetingPlatform(firstNonEmpty(input.platform, platform));
  const rawAnnotation = annotationInput(input);
  const clock = buildMeetingPlatformClockSyncReport(key, {
    ...input,
    annotation: rawAnnotation,
  }, optionGroup(options, 'clock'));
  const annotation = clock.calibrated_annotation ?? rawAnnotation;
  const binding = buildMeetingPlatformSessionBinding(key, {
    ...input,
    annotation,
  }, optionGroup(options, 'binding'));
  const intakeMeeting = meetingForIntake(key, input, binding);
  const intake = buildMeetingPlatformAnnotationIntake(key, {
    ...input,
    current_meeting: intakeMeeting,
    annotation,
  }, optionGroup(options, 'intake'));
  const status = finalStatus({ clock, binding, intake, options });
  const actions = actionsFor(status);
  return compactObject({
    type: 'meeting_platform_realtime_annotation',
    schema: MEETING_PLATFORM_REALTIME_ANNOTATION_SCHEMA,
    schema_version: MEETING_PLATFORM_REALTIME_ANNOTATION_SCHEMA_VERSION,
    platform: key,
    status,
    accepted_for_realtime: acceptedForRealtime(status),
    actions,
    clock_sync: clock,
    session_binding: binding,
    annotation_intake: intake,
    selected_meeting: binding.selected_meeting ?? intake.current_meeting,
    calibrated_annotation: annotation,
    start_payload: status === 'start_axis_then_insert'
      ? binding.start_payload
      : status === 'start_open_session_then_insert'
        ? intake.open_session_payload
        : undefined,
    insert_payload: actions.includes('insert_mark') ? intake.insert_payload : undefined,
    pending_payload: status === 'pending_real_meeting' ? intake.annotation : undefined,
    warnings: unique([
      ...(clock.warnings ?? []),
      ...(binding.conflicts ?? []).map((conflict) => `binding_conflict:${conflict.evidence?.join('+')}`),
      ...(intake.warnings ?? []),
    ]),
    next_actions: unique([
      ...(clock.next_actions ?? []),
      ...(binding.next_actions ?? []),
      ...(intake.next_actions ?? []),
      ...actions,
    ]),
  });
}
