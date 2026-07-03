import {
  buildTimelineMark,
  compactObject,
  normalizeAbsoluteMs,
} from '../index.mjs';
import {
  MEETING_PLATFORM_KEYS,
  buildPlatformIntegrationPlan,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_ANNOTATION_INTAKE_PLAN_SCHEMA = 'meeting_platform_annotation_intake_plan';
export const MEETING_PLATFORM_ANNOTATION_INTAKE_MATRIX_SCHEMA = 'meeting_platform_annotation_intake_matrix';
export const MEETING_PLATFORM_ANNOTATION_INTAKE_SCHEMA = 'meeting_platform_annotation_intake';
export const MEETING_PLATFORM_ANNOTATION_INTAKE_SCHEMA_VERSION = 1;

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

function numberOption(options, keys, fallback) {
  const value = firstNonEmpty(...keys.map((key) => options[key]));
  const parsed = Number(value ?? fallback);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function selectedPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_PLATFORM_KEYS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

function maybeAbsoluteMs(value, fieldName) {
  if (value == null || value === '') return undefined;
  try {
    return normalizeAbsoluteMs(value, fieldName);
  } catch {
    return undefined;
  }
}

function deepestPoints(value) {
  const rows = [];
  for (const item of asArray(value)) {
    if (Array.isArray(item)) rows.push(...deepestPoints(item));
    else if (item && typeof item === 'object') rows.push(item);
  }
  return rows;
}

function capturedAtMs(input = {}) {
  const direct = firstNonEmpty(
    input.captured_at_ms,
    input.capturedAtMs,
    input.captured_at,
    input.capturedAt,
    input.ink_end_at_ms,
    input.inkEndAtMs,
    input.ink_end_time_ms,
    input.inkEndTimeMs,
    input.stroke_end_at_ms,
    input.strokeEndAtMs,
    input.stroke_end_time_ms,
    input.strokeEndTimeMs,
    input.timestamp_ms,
    input.timestampMs,
    input.timestamp,
    input.ts,
    input.created_at_ms,
    input.createdAtMs,
    input.device_time_ms,
    input.deviceTimeMs,
    input.timing?.captured_at_ms,
    input.timing?.capturedAtMs,
    input.payload?.timing?.captured_at_ms,
    input.payload?.timing?.capturedAtMs,
  );
  const directMs = maybeAbsoluteMs(direct, 'captured_at_ms');
  if (directMs != null) return { value: directMs, source: 'captured_at_ms' };
  const points = [
    ...deepestPoints(input.strokes),
    ...deepestPoints(input.stroke_points),
    ...deepestPoints(input.strokePoints),
    ...deepestPoints(input.payload?.strokes),
    ...deepestPoints(input.payload?.stroke_points),
    ...deepestPoints(input.payload?.strokePoints),
  ];
  const pointTimes = points
    .map((point) => maybeAbsoluteMs(firstNonEmpty(point.t, point.ts, point.timestamp_ms, point.timestampMs, point.time_ms, point.timeMs), 'stroke.timestamp_ms'))
    .filter((value) => value != null);
  if (pointTimes.length > 0) return { value: Math.max(...pointTimes), source: 'stroke_point_time' };
  return { value: undefined, source: null };
}

function annotationInput(input = {}) {
  return input.annotation ?? input.mark ?? input.item ?? input;
}

function meetingInput(input = {}) {
  return firstNonEmpty(
    input.current_meeting,
    input.currentMeeting,
    input.currentAxis,
    input.current_axis,
    input.state?.meeting,
    input.timelineState?.meeting,
    input.timeline_state?.meeting,
    input.meeting,
    {},
  );
}

function normalizeMeeting(input = {}, platform) {
  const meeting = meetingInput(input);
  const key = normalizeMeetingPlatform(firstNonEmpty(meeting.platform, input.platform, platform));
  const start = maybeAbsoluteMs(firstNonEmpty(
    meeting.start_time_ms,
    meeting.startTimeMs,
    meeting.start_time,
    meeting.startTime,
    meeting.started_at_ms,
    meeting.startedAtMs,
    meeting.started_at,
    meeting.startedAt,
  ), 'meeting.start_time_ms');
  const end = maybeAbsoluteMs(firstNonEmpty(
    meeting.end_time_ms,
    meeting.endTimeMs,
    meeting.end_time,
    meeting.endTime,
    meeting.ended_at_ms,
    meeting.endedAtMs,
    meeting.ended_at,
    meeting.endedAt,
  ), 'meeting.end_time_ms');
  return compactObject({
    platform: key,
    meeting_id: firstNonEmpty(meeting.meeting_id, meeting.meetingId, meeting.id),
    external_meeting_id: firstNonEmpty(meeting.external_meeting_id, meeting.externalMeetingId),
    meeting_url: firstNonEmpty(meeting.meeting_url, meeting.meetingUrl, meeting.url),
    title: firstNonEmpty(meeting.title, meeting.topic, meeting.name),
    source: firstNonEmpty(meeting.source, input.source),
    pending_binding: Boolean(meeting.pending_binding ?? meeting.pendingBinding),
    start_time_ms: start,
    end_time_ms: end,
  });
}

function isDemoAxis(meeting = {}) {
  return meeting.meeting_id === 'demo-lark-meeting-001' && !meeting.source && !meeting.pending_binding;
}

function axisState(meeting = {}) {
  if (meeting.pending_binding) return 'pending_axis';
  if (!meeting.meeting_id) return 'no_axis';
  if (isDemoAxis(meeting)) return 'demo_axis';
  if (meeting.source === 'local_simulation') return 'local_simulation_axis';
  if (meeting.end_time_ms != null) return 'ended_axis';
  return 'active_axis';
}

function routePolicy(options = {}) {
  return {
    allow_pending_when_no_axis: boolOption(options, ['allowPendingWhenNoAxis', 'allow_pending_when_no_axis'], true),
    allow_open_session_when_no_axis: boolOption(options, ['allowOpenSessionWhenNoAxis', 'allow_open_session_when_no_axis'], true),
    allow_local_simulation_axis: boolOption(options, ['allowLocalSimulationAxis', 'allow_local_simulation_axis'], false),
    allowed_before_start_ms: numberOption(options, ['allowedBeforeStartMs', 'allowed_before_start_ms'], 0),
    open_session_source: String(firstNonEmpty(options.openSessionSource, options.open_session_source, 'first_reliable_annotation')),
    default_title: String(firstNonEmpty(options.defaultTitle, options.default_title, 'Realtime annotation session')),
  };
}

function buildOpenSessionPayload(platform, annotation = {}, meeting = {}, captured = {}, policy = {}) {
  return compactObject({
    platform,
    meeting_id: firstNonEmpty(
      annotation.meeting_id,
      annotation.meetingId,
      annotation.session_id,
      annotation.sessionId,
      meeting.meeting_id,
      captured.value == null ? undefined : `annotation-open-session-${captured.value}`,
    ),
    external_meeting_id: firstNonEmpty(annotation.external_meeting_id, annotation.externalMeetingId, meeting.external_meeting_id),
    meeting_url: firstNonEmpty(annotation.meeting_url, annotation.meetingUrl, annotation.url, annotation.join_url, annotation.joinUrl, meeting.meeting_url),
    title: firstNonEmpty(annotation.meeting_title, annotation.meetingTitle, annotation.context?.meeting_title, meeting.title, policy.default_title),
    start_time_ms: captured.value,
    detector_source: firstNonEmpty(annotation.detector_source, annotation.detectorSource, annotation.source, annotation.device?.type, policy.open_session_source),
    suppress_auto_annotations: true,
  });
}

function buildInsertPayload(annotation = {}, captured = {}, normalizedTimeMs) {
  return buildTimelineMark({
    ...annotation,
    captured_at_ms: captured.value,
    time_ms: normalizedTimeMs,
  }, {
    requireCapturedAt: false,
  });
}

function safeBuildTimelineMark(annotation = {}, fallbackCapturedAtMs) {
  try {
    return buildTimelineMark(annotation, { requireCapturedAt: false });
  } catch {
    return compactObject({
      id: firstNonEmpty(annotation.id, annotation.annotation_id, annotation.annotationId, annotation.mark_id, annotation.markId),
      source: annotation.source ?? annotation.device?.type ?? 'external_annotation',
      captured_at_ms: fallbackCapturedAtMs,
      kind: String(firstNonEmpty(annotation.kind, annotation.type, annotation.intent, annotation.action, 'annotation')),
      label: String(firstNonEmpty(annotation.label, annotation.text, annotation.reading, annotation.intent, '实时标注')),
      text: firstNonEmpty(annotation.text, annotation.reading),
      intent: annotation.intent ?? annotation.payload?.intent,
      mark: annotation.mark ?? annotation.payload?.mark,
      target: annotation.target ?? annotation.target_region ?? annotation.targetRegion ?? annotation.payload?.target,
      strokes: annotation.strokes ?? annotation.stroke_points ?? annotation.strokePoints ?? annotation.payload?.strokes,
      payload: annotation.payload,
    });
  }
}

function nextActionsFor(status) {
  return {
    ready_to_insert_current_axis: ['insert_mark_with_captured_at_ms'],
    start_open_session_then_insert: ['start_open_meeting_session_with_captured_at_ms', 'insert_mark_after_session_start'],
    pending_real_meeting: ['store_mark_as_pending', 'rebind_when_provider_or_local_meeting_start_arrives'],
    needs_device_captured_at: ['resend_mark_with_captured_at_ms_or_absolute_stroke_timestamps'],
    after_meeting_end: ['store_for_audit_but_do_not_count_as_realtime_mark', 'verify_device_clock_and_meeting_end_time'],
    before_meeting_start: ['verify_device_clock_and_meeting_start_time', 'store_for_audit_or_wait_for_matching_axis'],
    local_simulation_axis: ['avoid_counting_local_simulation_as_real_meeting_axis'],
    unbound_no_axis: ['start_meeting_session_or_enable_pending_route_before_inserting'],
  }[status] ?? ['inspect_annotation_intake_status'];
}

export function buildMeetingPlatformAnnotationIntakePlan(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const integration = buildPlatformIntegrationPlan(key, options);
  return compactObject({
    type: 'meeting_platform_annotation_intake_plan',
    schema: MEETING_PLATFORM_ANNOTATION_INTAKE_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_ANNOTATION_INTAKE_SCHEMA_VERSION,
    platform: key,
    display_name: integration.display_name,
    status: 'annotation_intake_contract_ready',
    accepted_time_fields: [
      'captured_at_ms',
      'captured_at',
      'ink_end_at_ms',
      'stroke_end_at_ms',
      'timestamp_ms',
      'device_time_ms',
      'strokes[*].t',
      'stroke_points[*].timestamp_ms',
    ],
    route_policy: routePolicy(options),
    output_contract: {
      annotation_timestamp_field: 'captured_at_ms',
      current_axis_action: 'insert_mark',
      no_axis_action: 'start_open_session_or_pending',
      missing_time_action: 'do_not_count_as_realtime',
      after_end_action: 'store_for_audit_only',
    },
    realtime_policy: {
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      device_clock_sync_required: true,
    },
    next_actions: [
      'normalize_annotation_with_captured_at_ms',
      'route_to_current_axis_pending_or_open_session',
      'poll_annotation_status_until_real_axis_bound_when_pending',
    ],
  });
}

export function buildMeetingPlatformAnnotationIntakeMatrix(options = {}) {
  const plans = selectedPlatforms(options).map((platform) => buildMeetingPlatformAnnotationIntakePlan(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_annotation_intake_matrix',
    schema: MEETING_PLATFORM_ANNOTATION_INTAKE_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_ANNOTATION_INTAKE_SCHEMA_VERSION,
    platform_count: plans.length,
    provider_blocking_count: plans.filter((plan) => plan.realtime_policy.provider_events_block_realtime).length,
    transcript_blocking_count: plans.filter((plan) => plan.realtime_policy.transcript_blocks_realtime).length,
    platforms: plans.map((plan) => plan.platform),
    rows: plans.map((plan) => ({
      platform: plan.platform,
      display_name: plan.display_name,
      status: plan.status,
      no_axis_action: plan.output_contract.no_axis_action,
      provider_events_block_realtime: plan.realtime_policy.provider_events_block_realtime,
      transcript_blocks_realtime: plan.realtime_policy.transcript_blocks_realtime,
    })),
    plans,
    next_actions: unique(plans.flatMap((plan) => plan.next_actions ?? [])),
  };
}

export function buildMeetingPlatformAnnotationIntake(platform, input = {}, options = {}) {
  const policy = routePolicy(options);
  const annotation = annotationInput(input);
  const key = normalizeMeetingPlatform(firstNonEmpty(annotation.platform, input.platform, platform));
  const meeting = normalizeMeeting(input, key);
  const captured = capturedAtMs(annotation);
  const axis = axisState(meeting);
  const warnings = [];
  let status = 'unbound_no_axis';
  let action = 'do_not_insert';
  let normalizedTimeMs;
  let afterMeetingEndMs = 0;

  if (captured.value == null) {
    status = 'needs_device_captured_at';
    action = 'reject_or_store_unbound';
    warnings.push('missing_captured_at_ms');
  } else if (axis === 'active_axis' && meeting.start_time_ms != null) {
    normalizedTimeMs = captured.value - meeting.start_time_ms;
    if (normalizedTimeMs < -policy.allowed_before_start_ms) {
      status = 'before_meeting_start';
      action = 'store_for_audit_only';
      warnings.push('before_meeting_start');
    } else {
      normalizedTimeMs = Math.max(0, normalizedTimeMs);
      status = 'ready_to_insert_current_axis';
      action = 'insert_mark';
    }
  } else if (axis === 'active_axis' && meeting.start_time_ms == null) {
    status = 'ready_to_insert_current_axis';
    action = 'insert_mark_without_relative_time';
    warnings.push('current_axis_missing_start_time');
  } else if (axis === 'ended_axis' && meeting.start_time_ms != null) {
    normalizedTimeMs = captured.value - meeting.start_time_ms;
    afterMeetingEndMs = meeting.end_time_ms == null ? 0 : Math.max(0, captured.value - meeting.end_time_ms);
    if (afterMeetingEndMs > 0) {
      status = 'after_meeting_end';
      action = 'store_for_audit_only';
      warnings.push('normalized_after_meeting_end');
    } else if (normalizedTimeMs < -policy.allowed_before_start_ms) {
      status = 'before_meeting_start';
      action = 'store_for_audit_only';
      warnings.push('before_meeting_start');
    } else {
      normalizedTimeMs = Math.max(0, normalizedTimeMs);
      status = 'ready_to_insert_current_axis';
      action = 'insert_mark';
    }
  } else if (axis === 'pending_axis') {
    status = 'pending_real_meeting';
    action = 'store_pending_mark';
    warnings.push('pending_real_meeting');
  } else if (axis === 'local_simulation_axis' && !policy.allow_local_simulation_axis) {
    status = 'local_simulation_axis';
    action = 'store_for_audit_only';
    warnings.push('not_on_real_axis');
  } else if (policy.allow_open_session_when_no_axis && ['no_axis', 'demo_axis', 'local_simulation_axis'].includes(axis)) {
    status = 'start_open_session_then_insert';
    action = 'start_open_session_then_insert_mark';
  } else if (policy.allow_pending_when_no_axis) {
    status = 'pending_real_meeting';
    action = 'store_pending_mark';
    warnings.push('pending_real_meeting');
  } else {
    status = 'unbound_no_axis';
    action = 'do_not_insert';
    warnings.push('not_on_real_axis');
  }

  const acceptedForRealtime = status === 'ready_to_insert_current_axis' || status === 'start_open_session_then_insert' || status === 'pending_real_meeting';
  return compactObject({
    type: 'meeting_platform_annotation_intake',
    schema: MEETING_PLATFORM_ANNOTATION_INTAKE_SCHEMA,
    schema_version: MEETING_PLATFORM_ANNOTATION_INTAKE_SCHEMA_VERSION,
    platform: key,
    status,
    action,
    accepted_for_realtime: acceptedForRealtime,
    timing_reliable: captured.value != null,
    requires_device_captured_at: captured.value == null,
    axis_state: axis,
    current_meeting: meeting,
    captured_at_ms: captured.value,
    captured_at_source: captured.source,
    normalized_time_ms: normalizedTimeMs == null ? undefined : Math.round(normalizedTimeMs),
    after_meeting_end_ms: afterMeetingEndMs,
    warnings,
    annotation: safeBuildTimelineMark(annotation, captured.value),
    insert_payload: captured.value == null ? undefined : buildInsertPayload(annotation, captured, normalizedTimeMs),
    open_session_payload: status === 'start_open_session_then_insert'
      ? buildOpenSessionPayload(key, annotation, meeting, captured, policy)
      : undefined,
    route_policy: policy,
    next_actions: nextActionsFor(status),
  });
}
