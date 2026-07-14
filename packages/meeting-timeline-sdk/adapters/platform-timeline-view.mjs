import {
  compactObject,
  normalizeAbsoluteMs,
} from './internal-utils.mjs';
import {
  MEETING_PLATFORM_KEYS,
  buildPlatformIntegrationPlan,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_TIMELINE_VIEW_PLAN_SCHEMA = 'meeting_platform_timeline_view_plan';
export const MEETING_PLATFORM_TIMELINE_VIEW_MATRIX_SCHEMA = 'meeting_platform_timeline_view_matrix';
export const MEETING_PLATFORM_TIMELINE_VIEW_SCHEMA = 'meeting_platform_timeline_view';
export const MEETING_PLATFORM_TIMELINE_VIEW_SCHEMA_VERSION = 1;

const DEFAULT_FULL_DURATION_MS = 10 * 60 * 1000;
const DEFAULT_MIN_VIEW_MS = 60 * 1000;
const DEFAULT_RAILS = Object.freeze([
  { id: 'speaker', label: 'Speaker', role: 'who_is_speaking', order: 10 },
  { id: 'participant', label: 'Participants', role: 'who_joined_or_left', order: 20 },
  { id: 'events', label: 'Events', role: 'meeting_and_provider_events', order: 30 },
  { id: 'annotations', label: 'Annotations', role: 'human_marks', order: 40 },
  { id: 'artifacts', label: 'Artifacts', role: 'post_meeting_outputs', order: 50 },
  { id: 'transcript', label: 'Transcript', role: 'post_meeting_transcript', order: 60 },
]);

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

function normalizeMeeting(input = {}, platform, options = {}) {
  const meeting = input.meeting ?? input.session ?? {};
  const key = normalizeMeetingPlatform(firstNonEmpty(meeting.platform, input.platform, platform));
  const start = maybeAbsoluteMs(firstNonEmpty(
    meeting.start_time_ms,
    meeting.startTimeMs,
    meeting.started_at_ms,
    meeting.startedAtMs,
    meeting.start_time,
    meeting.startTime,
    meeting.started_at,
    meeting.startedAt,
    input.start_time_ms,
    input.startTimeMs,
  ), 'meeting.start_time_ms');
  const end = maybeAbsoluteMs(firstNonEmpty(
    meeting.end_time_ms,
    meeting.endTimeMs,
    meeting.ended_at_ms,
    meeting.endedAtMs,
    meeting.end_time,
    meeting.endTime,
    meeting.ended_at,
    meeting.endedAt,
    input.end_time_ms,
    input.endTimeMs,
  ), 'meeting.end_time_ms');
  const configuredDuration = Number(firstNonEmpty(
    meeting.duration_ms,
    meeting.durationMs,
    input.duration_ms,
    input.durationMs,
    options.fullDurationMs,
    options.full_duration_ms,
    options.defaultDurationMs,
    options.default_duration_ms,
  ));
  const duration = Number.isFinite(configuredDuration) && configuredDuration > 0
    ? configuredDuration
    : end != null && start != null && end > start
      ? end - start
      : DEFAULT_FULL_DURATION_MS;
  return compactObject({
    platform: key,
    meeting_id: firstNonEmpty(meeting.meeting_id, meeting.meetingId, meeting.id, input.meeting_id, input.meetingId),
    external_meeting_id: firstNonEmpty(meeting.external_meeting_id, meeting.externalMeetingId),
    meeting_url: firstNonEmpty(meeting.meeting_url, meeting.meetingUrl, meeting.url),
    title: firstNonEmpty(meeting.title, meeting.topic, meeting.name),
    start_time_ms: start,
    end_time_ms: end,
    duration_ms: Math.max(1, Math.round(duration)),
  });
}

function markerRail(item = {}) {
  const explicit = firstNonEmpty(item.rail, item.track, item.payload?.rail, item.payload?.track);
  if (explicit) return String(explicit);
  const intent = String(firstNonEmpty(item.intent, item.payload?.intent, '')).toLowerCase();
  const kind = String(firstNonEmpty(item.kind, item.type, '')).toLowerCase();
  if (intent.includes('speaker') || kind.includes('speaker')) return 'speaker';
  if (intent.includes('participant') || kind.includes('participant')) return 'participant';
  if (
    item.artifact_kind
    || item.artifactKind
    || intent.includes('artifact')
    || kind.includes('artifact')
    || kind.includes('transcript')
    || kind.includes('recording')
  ) return 'artifacts';
  if (intent.includes('event') || kind.includes('meeting_') || kind.includes('subscription')) return 'events';
  if (intent.includes('transcript')) return 'transcript';
  return 'annotations';
}

function markerAbsoluteMs(item = {}) {
  return maybeAbsoluteMs(firstNonEmpty(
    item.captured_at_ms,
    item.capturedAtMs,
    item.occurred_at_ms,
    item.occurredAtMs,
    item.timestamp_ms,
    item.timestampMs,
    item.payload?.timing?.captured_at_ms,
    item.payload?.timing?.capturedAtMs,
  ), 'marker.captured_at_ms');
}

function markerTimeMs(item = {}, meeting = {}) {
  const explicit = Number(firstNonEmpty(
    item.time_ms,
    item.timeMs,
    item.offset_ms,
    item.offsetMs,
    item.relative_ms,
    item.relativeMs,
    item.payload?.timing?.normalized_time_ms,
  ));
  if (Number.isFinite(explicit)) return explicit;
  const absolute = markerAbsoluteMs(item);
  if (absolute != null && meeting.start_time_ms != null) return absolute - meeting.start_time_ms;
  return undefined;
}

function markerLabel(item = {}) {
  return String(firstNonEmpty(
    item.label,
    item.title,
    item.text,
    item.kind,
    item.type,
    item.artifact_kind ? `${item.artifact_kind} ready` : undefined,
    'marker',
  ));
}

function normalizeMarker(item = {}, source, meeting = {}) {
  const timeMs = markerTimeMs(item, meeting);
  const absolute = markerAbsoluteMs(item);
  const rail = markerRail(item);
  const warnings = [];
  if (!Number.isFinite(timeMs)) warnings.push('missing_relative_time');
  if (Number.isFinite(timeMs) && timeMs < 0) warnings.push('before_meeting_start');
  if (
    Number.isFinite(timeMs)
    && meeting.end_time_ms != null
    && meeting.start_time_ms != null
    && timeMs > meeting.end_time_ms - meeting.start_time_ms
  ) warnings.push('after_meeting_end');
  return compactObject({
    id: String(firstNonEmpty(item.id, item.source_event_id, item.sourceEventId, `${source}-${rail}-${absolute ?? timeMs ?? markerLabel(item)}`)),
    rail,
    source,
    kind: firstNonEmpty(item.kind, item.type, rail),
    label: markerLabel(item),
    captured_at_ms: absolute,
    time_ms: Number.isFinite(timeMs) ? Math.round(timeMs) : undefined,
    calibrated: Number.isFinite(timeMs) && timeMs >= 0,
    warnings,
    raw: item,
  });
}

function normalizeEventMarker(signal = {}, source, meeting = {}) {
  return normalizeMarker({
    ...signal,
    kind: signal.type,
    label: String(signal.type ?? 'event').replace(/_/g, ' '),
  }, source, meeting);
}

function collectMarkers(input = {}, meeting = {}) {
  const tracks = input.tracks ?? {};
  const rows = [
    ...asArray(input.annotations).map((item) => normalizeMarker(item, 'annotation', meeting)),
    ...asArray(input.sequence).map((item) => normalizeMarker(item, 'annotation', meeting)),
    ...asArray(input.marks).map((item) => normalizeMarker(item, 'annotation', meeting)),
    ...asArray(input.events).map((item) => normalizeEventMarker(item, 'event', meeting)),
    ...asArray(input.signals).map((item) => normalizeEventMarker(item, 'event', meeting)),
    ...asArray(input.speakerTrack?.marks ?? input.speaker_track?.marks ?? tracks.speaker?.marks).map((item) => normalizeMarker(item, 'speaker_track', meeting)),
    ...asArray(input.participantTrack?.marks ?? input.participant_track?.marks ?? tracks.participant?.marks).map((item) => normalizeMarker(item, 'participant_track', meeting)),
    ...asArray(input.artifactHandoff?.rows ?? input.artifact_handoff?.rows ?? tracks.artifacts?.rows).map((item) => normalizeMarker(item, 'artifact_handoff', meeting)),
    ...asArray(input.transcript?.segments ?? input.transcript_segments ?? tracks.transcript?.segments).map((item) => normalizeMarker({
      ...item,
      rail: 'transcript',
      kind: 'transcript_segment',
      label: firstNonEmpty(item.text, item.label, item.speaker_name, 'transcript'),
      time_ms: firstNonEmpty(item.start_ms, item.startMs, item.offset_ms, item.offsetMs),
    }, 'transcript', meeting)),
  ];
  return rows
    .filter(Boolean)
    .sort((left, right) => {
      const leftTime = Number.isFinite(left.time_ms) ? left.time_ms : Infinity;
      const rightTime = Number.isFinite(right.time_ms) ? right.time_ms : Infinity;
      return leftTime - rightTime || String(left.rail).localeCompare(String(right.rail));
    });
}

function timelineDuration(meeting = {}, markers = [], options = {}) {
  const markerMax = Math.max(0, ...markers
    .map((marker) => Number(marker.time_ms))
    .filter((value) => Number.isFinite(value) && value >= 0));
  const configured = numberOption(options, ['fullDurationMs', 'full_duration_ms'], meeting.duration_ms);
  return Math.max(1, Math.ceil(Math.max(configured, markerMax + numberOption(options, ['tailPaddingMs', 'tail_padding_ms'], 30_000))));
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function buildViewport(fullDurationMs, options = {}) {
  const minView = Math.min(fullDurationMs, numberOption(options, ['minViewMs', 'min_view_ms'], DEFAULT_MIN_VIEW_MS));
  const duration = clamp(numberOption(options, ['viewportDurationMs', 'viewport_duration_ms', 'windowDurationMs', 'window_duration_ms'], fullDurationMs), minView, fullDurationMs);
  const maxStart = Math.max(0, fullDurationMs - duration);
  const start = clamp(numberOption(options, ['viewportStartMs', 'viewport_start_ms', 'windowStartMs', 'window_start_ms'], 0), 0, maxStart);
  return {
    start_ms: Math.round(start),
    end_ms: Math.round(start + duration),
    duration_ms: Math.round(duration),
    full_duration_ms: Math.round(fullDurationMs),
    min_view_ms: Math.round(minView),
    is_full: start <= 0 && duration >= fullDurationMs,
  };
}

function tickStep(durationMs) {
  return [
    60 * 1000,
    2 * 60 * 1000,
    5 * 60 * 1000,
    10 * 60 * 1000,
    15 * 60 * 1000,
    30 * 60 * 1000,
    60 * 60 * 1000,
    2 * 60 * 60 * 1000,
  ].find((candidate) => durationMs / candidate <= 8) ?? 4 * 60 * 60 * 1000;
}

function formatOffset(ms) {
  const totalSeconds = Math.max(0, Math.round(Number(ms) / 1000));
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
  const two = (value) => String(value).padStart(2, '0');
  return hours > 0 ? `${hours}:${two(minutes)}:${two(seconds)}` : `${two(minutes)}:${two(seconds)}`;
}

function buildTicks(viewport = {}) {
  const ticks = [viewport.start_ms];
  const step = tickStep(viewport.duration_ms);
  const first = Math.ceil(viewport.start_ms / step) * step;
  for (let ms = first; ms <= viewport.end_ms; ms += step) ticks.push(ms);
  ticks.push(viewport.end_ms);
  return [...new Set(ticks.map((ms) => Math.round(ms)))]
    .sort((left, right) => left - right)
    .map((ms) => ({ ms, label: formatOffset(ms) }));
}

function positionMarkers(markers = [], viewport = {}) {
  const duration = Math.max(1, viewport.duration_ms);
  return markers.map((marker) => {
    const visible = Number.isFinite(marker.time_ms)
      && marker.time_ms >= viewport.start_ms
      && marker.time_ms <= viewport.end_ms;
    const xRatio = Number.isFinite(marker.time_ms)
      ? clamp((marker.time_ms - viewport.start_ms) / duration, 0, 1)
      : undefined;
    return compactObject({
      ...marker,
      visible,
      x_ratio: xRatio == null ? undefined : Number(xRatio.toFixed(6)),
    });
  });
}

function buildRails(markers = [], options = {}) {
  const configured = asArray(options.rails ?? options.timelineRails ?? options.timeline_rails);
  const rails = configured.length > 0 ? configured : DEFAULT_RAILS;
  const counts = markers.reduce((acc, marker) => {
    acc[marker.rail] = (acc[marker.rail] ?? 0) + 1;
    return acc;
  }, {});
  return rails.map((rail, index) => compactObject({
    id: String(rail.id ?? rail),
    label: String(rail.label ?? rail.id ?? rail),
    role: rail.role,
    order: Number(rail.order ?? (index + 1) * 10),
    marker_count: counts[String(rail.id ?? rail)] ?? 0,
  })).sort((left, right) => left.order - right.order);
}

export function buildMeetingPlatformTimelineViewPlan(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const integration = buildPlatformIntegrationPlan(key, options);
  return compactObject({
    type: 'meeting_platform_timeline_view_plan',
    schema: MEETING_PLATFORM_TIMELINE_VIEW_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_TIMELINE_VIEW_SCHEMA_VERSION,
    platform: key,
    display_name: integration.display_name,
    status: 'timeline_view_contract_ready',
    rails: buildRails([], options),
    input_contract: {
      meeting_required: true,
      supported_tracks: ['annotations', 'speakerTrack', 'participantTrack', 'artifactHandoff', 'transcript', 'events'],
      timestamp_fields: ['captured_at_ms', 'occurred_at_ms', 'time_ms', 'start_ms'],
      annotation_timestamp_field: 'captured_at_ms',
    },
    output_contract: {
      viewport: 'relative_ms_window',
      marker_position: 'x_ratio_0_to_1',
      renderer_agnostic: true,
    },
    realtime_policy: {
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      view_building_is_local_only: true,
    },
    next_actions: [
      'feed_current_axis_and_track_marks_into_timeline_view',
      'render_rails_with_host_ui',
      'use_viewport_start_and_duration_for_zoom_or_pan',
    ],
  });
}

export function buildMeetingPlatformTimelineViewMatrix(options = {}) {
  const plans = selectedPlatforms(options).map((platform) => buildMeetingPlatformTimelineViewPlan(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_timeline_view_matrix',
    schema: MEETING_PLATFORM_TIMELINE_VIEW_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_TIMELINE_VIEW_SCHEMA_VERSION,
    platform_count: plans.length,
    renderer_agnostic_count: plans.filter((plan) => plan.output_contract.renderer_agnostic).length,
    provider_blocking_count: plans.filter((plan) => plan.realtime_policy.provider_events_block_realtime).length,
    transcript_blocking_count: plans.filter((plan) => plan.realtime_policy.transcript_blocks_realtime).length,
    platforms: plans.map((plan) => plan.platform),
    rows: plans.map((plan) => ({
      platform: plan.platform,
      display_name: plan.display_name,
      status: plan.status,
      rails: plan.rails.map((rail) => rail.id),
      provider_events_block_realtime: plan.realtime_policy.provider_events_block_realtime,
      transcript_blocks_realtime: plan.realtime_policy.transcript_blocks_realtime,
    })),
    plans,
    next_actions: unique(plans.flatMap((plan) => plan.next_actions ?? [])),
  };
}

export function buildMeetingPlatformTimelineView(platform, input = {}, options = {}) {
  const meeting = normalizeMeeting(input, platform, options);
  const markers = collectMarkers(input, meeting);
  const fullDurationMs = timelineDuration(meeting, markers, options);
  const viewport = buildViewport(fullDurationMs, options);
  const positionedMarkers = positionMarkers(markers, viewport);
  const rails = buildRails(positionedMarkers, options);
  const visibleMarkers = positionedMarkers.filter((marker) => marker.visible);
  const uncalibratedMarkers = positionedMarkers.filter((marker) => !marker.calibrated);
  return compactObject({
    type: 'meeting_platform_timeline_view',
    schema: MEETING_PLATFORM_TIMELINE_VIEW_SCHEMA,
    schema_version: MEETING_PLATFORM_TIMELINE_VIEW_SCHEMA_VERSION,
    platform: meeting.platform,
    status: markers.length > 0 ? 'timeline_view_ready' : 'empty_timeline_view',
    meeting,
    viewport,
    ticks: buildTicks(viewport),
    rails,
    markers: positionedMarkers,
    visible_markers: visibleMarkers,
    uncalibrated_markers: uncalibratedMarkers,
    diagnostics: {
      marker_count: positionedMarkers.length,
      visible_marker_count: visibleMarkers.length,
      uncalibrated_marker_count: uncalibratedMarkers.length,
      warning_count: positionedMarkers.filter((marker) => marker.warnings?.length > 0).length,
      rail_counts: Object.fromEntries(rails.map((rail) => [rail.id, rail.marker_count])),
    },
    next_actions: markers.length > 0
      ? ['render_visible_markers_by_rail', 'keep_captured_at_ms_for_realtime_annotations']
      : ['wait_for_axis_or_track_mark_input'],
  });
}

export function zoomMeetingPlatformTimelineViewport(viewport = {}, factor = 1, anchorRatio = 0.5, options = {}) {
  const full = Math.max(1, Number(viewport.full_duration_ms ?? viewport.fullDurationMs ?? viewport.duration_ms ?? DEFAULT_FULL_DURATION_MS));
  const currentDuration = Math.max(1, Number(viewport.duration_ms ?? viewport.durationMs ?? full));
  const currentStart = Math.max(0, Number(viewport.start_ms ?? viewport.startMs ?? 0));
  const nextDuration = currentDuration / Math.max(0.001, Number(factor) || 1);
  const anchor = currentStart + currentDuration * clamp(Number(anchorRatio), 0, 1);
  return buildViewport(full, {
    ...options,
    viewportStartMs: anchor - nextDuration * clamp(Number(anchorRatio), 0, 1),
    viewportDurationMs: nextDuration,
  });
}
