import { compactObject } from '../index.mjs';
import { normalizeMeetingSignal } from './core.mjs';
import { observeActiveSpeakerSamples } from './active-speaker.mjs';
import { buildMeetingPlatformAdaptationStrategy } from './platform-strategy.mjs';
import {
  MEETING_PLATFORM_KEYS,
  buildPlatformIntegrationPlan,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_SPEAKER_TRACK_PLAN_SCHEMA = 'meeting_platform_speaker_track_plan';
export const MEETING_PLATFORM_SPEAKER_TRACK_MATRIX_SCHEMA = 'meeting_platform_speaker_track_matrix';
export const MEETING_PLATFORM_SPEAKER_TRACK_SCHEMA = 'meeting_platform_speaker_track';
export const MEETING_PLATFORM_SPEAKER_TRACK_SCHEMA_VERSION = 1;

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

function boolOption(options, keys, fallback) {
  const value = firstNonEmpty(...keys.map((key) => options[key]));
  return value == null ? fallback : value !== false && value !== 'false';
}

function selectedPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_PLATFORM_KEYS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

function speakerFilterPolicy(options = {}) {
  return {
    min_stable_ms: numberOption(options, ['minStableMs', 'min_stable_ms'], 500),
    switch_stable_ms: numberOption(options, ['switchStableMs', 'switch_stable_ms'], 700),
    end_idle_ms: numberOption(options, ['endIdleMs', 'end_idle_ms', 'silenceEndMs', 'silence_end_ms'], 1_500),
    min_segment_ms: numberOption(options, ['minSegmentMs', 'min_segment_ms'], 800),
    merge_gap_ms: numberOption(options, ['mergeGapMs', 'merge_gap_ms'], 1_200),
    duplicate_window_ms: numberOption(options, ['duplicateWindowMs', 'duplicate_window_ms'], 1_200),
    emit_speaker_end: boolOption(options, ['emitSpeakerEnd', 'emit_speaker_end'], true),
    output_mark_kind: String(firstNonEmpty(options.outputMarkKind, options.output_mark_kind, 'speaker_started')),
  };
}

function sourceName(options = {}) {
  return String(firstNonEmpty(options.source, options.detectorSource, options.detector_source, 'local_active_speaker_observer'));
}

function speakerSourceCatalog(platform) {
  const key = normalizeMeetingPlatform(platform);
  return [
    {
      id: 'local_active_speaker_observer',
      role: 'primary_realtime_speaker_track',
      realtime: true,
      required: true,
      input: 'active_speaker_samples',
      notes: [
        'Use browser DOM, embedded WebView, native accessibility, or host SDK snapshots.',
        'Apply stable-window filtering before inserting timeline marks.',
      ],
    },
    {
      id: 'host_speaker_signal',
      role: 'primary_when_host_has_direct_audio_or_ui_state',
      realtime: true,
      required: false,
      input: 'speaker_started_or_speaker_ended_signal',
      notes: ['Use this when the meeting host already emits debounced speaker_started/speaker_ended signals.'],
    },
    {
      id: 'provider_participant_events',
      role: 'presence_reconcile_not_active_speaker',
      realtime: false,
      required: false,
      input: key === 'microsoft_teams'
        ? 'meetingCallEvents.rosterUpdated'
        : key === 'zoom'
          ? 'meeting.participant_joined/left'
          : key === 'webex'
            ? 'meetingParticipants.joined/left'
            : key === 'google_meet'
              ? 'participant.v2.joined/left'
              : 'join/leave_events',
      notes: ['Provider participant events can help audit participants, but they do not reliably identify current speaker position.'],
    },
    {
      id: 'post_meeting_transcript_speaker_labels',
      role: 'post_meeting_backfill_only',
      realtime: false,
      required: false,
      input: 'transcript_entries_or_artifact',
      notes: ['Transcript speaker labels must not block live annotation insertion.'],
    },
  ];
}

function normalizeSignalArray(input = {}, options = {}) {
  const defaults = options.defaults ?? input.defaults ?? {};
  return asArray(input.signals ?? input.events ?? input)
    .map((signal) => {
      try {
        return normalizeMeetingSignal(signal, defaults);
      } catch {
        return null;
      }
    })
    .filter((signal) => signal && (signal.type === 'speaker_started' || signal.type === 'speaker_ended'));
}

function sampleArray(input = {}) {
  if (Array.isArray(input)) return input;
  return asArray(firstNonEmpty(
    input.samples,
    input.activeSpeakerSamples,
    input.active_speaker_samples,
    input.observations,
    input.rows,
    [],
  ));
}

function speakerKey(signal = {}) {
  return firstNonEmpty(signal.speaker_id, signal.speaker_name, signal.participant_id, signal.participant_name);
}

function sameSpeaker(left = {}, right = {}) {
  const leftKey = speakerKey(left);
  const rightKey = speakerKey(right);
  return Boolean(leftKey && rightKey && String(leftKey) === String(rightKey));
}

function segmentId(segment = {}) {
  return [
    segment.meeting?.platform,
    segment.meeting?.meeting_id,
    speakerKey(segment),
    segment.start_at_ms,
  ].filter(Boolean).join(':');
}

function signalSortValue(signal = {}) {
  return Number(signal.occurred_at_ms ?? 0);
}

function closeSegment(active, endedAtMs, endSignal) {
  if (!active) return null;
  return compactObject({
    ...active,
    end_at_ms: endedAtMs,
    duration_ms: endedAtMs == null ? undefined : Math.max(0, endedAtMs - active.start_at_ms),
    end_source_event_id: endSignal?.source_event_id,
  });
}

function startSegment(signal) {
  return compactObject({
    meeting: signal.meeting,
    platform: signal.meeting?.platform,
    meeting_id: signal.meeting?.meeting_id,
    speaker_id: signal.speaker_id,
    speaker_name: signal.speaker_name,
    participant_id: signal.participant_id,
    participant_name: signal.participant_name,
    start_at_ms: signal.occurred_at_ms,
    start_source_event_id: signal.source_event_id,
    source: signal.source,
    raw: signal.raw,
  });
}

function mergeSegments(segments = [], policy = {}) {
  const merged = [];
  let mergedCount = 0;
  for (const segment of segments) {
    const previous = merged[merged.length - 1];
    const gap = previous?.end_at_ms == null ? Infinity : segment.start_at_ms - previous.end_at_ms;
    if (previous && sameSpeaker(previous, segment) && gap >= 0 && gap <= policy.merge_gap_ms) {
      merged[merged.length - 1] = compactObject({
        ...previous,
        end_at_ms: segment.end_at_ms,
        duration_ms: segment.end_at_ms == null ? undefined : segment.end_at_ms - previous.start_at_ms,
        end_source_event_id: segment.end_source_event_id,
      });
      mergedCount += 1;
    } else {
      merged.push(segment);
    }
  }
  return { segments: merged, merged_count: mergedCount };
}

function segmentsFromSignals(signals = [], policy = {}, options = {}) {
  const ordered = [...signals].sort((left, right) => signalSortValue(left) - signalSortValue(right));
  const rawSegments = [];
  let active = null;
  let duplicateCount = 0;
  let orphanEndCount = 0;
  let implicitEndCount = 0;
  let lastStart = null;

  for (const signal of ordered) {
    if (signal.type === 'speaker_started') {
      if (
        lastStart
        && sameSpeaker(lastStart, signal)
        && Math.abs(signal.occurred_at_ms - lastStart.occurred_at_ms) <= policy.duplicate_window_ms
      ) {
        duplicateCount += 1;
        continue;
      }
      if (active) {
        rawSegments.push(closeSegment(active, signal.occurred_at_ms, signal));
        implicitEndCount += 1;
      }
      active = startSegment(signal);
      lastStart = signal;
      continue;
    }
    if (signal.type === 'speaker_ended') {
      if (!active || !sameSpeaker(active, signal)) {
        orphanEndCount += 1;
        continue;
      }
      rawSegments.push(closeSegment(active, signal.occurred_at_ms, signal));
      active = null;
    }
  }

  const closeAt = firstNonEmpty(options.closeOpenSegmentsAtMs, options.close_open_segments_at_ms);
  if (active && closeAt != null) {
    rawSegments.push(closeSegment(active, Number(closeAt), undefined));
    active = null;
  }

  const { segments: mergedSegments, merged_count } = mergeSegments(rawSegments, policy);
  const kept = [];
  let droppedShortCount = 0;
  for (const segment of mergedSegments) {
    if (segment.duration_ms != null && segment.duration_ms < policy.min_segment_ms) {
      droppedShortCount += 1;
      continue;
    }
    kept.push(compactObject({
      id: segmentId(segment),
      ...segment,
    }));
  }

  return {
    segments: kept,
    active_segment: active ? compactObject({ id: segmentId(active), ...active }) : undefined,
    diagnostics: {
      input_signal_count: ordered.length,
      raw_segment_count: rawSegments.length,
      merged_segment_count: mergedSegments.length,
      kept_segment_count: kept.length,
      dropped_duplicate_count: duplicateCount,
      dropped_short_segment_count: droppedShortCount,
      merged_gap_count: merged_count,
      orphan_end_count: orphanEndCount,
      implicit_end_count: implicitEndCount,
      open_segment_count: active ? 1 : 0,
    },
  };
}

function speakerName(segment = {}) {
  return firstNonEmpty(segment.speaker_name, segment.speaker_id, segment.participant_name, segment.participant_id, 'speaker');
}

function segmentToMark(segment = {}, options = {}) {
  const kind = String(firstNonEmpty(options.outputMarkKind, options.output_mark_kind, 'speaker_started'));
  return compactObject({
    id: String(firstNonEmpty(options.idPrefix, options.id_prefix, segment.id)),
    source: `${segment.platform ?? 'meeting'}_speaker_track`,
    captured_at_ms: segment.start_at_ms,
    kind,
    label: `${speakerName(segment)} speaking`,
    intent: 'speaker_track',
    payload: {
      meeting: segment.meeting,
      speaker_id: segment.speaker_id,
      speaker_name: segment.speaker_name,
      participant_id: segment.participant_id,
      participant_name: segment.participant_name,
      segment: {
        start_at_ms: segment.start_at_ms,
        end_at_ms: segment.end_at_ms,
        duration_ms: segment.duration_ms,
      },
      source_event_ids: unique([segment.start_source_event_id, segment.end_source_event_id]),
    },
  });
}

export function buildMeetingPlatformSpeakerTrackPlan(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const strategy = buildMeetingPlatformAdaptationStrategy(key, options);
  const integration = buildPlatformIntegrationPlan(key, options);
  return compactObject({
    type: 'meeting_platform_speaker_track_plan',
    schema: MEETING_PLATFORM_SPEAKER_TRACK_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_SPEAKER_TRACK_SCHEMA_VERSION,
    platform: key,
    display_name: strategy.display_name ?? integration.display_name,
    status: 'host_speaker_observer_required',
    realtime_ready_when_samples_available: true,
    provider_events_block_realtime: false,
    transcript_blocks_realtime: false,
    track_source_order: speakerSourceCatalog(key),
    filter_policy: speakerFilterPolicy(options),
    output_contract: {
      track: 'speaker',
      mark_intent: 'speaker_track',
      default_mark_kind: 'speaker_started',
      timestamp_field: 'captured_at_ms',
      transcript_required: false,
      provider_event_required: false,
      inserts_content_text: false,
    },
    required_input_fields: ['meeting_id', 'captured_at_ms_or_observed_at_ms', 'speaker_id_or_speaker_name'],
    next_actions: [
      'capture_active_speaker_samples_from_host_observer',
      'apply_stable_window_and_duplicate_filter_before_insert',
      'insert_only_speaker_position_markers_without_transcript_content',
    ],
  });
}

export function buildMeetingPlatformSpeakerTrackMatrix(options = {}) {
  const plans = selectedPlatforms(options).map((platform) => buildMeetingPlatformSpeakerTrackPlan(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_speaker_track_matrix',
    schema: MEETING_PLATFORM_SPEAKER_TRACK_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_SPEAKER_TRACK_SCHEMA_VERSION,
    platform_count: plans.length,
    realtime_ready_when_samples_available_count: plans.filter((plan) => plan.realtime_ready_when_samples_available).length,
    provider_blocking_count: plans.filter((plan) => plan.provider_events_block_realtime).length,
    transcript_blocking_count: plans.filter((plan) => plan.transcript_blocks_realtime).length,
    platforms: plans.map((plan) => plan.platform),
    rows: plans.map((plan) => ({
      platform: plan.platform,
      display_name: plan.display_name,
      status: plan.status,
      source_order: plan.track_source_order.map((source) => source.id),
      min_stable_ms: plan.filter_policy.min_stable_ms,
      switch_stable_ms: plan.filter_policy.switch_stable_ms,
      end_idle_ms: plan.filter_policy.end_idle_ms,
      provider_events_block_realtime: plan.provider_events_block_realtime,
      transcript_blocks_realtime: plan.transcript_blocks_realtime,
      next_actions: plan.next_actions,
    })),
    plans,
    next_actions: unique(plans.flatMap((plan) => plan.next_actions ?? [])),
  };
}

export function buildMeetingPlatformSpeakerTrack(platform, input = {}, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const policy = speakerFilterPolicy(options);
  const samples = sampleArray(input);
  const observed = samples.length > 0
    ? observeActiveSpeakerSamples(options.initialState ?? options.initial_state ?? null, samples, {
      minStableMs: policy.min_stable_ms,
      switchStableMs: policy.switch_stable_ms,
      endIdleMs: policy.end_idle_ms,
      emitSpeakerEnd: policy.emit_speaker_end,
      source: sourceName(options),
      ...options.observerOptions,
      ...options.observer_options,
    })
    : { state: options.initialState ?? options.initial_state ?? null, signals: [], observations: [] };
  const explicitSignals = normalizeSignalArray(input, options);
  const signals = [...observed.signals, ...explicitSignals];
  const segmentResult = segmentsFromSignals(signals, policy, options);
  const marks = segmentResult.segments.map((segment) => segmentToMark(segment, {
    outputMarkKind: policy.output_mark_kind,
  }));
  return compactObject({
    type: 'meeting_platform_speaker_track',
    schema: MEETING_PLATFORM_SPEAKER_TRACK_SCHEMA,
    schema_version: MEETING_PLATFORM_SPEAKER_TRACK_SCHEMA_VERSION,
    platform: key,
    status: marks.length > 0 ? 'speaker_track_ready' : 'no_speaker_markers',
    input_sample_count: samples.length,
    normalized_signal_count: explicitSignals.length,
    generated_signal_count: observed.signals.length,
    signal_count: signals.length,
    segment_count: segmentResult.segments.length,
    mark_count: marks.length,
    filter_policy: policy,
    signals,
    segments: segmentResult.segments,
    active_segment: segmentResult.active_segment,
    marks,
    observer_state: observed.state,
    diagnostics: {
      observation_count: observed.observations.length,
      ...segmentResult.diagnostics,
    },
    next_actions: marks.length > 0
      ? ['insert_speaker_track_marks_on_current_axis']
      : ['collect_more_stable_active_speaker_samples'],
  });
}
