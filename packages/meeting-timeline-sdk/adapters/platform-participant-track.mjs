import { compactObject } from '../index.mjs';
import { normalizeMeetingSignal } from './core.mjs';
import { buildMeetingPlatformAdaptationStrategy } from './platform-strategy.mjs';
import {
  MEETING_PLATFORM_KEYS,
  buildPlatformIntegrationPlan,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_PARTICIPANT_TRACK_PLAN_SCHEMA = 'meeting_platform_participant_track_plan';
export const MEETING_PLATFORM_PARTICIPANT_TRACK_MATRIX_SCHEMA = 'meeting_platform_participant_track_matrix';
export const MEETING_PLATFORM_PARTICIPANT_TRACK_SCHEMA = 'meeting_platform_participant_track';
export const MEETING_PLATFORM_PARTICIPANT_TRACK_SCHEMA_VERSION = 1;

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

function participantFilterPolicy(options = {}) {
  return {
    duplicate_window_ms: numberOption(options, ['duplicateWindowMs', 'duplicate_window_ms'], 1_500),
    suppress_reconnect_gap_ms: numberOption(options, ['suppressReconnectGapMs', 'suppress_reconnect_gap_ms'], 10_000),
    leave_stable_ms: numberOption(options, ['leaveStableMs', 'leave_stable_ms'], 1_500),
    emit_initial_roster: boolOption(options, ['emitInitialRoster', 'emit_initial_roster'], false),
    output_join_kind: String(firstNonEmpty(options.outputJoinKind, options.output_join_kind, 'participant_joined')),
    output_left_kind: String(firstNonEmpty(options.outputLeftKind, options.output_left_kind, 'participant_left')),
  };
}

function sourceName(options = {}) {
  return String(firstNonEmpty(options.source, options.detectorSource, options.detector_source, 'local_roster_observer'));
}

function participantSourceCatalog(platform) {
  const key = normalizeMeetingPlatform(platform);
  return [
    {
      id: 'local_roster_observer',
      role: 'primary_realtime_participant_track',
      realtime: true,
      required: true,
      input: 'participant_roster_snapshots',
      notes: [
        'Use browser DOM, native accessibility, or host SDK roster snapshots.',
        'Apply duplicate and reconnect filters before inserting timeline marks.',
      ],
    },
    {
      id: 'host_participant_signal',
      role: 'primary_when_host_has_direct_join_leave_events',
      realtime: true,
      required: false,
      input: 'participant_joined_or_participant_left_signal',
      notes: ['Use this when the host already emits normalized participant_joined/participant_left signals.'],
    },
    {
      id: 'provider_participant_events',
      role: 'reconcile_and_backfill_only_when_delayed',
      realtime: false,
      required: false,
      input: key === 'microsoft_teams'
        ? 'meetingCallEvents.rosterUpdated_or_Teams_bot_participant_events'
        : key === 'zoom'
          ? 'meeting.participant_joined/left'
          : key === 'webex'
            ? 'meetingParticipants.joined/left'
            : key === 'google_meet'
              ? 'participant.v2.joined/left'
              : 'join/leave_events',
      notes: [
        'Provider participant events can be delayed or permission-gated.',
        'They must not be required for realtime annotation insertion.',
      ],
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
    .filter((signal) => signal && (signal.type === 'participant_joined' || signal.type === 'participant_left'));
}

function snapshotArray(input = {}) {
  if (Array.isArray(input)) return [];
  return asArray(firstNonEmpty(
    input.snapshots,
    input.participantSnapshots,
    input.participant_snapshots,
    input.rosterSnapshots,
    input.roster_snapshots,
    [],
  ));
}

function snapshotTime(snapshot = {}) {
  return Number(firstNonEmpty(
    snapshot.observed_at_ms,
    snapshot.observedAtMs,
    snapshot.occurred_at_ms,
    snapshot.occurredAtMs,
    snapshot.timestamp_ms,
    snapshot.timestampMs,
    snapshot.time_ms,
    snapshot.timeMs,
  ));
}

function participantRows(snapshot = {}) {
  return asArray(firstNonEmpty(
    snapshot.participants,
    snapshot.roster,
    snapshot.members,
    snapshot.people,
    [],
  ));
}

function participantName(participant = {}) {
  return firstNonEmpty(
    participant.participant_name,
    participant.participantName,
    participant.display_name,
    participant.displayName,
    participant.name,
    participant.label,
    participant.id,
    participant.participant_id,
    participant.participantId,
    'participant',
  );
}

function participantId(participant = {}) {
  return firstNonEmpty(
    participant.participant_id,
    participant.participantId,
    participant.id,
    participant.user_id,
    participant.userId,
    participant.email,
    participant.name,
    participant.display_name,
    participant.displayName,
  );
}

function participantKey(value = {}) {
  return firstNonEmpty(
    value.participant_id,
    value.participantId,
    value.id,
    value.user_id,
    value.userId,
    value.email,
    value.participant_name,
    value.participantName,
    value.display_name,
    value.displayName,
    value.name,
  );
}

function normalizeParticipant(participant = {}) {
  const id = participantId(participant);
  const name = participantName(participant);
  return compactObject({
    participant_id: id == null ? undefined : String(id),
    participant_name: name == null ? undefined : String(name),
    raw: participant.raw ?? participant,
  });
}

function signalParticipantKey(signal = {}) {
  return firstNonEmpty(signal.participant_id, signal.participant_name);
}

function signalSortValue(signal = {}) {
  return Number(signal.occurred_at_ms ?? 0);
}

function signalId(signal = {}) {
  return [
    signal.meeting?.platform,
    signal.meeting?.meeting_id,
    signal.type,
    signalParticipantKey(signal),
    signal.occurred_at_ms,
  ].filter(Boolean).join(':');
}

function makeParticipantSignal(type, snapshot = {}, participant = {}, occurredAtMs, options = {}) {
  const normalized = normalizeParticipant(participant);
  return compactObject({
    type,
    meeting: snapshot.meeting ?? options.defaults?.meeting,
    occurred_at_ms: occurredAtMs,
    source_event_id: firstNonEmpty(snapshot.source_event_id, snapshot.sourceEventId),
    source: sourceName(options),
    participant_id: normalized.participant_id,
    participant_name: normalized.participant_name,
    raw: normalized.raw,
  });
}

function normalizeSnapshotSignal(signal, defaults = {}) {
  try {
    return normalizeMeetingSignal(signal, defaults);
  } catch {
    return null;
  }
}

function signalsFromSnapshots(snapshots = [], policy = {}, options = {}) {
  const ordered = [...snapshots]
    .map((snapshot, index) => ({ snapshot, index, observed_at_ms: snapshotTime(snapshot) }))
    .filter((row) => Number.isFinite(row.observed_at_ms))
    .sort((left, right) => left.observed_at_ms - right.observed_at_ms || left.index - right.index);
  const signals = [];
  const present = new Map();
  const pendingLeaves = new Map();
  let initialized = false;
  let suppressedPendingLeaveCount = 0;
  let emittedPendingLeaveCount = 0;

  for (const row of ordered) {
    const current = new Map();
    for (const participant of participantRows(row.snapshot)) {
      const key = participantKey(participant);
      if (!key) continue;
      current.set(String(key), participant);
    }
    for (const [key, participant] of current) {
      if (pendingLeaves.has(key)) {
        pendingLeaves.delete(key);
        suppressedPendingLeaveCount += 1;
      }
      if (!present.has(key)) {
        if (initialized || policy.emit_initial_roster) {
          signals.push(makeParticipantSignal('participant_joined', row.snapshot, participant, row.observed_at_ms, options));
        }
        present.set(key, participant);
      } else {
        present.set(key, participant);
      }
    }
    initialized = true;
    for (const [key, participant] of present) {
      if (current.has(key) || pendingLeaves.has(key)) continue;
      pendingLeaves.set(key, {
        participant,
        snapshot: row.snapshot,
        left_at_ms: row.observed_at_ms,
      });
    }
    for (const [key, pending] of [...pendingLeaves.entries()]) {
      if (row.observed_at_ms - pending.left_at_ms < policy.leave_stable_ms) continue;
      signals.push(makeParticipantSignal('participant_left', pending.snapshot, pending.participant, pending.left_at_ms, options));
      pendingLeaves.delete(key);
      present.delete(key);
      emittedPendingLeaveCount += 1;
    }
  }

  const closeAt = firstNonEmpty(options.closePendingLeavesAtMs, options.close_pending_leaves_at_ms);
  if (closeAt != null) {
    const closeMs = Number(closeAt);
    for (const [key, pending] of [...pendingLeaves.entries()]) {
      if (!Number.isFinite(closeMs) || closeMs - pending.left_at_ms < policy.leave_stable_ms) continue;
      signals.push(makeParticipantSignal('participant_left', pending.snapshot, pending.participant, pending.left_at_ms, options));
      pendingLeaves.delete(key);
      present.delete(key);
      emittedPendingLeaveCount += 1;
    }
  }

  const normalized = signals
    .map((signal) => normalizeSnapshotSignal(signal, options.defaults ?? {}))
    .filter(Boolean);
  return {
    signals: normalized,
    diagnostics: {
      snapshot_count: ordered.length,
      generated_signal_count: normalized.length,
      suppressed_pending_leave_count: suppressedPendingLeaveCount,
      emitted_pending_leave_count: emittedPendingLeaveCount,
      pending_leave_count: pendingLeaves.size,
      observed_participant_count: present.size,
    },
  };
}

function filterDuplicateSignals(signals = [], policy = {}) {
  const ordered = [...signals].sort((left, right) => signalSortValue(left) - signalSortValue(right));
  const kept = [];
  let droppedDuplicateCount = 0;
  for (const signal of ordered) {
    const previous = kept[kept.length - 1];
    const sameKey = previous
      && previous.type === signal.type
      && signalParticipantKey(previous) === signalParticipantKey(signal)
      && previous.meeting?.meeting_id === signal.meeting?.meeting_id;
    const gap = previous ? signal.occurred_at_ms - previous.occurred_at_ms : Infinity;
    if (sameKey && gap >= 0 && gap <= policy.duplicate_window_ms) {
      droppedDuplicateCount += 1;
      continue;
    }
    kept.push(signal);
  }
  return { signals: kept, dropped_duplicate_count: droppedDuplicateCount };
}

function suppressReconnectPairs(signals = [], policy = {}) {
  const suppressed = new Set();
  let suppressedPairCount = 0;
  for (let index = 0; index < signals.length; index += 1) {
    if (suppressed.has(index) || signals[index].type !== 'participant_left') continue;
    const left = signals[index];
    const matchIndex = signals.findIndex((candidate, candidateIndex) => (
      candidateIndex > index
      && !suppressed.has(candidateIndex)
      && candidate.type === 'participant_joined'
      && candidate.meeting?.meeting_id === left.meeting?.meeting_id
      && signalParticipantKey(candidate) === signalParticipantKey(left)
      && candidate.occurred_at_ms - left.occurred_at_ms >= 0
      && candidate.occurred_at_ms - left.occurred_at_ms <= policy.suppress_reconnect_gap_ms
    ));
    if (matchIndex < 0) continue;
    suppressed.add(index);
    suppressed.add(matchIndex);
    suppressedPairCount += 1;
  }
  return {
    signals: signals.filter((_, index) => !suppressed.has(index)),
    suppressed_reconnect_pair_count: suppressedPairCount,
  };
}

function participantLabel(signal = {}) {
  return firstNonEmpty(signal.participant_name, signal.participant_id, 'participant');
}

function signalToMark(signal = {}, policy = {}) {
  const isLeft = signal.type === 'participant_left';
  return compactObject({
    id: firstNonEmpty(signal.source_event_id, signalId(signal)),
    source: `${signal.meeting?.platform ?? 'meeting'}_participant_track`,
    captured_at_ms: signal.occurred_at_ms,
    kind: isLeft ? policy.output_left_kind : policy.output_join_kind,
    label: `${participantLabel(signal)} ${isLeft ? 'left' : 'joined'}`,
    intent: 'participant_track',
    payload: {
      meeting: signal.meeting,
      participant_id: signal.participant_id,
      participant_name: signal.participant_name,
      source_event_id: signal.source_event_id,
      source: signal.source,
      raw: signal.raw,
    },
  });
}

export function buildMeetingPlatformParticipantTrackPlan(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const strategy = buildMeetingPlatformAdaptationStrategy(key, options);
  const integration = buildPlatformIntegrationPlan(key, options);
  return compactObject({
    type: 'meeting_platform_participant_track_plan',
    schema: MEETING_PLATFORM_PARTICIPANT_TRACK_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_PARTICIPANT_TRACK_SCHEMA_VERSION,
    platform: key,
    display_name: strategy.display_name ?? integration.display_name,
    status: 'host_roster_observer_or_provider_events_required',
    realtime_ready_when_snapshots_available: true,
    provider_events_block_realtime: false,
    transcript_blocks_realtime: false,
    track_source_order: participantSourceCatalog(key),
    filter_policy: participantFilterPolicy(options),
    output_contract: {
      track: 'participant',
      mark_intent: 'participant_track',
      default_join_kind: 'participant_joined',
      default_left_kind: 'participant_left',
      timestamp_field: 'captured_at_ms',
      transcript_required: false,
      provider_event_required: false,
      inserts_content_text: false,
    },
    required_input_fields: ['meeting_id', 'captured_at_ms_or_observed_at_ms', 'participant_id_or_participant_name'],
    next_actions: [
      'capture_roster_snapshots_or_join_leave_signals_from_host_observer',
      'apply_duplicate_and_reconnect_filters_before_insert',
      'insert_only_participant_position_markers_without_transcript_content',
    ],
  });
}

export function buildMeetingPlatformParticipantTrackMatrix(options = {}) {
  const plans = selectedPlatforms(options).map((platform) => buildMeetingPlatformParticipantTrackPlan(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_participant_track_matrix',
    schema: MEETING_PLATFORM_PARTICIPANT_TRACK_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_PARTICIPANT_TRACK_SCHEMA_VERSION,
    platform_count: plans.length,
    realtime_ready_when_snapshots_available_count: plans.filter((plan) => plan.realtime_ready_when_snapshots_available).length,
    provider_blocking_count: plans.filter((plan) => plan.provider_events_block_realtime).length,
    transcript_blocking_count: plans.filter((plan) => plan.transcript_blocks_realtime).length,
    platforms: plans.map((plan) => plan.platform),
    rows: plans.map((plan) => ({
      platform: plan.platform,
      display_name: plan.display_name,
      status: plan.status,
      source_order: plan.track_source_order.map((source) => source.id),
      duplicate_window_ms: plan.filter_policy.duplicate_window_ms,
      suppress_reconnect_gap_ms: plan.filter_policy.suppress_reconnect_gap_ms,
      leave_stable_ms: plan.filter_policy.leave_stable_ms,
      provider_events_block_realtime: plan.provider_events_block_realtime,
      transcript_blocks_realtime: plan.transcript_blocks_realtime,
      next_actions: plan.next_actions,
    })),
    plans,
    next_actions: unique(plans.flatMap((plan) => plan.next_actions ?? [])),
  };
}

export function buildMeetingPlatformParticipantTrack(platform, input = {}, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const policy = participantFilterPolicy(options);
  const snapshots = snapshotArray(input);
  const snapshotSignals = signalsFromSnapshots(snapshots, policy, options);
  const explicitSignals = normalizeSignalArray(input, options);
  const deduped = filterDuplicateSignals([...snapshotSignals.signals, ...explicitSignals], policy);
  const filtered = suppressReconnectPairs(deduped.signals, policy);
  const marks = filtered.signals.map((signal) => signalToMark(signal, policy));
  return compactObject({
    type: 'meeting_platform_participant_track',
    schema: MEETING_PLATFORM_PARTICIPANT_TRACK_SCHEMA,
    schema_version: MEETING_PLATFORM_PARTICIPANT_TRACK_SCHEMA_VERSION,
    platform: key,
    status: marks.length > 0 ? 'participant_track_ready' : 'no_participant_markers',
    input_snapshot_count: snapshots.length,
    normalized_signal_count: explicitSignals.length,
    generated_signal_count: snapshotSignals.signals.length,
    signal_count: filtered.signals.length,
    mark_count: marks.length,
    filter_policy: policy,
    signals: filtered.signals,
    marks,
    diagnostics: {
      ...snapshotSignals.diagnostics,
      dropped_duplicate_count: deduped.dropped_duplicate_count,
      suppressed_reconnect_pair_count: filtered.suppressed_reconnect_pair_count,
    },
    next_actions: marks.length > 0
      ? ['insert_participant_track_marks_on_current_axis']
      : ['collect_roster_snapshots_or_join_leave_signals'],
  });
}
