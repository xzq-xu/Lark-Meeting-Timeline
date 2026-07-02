import { compactObject } from '../index.mjs';
import { normalizeMeetingSignal } from './core.mjs';

const LOCAL_SOURCE_RE = /(^|_)(local|observer|desktop|browser|manual|device)(_|$)/i;

function asArray(value) {
  return Array.isArray(value) ? value : value == null ? [] : [value];
}

function normalizeUrl(value) {
  if (!value) return null;
  try {
    const parsed = new URL(String(value));
    parsed.hash = '';
    parsed.search = '';
    return `${parsed.hostname.toLowerCase()}${parsed.pathname.replace(/\/+$/, '')}`;
  } catch {
    return String(value).trim().toLowerCase().replace(/[?#].*$/, '').replace(/\/+$/, '') || null;
  }
}

function signalTime(signal) {
  return Number(signal.occurred_at_ms ?? 0) || 0;
}

function sourceOf(signal = {}) {
  return String(signal.source || 'adapter');
}

function isLocalSource(source) {
  return LOCAL_SOURCE_RE.test(String(source || ''));
}

function isProviderSource(source) {
  return !isLocalSource(source);
}

function meetingRef(signal = {}) {
  const meeting = signal.meeting ?? {};
  return compactObject({
    platform: meeting.platform ?? signal.platform,
    meeting_id: meeting.meeting_id,
    external_meeting_id: meeting.external_meeting_id,
    meeting_url: normalizeUrl(meeting.meeting_url),
  });
}

function meetingIds(ref = {}) {
  return [ref.meeting_id, ref.external_meeting_id]
    .filter((value) => value != null && value !== '')
    .map((value) => String(value));
}

function sameMeetingRef(left = {}, right = {}) {
  if (!left.platform || !right.platform || String(left.platform) !== String(right.platform)) return false;
  const leftIds = meetingIds(left);
  const rightIds = meetingIds(right);
  if (leftIds.some((leftId) => rightIds.includes(leftId))) return true;
  return Boolean(left.meeting_url && right.meeting_url && left.meeting_url === right.meeting_url);
}

function subjectKey(signal = {}) {
  if (signal.type === 'participant_joined' || signal.type === 'participant_left') {
    return signal.participant_id ?? signal.participant_name ?? 'participant';
  }
  if (signal.type === 'speaker_started' || signal.type === 'speaker_ended') {
    return signal.speaker_id ?? signal.speaker_name ?? signal.participant_id ?? signal.participant_name ?? 'speaker';
  }
  if (signal.type === 'artifact_ready') {
    return `${signal.artifact_kind ?? 'artifact'}:${signal.artifact_id ?? signal.artifact_url ?? ''}`;
  }
  if (signal.type === 'subscription_lifecycle') {
    return `${signal.subscription_id ?? signal.subscription_name ?? ''}:${signal.lifecycle_type ?? ''}`;
  }
  return '';
}

export function meetingSignalFingerprint(signalInput = {}) {
  const signal = normalizeMeetingSignal(signalInput);
  const source = sourceOf(signal);
  if (signal.source_event_id) return `event:${source}:${signal.source_event_id}`;
  const ref = meetingRef(signal);
  return [
    signal.type,
    ref.platform ?? signal.platform ?? 'unknown',
    ref.meeting_id ?? ref.external_meeting_id ?? ref.meeting_url ?? 'unknown',
    signalTime(signal),
    subjectKey(signal),
  ].join('|');
}

function normalizeState(state = {}) {
  return {
    seen: Array.isArray(state.seen) ? [...state.seen] : [],
    active_meetings: Array.isArray(state.active_meetings) ? [...state.active_meetings] : [],
    ended_meetings: Array.isArray(state.ended_meetings) ? [...state.ended_meetings] : [],
    recent_speakers: Array.isArray(state.recent_speakers) ? [...state.recent_speakers] : [],
  };
}

function compactState(state, options = {}) {
  const maxSeen = Number(options.maxSeen ?? options.max_seen ?? 1_000);
  const maxMeetings = Number(options.maxMeetings ?? options.max_meetings ?? 100);
  const maxSpeakers = Number(options.maxSpeakers ?? options.max_speakers ?? 200);
  return {
    seen: state.seen.slice(-maxSeen),
    active_meetings: state.active_meetings.slice(-maxMeetings),
    ended_meetings: state.ended_meetings.slice(-maxMeetings),
    recent_speakers: state.recent_speakers.slice(-maxSpeakers),
  };
}

function rememberSeen(state, signal, fingerprint) {
  state.seen.push({
    key: fingerprint,
    type: signal.type,
    source: sourceOf(signal),
    occurred_at_ms: signalTime(signal),
  });
}

function seenEntry(state, fingerprint) {
  return state.seen.find((item) => item.key === fingerprint);
}

function meetingStateEntry(signal) {
  const ref = meetingRef(signal);
  return {
    ...ref,
    source: sourceOf(signal),
    occurred_at_ms: signalTime(signal),
    source_event_id: signal.source_event_id,
  };
}

function findMeeting(rows = [], signal = {}) {
  const ref = meetingRef(signal);
  return rows.find((item) => sameMeetingRef(item, ref));
}

function removeMeeting(rows = [], signal = {}) {
  const ref = meetingRef(signal);
  return rows.filter((item) => !sameMeetingRef(item, ref));
}

function near(leftMs, rightMs, windowMs) {
  return Math.abs(Number(leftMs ?? 0) - Number(rightMs ?? 0)) <= windowMs;
}

function skipped(signal, reason, details = {}) {
  return compactObject({
    action: 'skip',
    reason,
    signal,
    ...details,
  });
}

function applied(signal, reason, details = {}) {
  return compactObject({
    action: 'apply',
    reason,
    signal,
    ...details,
  });
}

function reconcileMeetingStarted(state, signal, options) {
  const active = findMeeting(state.active_meetings, signal);
  const source = sourceOf(signal);
  const duplicateWindowMs = Number(options.duplicateWindowMs ?? options.duplicate_window_ms ?? 30_000);
  if (!active) {
    state.active_meetings.push(meetingStateEntry(signal));
    state.ended_meetings = removeMeeting(state.ended_meetings, signal);
    return applied(signal);
  }
  if (isLocalSource(source) && isProviderSource(active.source)) {
    return skipped(signal, 'provider_axis_already_active', { existing: active });
  }
  if (isLocalSource(source) && isLocalSource(active.source) && near(active.occurred_at_ms, signalTime(signal), duplicateWindowMs)) {
    return skipped(signal, 'duplicate_local_meeting_start', { existing: active });
  }
  if (isProviderSource(source) && isProviderSource(active.source) && near(active.occurred_at_ms, signalTime(signal), duplicateWindowMs)) {
    return skipped(signal, 'duplicate_provider_meeting_start', { existing: active });
  }
  state.active_meetings = removeMeeting(state.active_meetings, signal);
  state.active_meetings.push(meetingStateEntry(signal));
  state.ended_meetings = removeMeeting(state.ended_meetings, signal);
  return applied(signal, isProviderSource(source) && isLocalSource(active.source) ? 'provider_reconciles_local_axis' : undefined, {
    existing: active,
  });
}

function reconcileMeetingEnded(state, signal, options) {
  const duplicateWindowMs = Number(options.duplicateWindowMs ?? options.duplicate_window_ms ?? 30_000);
  const ended = findMeeting(state.ended_meetings, signal);
  if (ended && near(ended.occurred_at_ms, signalTime(signal), duplicateWindowMs)) {
    return skipped(signal, 'duplicate_meeting_end', { existing: ended });
  }
  state.active_meetings = removeMeeting(state.active_meetings, signal);
  state.ended_meetings = removeMeeting(state.ended_meetings, signal);
  state.ended_meetings.push(meetingStateEntry(signal));
  return applied(signal);
}

function speakerStateKey(signal) {
  const ref = meetingRef(signal);
  return [
    ref.platform,
    ref.meeting_id ?? ref.external_meeting_id ?? ref.meeting_url,
    signal.speaker_id ?? signal.speaker_name ?? signal.participant_id ?? signal.participant_name ?? 'speaker',
  ].join('|');
}

function reconcileSpeaker(state, signal, options) {
  const speakerWindowMs = Number(options.speakerRepeatWindowMs ?? options.speaker_repeat_window_ms ?? 5_000);
  const key = speakerStateKey(signal);
  const existing = state.recent_speakers.find((item) => item.key === key && item.type === signal.type);
  if (existing && near(existing.occurred_at_ms, signalTime(signal), speakerWindowMs)) {
    return skipped(signal, 'duplicate_speaker_signal', { existing });
  }
  state.recent_speakers = state.recent_speakers.filter((item) => item.key !== key || item.type !== signal.type);
  state.recent_speakers.push({
    key,
    type: signal.type,
    occurred_at_ms: signalTime(signal),
    source: sourceOf(signal),
  });
  return applied(signal);
}

function reconcileOne(state, signalInput, options = {}) {
  const signal = normalizeMeetingSignal(signalInput);
  const fingerprint = meetingSignalFingerprint(signal);
  const seen = seenEntry(state, fingerprint);
  if (seen) return skipped(signal, 'duplicate_signal_fingerprint', { existing: seen, fingerprint });

  let decision;
  if (signal.type === 'meeting_started') decision = reconcileMeetingStarted(state, signal, options);
  else if (signal.type === 'meeting_ended') decision = reconcileMeetingEnded(state, signal, options);
  else if (signal.type === 'speaker_started' || signal.type === 'speaker_ended') decision = reconcileSpeaker(state, signal, options);
  else decision = applied(signal);

  if (decision.action === 'apply') rememberSeen(state, signal, fingerprint);
  return {
    ...decision,
    fingerprint,
  };
}

export function reconcileMeetingSignals(signals = [], state = {}, options = {}) {
  const nextState = normalizeState(state);
  const decisions = asArray(signals).map((signal) => reconcileOne(nextState, signal, options));
  return {
    signals: decisions.filter((item) => item.action === 'apply').map((item) => item.signal),
    skipped: decisions.filter((item) => item.action === 'skip'),
    decisions,
    state: compactState(nextState, options),
  };
}

export function createMeetingSignalReconciler(options = {}) {
  let state = normalizeState(options.initialState ?? options.initial_state);
  return {
    reconcile(signals = [], reconcileOptions = {}) {
      const result = reconcileMeetingSignals(signals, state, {
        ...options,
        ...reconcileOptions,
      });
      state = normalizeState(result.state);
      return result;
    },
    getState() {
      return compactState(state, options);
    },
    reset(nextState = {}) {
      state = normalizeState(nextState);
      return compactState(state, options);
    },
  };
}
