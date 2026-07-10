import { compactObject, normalizeAbsoluteMs } from './internal-utils.mjs';
import { applyMeetingSignals } from './core.mjs';
import { detectMeetingFromUrl, stableMeetingIdFromUrl } from './meeting-url.mjs';
import { normalizeMeetingPlatform } from './platform-setup.mjs';

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function getPath(raw, path) {
  const parts = path.split('.');
  let node = raw;
  for (const part of parts) node = node?.[part];
  return node;
}

function firstPath(raw, paths) {
  return firstNonEmpty(...paths.map((path) => getPath(raw, path)));
}

function firstBoolean(raw, paths) {
  for (const path of paths) {
    const value = getPath(raw, path);
    if (typeof value === 'boolean') return value;
  }
  return undefined;
}

function normalizeTimeMs(raw = {}, options = {}) {
  const value = firstNonEmpty(
    raw.observed_at_ms,
    raw.observedAtMs,
    raw.detected_at_ms,
    raw.detectedAtMs,
    raw.occurred_at_ms,
    raw.occurredAtMs,
    raw.timestamp_ms,
    raw.timestampMs,
    raw.timestamp,
    raw.time,
    raw.ts,
    options.observedAtMs,
    options.receivedAtMs,
    Date.now(),
  );
  return normalizeAbsoluteMs(value, 'active_speaker_observed_at_ms');
}

function explicitPlatform(raw = {}) {
  const value = firstPath(raw, [
    'meeting.platform',
    'meeting_platform',
    'meetingPlatform',
    'detected_platform',
    'detectedPlatform',
    'platform',
    'provider',
    'app_platform',
    'appPlatform',
  ]);
  if (!value || typeof value === 'object') return null;
  try {
    return normalizeMeetingPlatform(String(value));
  } catch {
    return String(value);
  }
}

function meetingUrl(raw = {}) {
  return firstPath(raw, [
    'meeting.meeting_url',
    'meeting.meetingUrl',
    'meeting.url',
    'meeting.join_url',
    'meeting.joinUrl',
    'meeting_url',
    'meetingUrl',
    'join_url',
    'joinUrl',
    'url',
    'window.url',
    'browser.url',
    'tab.url',
  ]);
}

function meetingTitle(raw = {}) {
  return firstPath(raw, [
    'meeting.title',
    'meeting.topic',
    'meeting.name',
    'title',
    'topic',
    'name',
    'window.title',
    'tab.title',
  ]);
}

function meetingIdentity(raw = {}) {
  const detected = detectMeetingFromUrl(raw) ?? detectMeetingFromUrl(meetingUrl(raw));
  const platform = explicitPlatform(raw) ?? detected?.platform ?? 'local_detector';
  const id = firstPath(raw, [
    'meeting.meeting_id',
    'meeting.meetingId',
    'meeting.id',
    'meeting.external_meeting_id',
    'meeting.externalMeetingId',
    'meeting_id',
    'meetingId',
    'external_meeting_id',
    'externalMeetingId',
    'session_id',
    'sessionId',
    'id',
  ]) ?? detected?.meeting_id ?? stableMeetingIdFromUrl(meetingUrl(raw));
  if (!id) return null;
  return compactObject({
    platform,
    meeting_id: String(id),
    external_meeting_id: firstPath(raw, [
      'meeting.external_meeting_id',
      'meeting.externalMeetingId',
      'external_meeting_id',
      'externalMeetingId',
      'meeting_no',
      'meetingNo',
    ]) ?? detected?.external_meeting_id,
    meeting_url: meetingUrl(raw) ?? detected?.meeting_url,
    title: meetingTitle(raw) ?? detected?.title,
  });
}

function speakerId(raw = {}) {
  return firstPath(raw, [
    'active_speaker.id',
    'active_speaker.speaker_id',
    'active_speaker.speakerId',
    'active_speaker.participant_id',
    'active_speaker.participantId',
    'activeSpeaker.id',
    'activeSpeaker.speaker_id',
    'activeSpeaker.speakerId',
    'activeSpeaker.participant_id',
    'activeSpeaker.participantId',
    'speaker.id',
    'speaker.speaker_id',
    'speaker.speakerId',
    'speaker.participant_id',
    'speaker.participantId',
    'participant.id',
    'participant.participant_id',
    'participant.participantId',
    'user.id',
    'speaker_id',
    'speakerId',
    'participant_id',
    'participantId',
    'user_id',
    'userId',
  ]);
}

function speakerName(raw = {}) {
  return firstPath(raw, [
    'active_speaker.name',
    'active_speaker.display_name',
    'active_speaker.displayName',
    'active_speaker.user_name',
    'active_speaker.userName',
    'activeSpeaker.name',
    'activeSpeaker.display_name',
    'activeSpeaker.displayName',
    'activeSpeaker.user_name',
    'activeSpeaker.userName',
    'speaker.name',
    'speaker.display_name',
    'speaker.displayName',
    'speaker.user_name',
    'speaker.userName',
    'participant.name',
    'participant.display_name',
    'participant.displayName',
    'user.name',
    'speaker_name',
    'speakerName',
    'participant_name',
    'participantName',
    'user_name',
    'userName',
    'name',
  ]);
}

function speakingFlag(raw = {}, hasSpeaker = false) {
  const value = firstBoolean(raw, [
    'active_speaker.speaking',
    'active_speaker.active',
    'active_speaker.is_speaking',
    'active_speaker.isSpeaking',
    'activeSpeaker.speaking',
    'activeSpeaker.active',
    'activeSpeaker.is_speaking',
    'activeSpeaker.isSpeaking',
    'speaker.speaking',
    'speaker.active',
    'speaker.is_speaking',
    'speaker.isSpeaking',
    'participant.speaking',
    'participant.active_speaker',
    'participant.activeSpeaker',
    'tile.speaking',
    'tile.active',
    'speaking',
    'is_speaking',
    'isSpeaking',
    'active',
  ]);
  if (value != null) return value;
  return hasSpeaker || undefined;
}

function speakerKey(sample = {}) {
  return sample.speaker_id ?? sample.speaker_name ?? null;
}

function sameSpeaker(left = {}, right = {}) {
  const leftKey = speakerKey(left);
  const rightKey = speakerKey(right);
  return Boolean(leftKey && rightKey && String(leftKey) === String(rightKey));
}

function normalizeActiveSpeakerState(state = null) {
  return {
    activeSpeaker: state?.activeSpeaker ?? null,
    activeSinceMs: state?.activeSinceMs ?? null,
    candidateSpeaker: state?.candidateSpeaker ?? null,
    candidateSinceMs: state?.candidateSinceMs ?? null,
    silenceSinceMs: state?.silenceSinceMs ?? null,
    lastObservedAtMs: state?.lastObservedAtMs ?? null,
  };
}

function stableWindowMs(options = {}, switching = false) {
  if (switching) {
    return Number(firstNonEmpty(
      options.switchStableMs,
      options.switch_stable_ms,
      options.minStableMs,
      options.min_stable_ms,
      400,
    ));
  }
  return Number(firstNonEmpty(options.minStableMs, options.min_stable_ms, 300));
}

function endIdleMs(options = {}) {
  return Number(firstNonEmpty(options.endIdleMs, options.end_idle_ms, options.silenceEndMs, options.silence_end_ms, 1_500));
}

function emitSpeakerEnd(options = {}) {
  return options.emitSpeakerEnd !== false && options.emit_speaker_end !== false;
}

function sourceName(options = {}) {
  return String(options.source ?? 'active_speaker_detector');
}

function signal(type, sample, occurredAtMs, options = {}) {
  return compactObject({
    type,
    meeting: sample.meeting,
    occurred_at_ms: occurredAtMs,
    source_event_id: type === 'speaker_started'
      ? firstNonEmpty(sample.source_event_id, sample.sourceEventId, `${sample.meeting.meeting_id}:${speakerKey(sample)}:${occurredAtMs}:start`)
      : firstNonEmpty(sample.source_event_id, sample.sourceEventId, `${sample.meeting.meeting_id}:${speakerKey(sample)}:${occurredAtMs}:end`),
    source: sourceName(options),
    speaker_id: sample.speaker_id,
    speaker_name: sample.speaker_name,
    participant_id: sample.participant_id,
    participant_name: sample.participant_name,
    raw: sample.raw,
  });
}

export function normalizeActiveSpeakerSample(raw = {}, options = {}) {
  const meeting = meetingIdentity(raw);
  if (!meeting) return null;
  const id = speakerId(raw);
  const name = speakerName(raw);
  const hasSpeaker = id != null || name != null;
  const speaking = speakingFlag(raw, hasSpeaker);
  return compactObject({
    meeting,
    observed_at_ms: normalizeTimeMs(raw, options),
    speaking: speaking === undefined ? false : speaking,
    speaker_id: id != null ? String(id) : undefined,
    speaker_name: name != null ? String(name) : undefined,
    participant_id: firstPath(raw, ['participant.id', 'participant.participant_id', 'participant.participantId', 'participant_id', 'participantId']),
    participant_name: firstPath(raw, ['participant.name', 'participant.display_name', 'participant.displayName', 'participant_name', 'participantName']),
    source_event_id: firstNonEmpty(raw.source_event_id, raw.sourceEventId, raw.event_id, raw.eventId, raw.id),
    raw,
  });
}

export function observeActiveSpeakerSample(state = null, raw = {}, options = {}) {
  const previous = normalizeActiveSpeakerState(state);
  const sample = normalizeActiveSpeakerSample(raw, options);
  const atMs = sample?.observed_at_ms ?? normalizeTimeMs(raw, options);
  const signals = [];

  if (!sample || sample.speaking === false || !speakerKey(sample)) {
    if (!previous.activeSpeaker) {
      return {
        state: {
          ...previous,
          candidateSpeaker: null,
          candidateSinceMs: null,
          silenceSinceMs: null,
          lastObservedAtMs: atMs,
        },
        signals,
        sample,
      };
    }
    const silenceSinceMs = previous.silenceSinceMs ?? atMs;
    if (atMs - silenceSinceMs >= endIdleMs(options)) {
      if (emitSpeakerEnd(options)) {
        signals.push(signal('speaker_ended', previous.activeSpeaker, silenceSinceMs, options));
      }
      return {
        state: {
          activeSpeaker: null,
          activeSinceMs: null,
          candidateSpeaker: null,
          candidateSinceMs: null,
          silenceSinceMs: null,
          lastObservedAtMs: atMs,
        },
        signals,
        sample,
      };
    }
    return {
      state: {
        ...previous,
        candidateSpeaker: null,
        candidateSinceMs: null,
        silenceSinceMs,
        lastObservedAtMs: atMs,
      },
      signals,
      sample,
    };
  }

  if (previous.activeSpeaker && sameSpeaker(previous.activeSpeaker, sample)) {
    return {
      state: {
        ...previous,
        activeSpeaker: { ...previous.activeSpeaker, ...sample },
        candidateSpeaker: null,
        candidateSinceMs: null,
        silenceSinceMs: null,
        lastObservedAtMs: atMs,
      },
      signals,
      sample,
    };
  }

  const switching = Boolean(previous.activeSpeaker);
  const existingCandidate = previous.candidateSpeaker && sameSpeaker(previous.candidateSpeaker, sample);
  const candidateSinceMs = existingCandidate ? previous.candidateSinceMs : atMs;
  const candidateSpeaker = existingCandidate
    ? { ...previous.candidateSpeaker, ...sample }
    : sample;
  const windowMs = stableWindowMs(options, switching);

  if (atMs - candidateSinceMs < windowMs) {
    return {
      state: {
        ...previous,
        candidateSpeaker,
        candidateSinceMs,
        silenceSinceMs: null,
        lastObservedAtMs: atMs,
      },
      signals,
      sample,
    };
  }

  if (previous.activeSpeaker && emitSpeakerEnd(options)) {
    signals.push(signal('speaker_ended', previous.activeSpeaker, candidateSinceMs, options));
  }
  signals.push(signal('speaker_started', candidateSpeaker, candidateSinceMs, options));
  return {
    state: {
      activeSpeaker: candidateSpeaker,
      activeSinceMs: candidateSinceMs,
      candidateSpeaker: null,
      candidateSinceMs: null,
      silenceSinceMs: null,
      lastObservedAtMs: atMs,
    },
    signals,
    sample,
  };
}

export function observeActiveSpeakerSamples(state = null, rows = [], options = {}) {
  let nextState = normalizeActiveSpeakerState(state);
  const observations = [];
  const signals = [];
  for (const row of rows) {
    const result = observeActiveSpeakerSample(nextState, row, options);
    nextState = result.state;
    observations.push(result);
    signals.push(...result.signals);
  }
  return { state: nextState, signals, observations };
}

export function createActiveSpeakerObserver(options = {}) {
  let state = normalizeActiveSpeakerState(options.initialState);
  return {
    observe(sample = {}, observeOptions = {}) {
      const result = observeActiveSpeakerSample(state, sample, {
        ...options,
        ...observeOptions,
      });
      state = result.state;
      return result;
    },
    observeMany(samples = [], observeOptions = {}) {
      const result = observeActiveSpeakerSamples(state, samples, {
        ...options,
        ...observeOptions,
      });
      state = result.state;
      return result;
    },
    getState() {
      return state;
    },
    reset(nextState = null) {
      state = normalizeActiveSpeakerState(nextState);
      return state;
    },
  };
}

export function createActiveSpeakerTimelineObserver(client, options = {}) {
  const observer = createActiveSpeakerObserver(options);
  const baseApplyOptions = options.applyOptions ?? {};
  return {
    async observe(sample = {}, observeOptions = {}) {
      const { applyOptions, ...nextOptions } = observeOptions;
      const observed = observer.observe(sample, nextOptions);
      const results = observed.signals.length === 0
        ? []
        : await applyMeetingSignals(client, observed.signals, {
          ...baseApplyOptions,
          ...applyOptions,
        });
      return { ...observed, results };
    },
    async observeMany(samples = [], observeOptions = {}) {
      const { applyOptions, ...nextOptions } = observeOptions;
      const observed = observer.observeMany(samples, nextOptions);
      const results = observed.signals.length === 0
        ? []
        : await applyMeetingSignals(client, observed.signals, {
          ...baseApplyOptions,
          ...applyOptions,
        });
      return { ...observed, results };
    },
    getState() {
      return observer.getState();
    },
    reset(nextState = null) {
      return observer.reset(nextState);
    },
  };
}
