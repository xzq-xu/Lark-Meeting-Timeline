import { MeetingTimelineSdkError, compactObject, normalizeAbsoluteMs } from './internal-utils.mjs';
import { applyMeetingSignals } from './core.mjs';
import { detectMeetingFromUrl } from './meeting-url.mjs';

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

function observedAtMs(snapshot = {}, options = {}) {
  const value = firstNonEmpty(
    snapshot.observed_at_ms,
    snapshot.observedAtMs,
    snapshot.detected_at_ms,
    snapshot.detectedAtMs,
    snapshot.timestamp_ms,
    snapshot.timestampMs,
    snapshot.timestamp,
    snapshot.time,
    snapshot.ts,
    options.observedAtMs,
    options.receivedAtMs,
    Date.now(),
  );
  return normalizeAbsoluteMs(value, 'local_observer_snapshot_time');
}

function urlSnapshot(input = {}) {
  return firstPath(input, [
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

function titleSnapshot(input = {}) {
  return firstPath(input, ['title', 'topic', 'name', 'window.title', 'tab.title']);
}

function explicitMeeting(input = {}) {
  const platform = firstPath(input, [
    'meeting.platform',
    'meeting_platform',
    'meetingPlatform',
    'detected_platform',
    'detectedPlatform',
    'platform',
    'provider',
  ]);
  const meetingId = firstPath(input, [
    'meeting.meeting_id',
    'meeting.meetingId',
    'meeting.id',
    'meeting_id',
    'meetingId',
    'id',
  ]);
  if (!platform || !meetingId) return null;
  return compactObject({
    platform,
    meeting_id: String(meetingId),
    external_meeting_id: firstPath(input, [
      'meeting.external_meeting_id',
      'meeting.externalMeetingId',
      'external_meeting_id',
      'externalMeetingId',
    ]),
    meeting_url: urlSnapshot(input),
    title: titleSnapshot(input),
    confidence: 'explicit',
  });
}

function meetingFromSnapshot(snapshot = {}) {
  return explicitMeeting(snapshot) ?? detectMeetingFromUrl(snapshot);
}

function sameMeeting(left = {}, right = {}) {
  return Boolean(left && right)
    && String(left.platform || '') === String(right.platform || '')
    && String(left.meeting_id || '') === String(right.meeting_id || '');
}

function nestedBoolean(snapshot = {}, paths = []) {
  for (const path of paths) {
    const value = getPath(snapshot, path);
    if (typeof value === 'boolean') return value;
  }
  return undefined;
}

function numericTime(value) {
  if (value == null || value === '') return null;
  try {
    return normalizeAbsoluteMs(value, 'candidate_time');
  } catch {
    const parsed = Date.parse(String(value));
    return Number.isFinite(parsed) ? parsed : null;
  }
}

function candidateList(input = {}) {
  const rows = Array.isArray(input)
    ? input
    : firstPath(input, ['snapshots', 'candidates', 'items', 'tabs', 'windows']) ?? [];
  return rows.flatMap((item) => {
    if (Array.isArray(item?.tabs)) {
      const { tabs, ...windowSnapshot } = item;
      return tabs.map((tab) => compactObject({
        ...tab,
        window: windowSnapshot,
        url: tab.url ?? windowSnapshot.url,
        title: tab.title ?? windowSnapshot.title,
      }));
    }
    return [item];
  });
}

function candidateRecencyMs(snapshot = {}) {
  return numericTime(firstPath(snapshot, [
    'lastFocusedAtMs',
    'last_focused_at_ms',
    'window.lastFocusedAtMs',
    'window.last_focused_at_ms',
    'tab.lastAccessed',
    'tab.last_accessed',
    'observedAtMs',
    'observed_at_ms',
    'timestampMs',
    'timestamp_ms',
  ]));
}

function confidenceScore(confidence) {
  if (confidence === 'explicit') return 45;
  if (confidence === 'high') return 35;
  if (confidence === 'medium') return 20;
  if (confidence === 'low') return 5;
  return 10;
}

function scoreMeetingCandidate(snapshot = {}, detected = {}, options = {}) {
  let score = confidenceScore(detected.confidence);
  if (sameMeeting(detected, options.preferredMeeting ?? options.activeMeeting)) score += 30;
  if (detected.platform === options.preferredPlatform) score += 12;
  if (String(detected.meeting_id) === String(options.preferredMeetingId ?? '')) score += 18;
  if (nestedBoolean(snapshot, ['focused', 'window.focused', 'tab.highlighted']) === true) score += 18;
  if (nestedBoolean(snapshot, ['tab.active', 'window.active', 'selected']) === true) score += 16;
  if (snapshot.active === true) score += 8;
  if (nestedBoolean(snapshot, ['audible', 'tab.audible', 'speaking']) === true) score += 6;
  if (nestedBoolean(snapshot, ['visible', 'window.visible', 'tab.visible']) !== false) score += 4;
  if (snapshot.active === false) score -= 2;
  const recency = candidateRecencyMs(snapshot);
  if (recency != null) score += Math.min(10, Math.max(0, (recency - Number(options.recencyBaseMs ?? 0)) / 60_000));
  return score;
}

function closedCandidate(snapshot = {}) {
  return snapshot.closed === true
    || snapshot.in_meeting === false
    || snapshot.inMeeting === false
    || snapshot.visible === false
    || snapshot.window?.visible === false
    || snapshot.tab?.visible === false
    || snapshot.discarded === true
    || snapshot.tab?.discarded === true;
}

function safeCandidateSnapshot(snapshot = {}, selected = {}) {
  const safe = {
    ...snapshot,
    active: snapshot.in_meeting ?? snapshot.inMeeting ?? undefined,
    selected_from_candidates: true,
    selection: {
      score: selected.score,
      rank: selected.rank,
      candidate_count: selected.candidate_count,
    },
    candidate_raw: snapshot,
  };
  return compactObject(safe);
}

export function selectMeetingSnapshot(input = {}, options = {}) {
  const rows = candidateList(input);
  const candidates = rows
    .map((snapshot, index) => {
      const detectedMeeting = meetingFromSnapshot(snapshot);
      if (!detectedMeeting || closedCandidate(snapshot)) return null;
      return {
        index,
        snapshot,
        detectedMeeting,
        score: scoreMeetingCandidate(snapshot, detectedMeeting, options),
      };
    })
    .filter(Boolean)
    .sort((left, right) => right.score - left.score || left.index - right.index)
    .map((candidate, rank, all) => ({
      ...candidate,
      rank,
      candidate_count: all.length,
    }));
  const selected = candidates[0] ?? null;
  return {
    selectedSnapshot: selected ? safeCandidateSnapshot(selected.snapshot, selected) : null,
    detectedMeeting: selected?.detectedMeeting ?? null,
    candidates,
  };
}

function observerSignal(type, meeting, atMs, snapshot = {}, options = {}) {
  return compactObject({
    type,
    meeting: {
      platform: meeting.platform,
      meeting_id: meeting.meeting_id,
      external_meeting_id: meeting.external_meeting_id,
      meeting_url: meeting.meeting_url,
      title: meeting.title,
    },
    occurred_at_ms: atMs,
    source_event_id: firstNonEmpty(snapshot.source_event_id, snapshot.sourceEventId, snapshot.event_id, snapshot.eventId, snapshot.id),
    source: options.source ?? 'local_observer',
    raw: snapshot,
  });
}

function normalizeObserverState(state = null) {
  if (!state?.activeMeeting) return { activeMeeting: null, activeSinceMs: null, lastObservedAtMs: null };
  return {
    activeMeeting: state.activeMeeting,
    activeSinceMs: state.activeSinceMs ?? null,
    lastObservedAtMs: state.lastObservedAtMs ?? null,
  };
}

function splitTimelineObserverOptions(options = {}) {
  const { applyOptions, ...observerOptions } = options ?? {};
  return {
    observerOptions,
    applyOptions: applyOptions ?? {},
  };
}

export function observeMeetingSnapshot(state = null, snapshot = {}, options = {}) {
  const previous = normalizeObserverState(state);
  const atMs = observedAtMs(snapshot, options);
  const detected = meetingFromSnapshot(snapshot);
  const endedBySnapshot = snapshot.active === false
    || snapshot.in_meeting === false
    || snapshot.inMeeting === false
    || snapshot.closed === true
    || snapshot.visible === false;
  const signals = [];

  if (!detected || endedBySnapshot) {
    if (previous.activeMeeting) {
      signals.push(observerSignal('meeting_ended', previous.activeMeeting, atMs, snapshot, options));
    }
    return {
      state: {
        activeMeeting: null,
        activeSinceMs: null,
        lastObservedAtMs: atMs,
      },
      signals,
      detectedMeeting: detected,
    };
  }

  if (!previous.activeMeeting) {
    signals.push(observerSignal('meeting_started', detected, atMs, snapshot, options));
    return {
      state: {
        activeMeeting: detected,
        activeSinceMs: atMs,
        lastObservedAtMs: atMs,
      },
      signals,
      detectedMeeting: detected,
    };
  }

  if (!sameMeeting(previous.activeMeeting, detected)) {
    signals.push(observerSignal('meeting_ended', previous.activeMeeting, atMs, snapshot, options));
    signals.push(observerSignal('meeting_started', detected, atMs, snapshot, options));
    return {
      state: {
        activeMeeting: detected,
        activeSinceMs: atMs,
        lastObservedAtMs: atMs,
      },
      signals,
      detectedMeeting: detected,
    };
  }

  return {
    state: {
      activeMeeting: {
        ...previous.activeMeeting,
        ...detected,
      },
      activeSinceMs: previous.activeSinceMs ?? atMs,
      lastObservedAtMs: atMs,
    },
    signals,
    detectedMeeting: detected,
  };
}

export function createLocalMeetingObserver(options = {}) {
  let state = normalizeObserverState(options.initialState);
  return {
    observe(snapshot = {}, observeOptions = {}) {
      const result = observeMeetingSnapshot(state, snapshot, {
        ...options,
        ...observeOptions,
      });
      state = result.state;
      return result;
    },
    observeCandidates(candidates = [], observeOptions = {}) {
      const selected = selectMeetingSnapshot(candidates, {
        ...options,
        ...observeOptions,
        activeMeeting: state.activeMeeting,
      });
      const snapshot = selected.selectedSnapshot ?? {
        active: false,
        observedAtMs: observeOptions.observedAtMs ?? observeOptions.receivedAtMs,
        candidate_count: selected.candidates.length,
      };
      const result = observeMeetingSnapshot(state, snapshot, {
        ...options,
        ...observeOptions,
      });
      state = result.state;
      return {
        ...result,
        selection: selected,
      };
    },
    getState() {
      return state;
    },
    reset(nextState = null) {
      state = normalizeObserverState(nextState);
      return state;
    },
  };
}

export function createLocalMeetingTimelineObserver(client, options = {}) {
  if (!client || typeof client !== 'object') {
    throw new MeetingTimelineSdkError('Meeting timeline client is required for createLocalMeetingTimelineObserver');
  }

  const base = splitTimelineObserverOptions(options);
  const observer = createLocalMeetingObserver(base.observerOptions);
  return {
    async observe(snapshot = {}, observeOptions = {}) {
      const next = splitTimelineObserverOptions(observeOptions);
      const observed = observer.observe(snapshot, next.observerOptions);
      const results = observed.signals.length === 0
        ? []
        : await applyMeetingSignals(client, observed.signals, {
          ...base.applyOptions,
          ...next.applyOptions,
        });
      return {
        ...observed,
        results,
      };
    },
    async observeCandidates(candidates = [], observeOptions = {}) {
      const next = splitTimelineObserverOptions(observeOptions);
      const observed = observer.observeCandidates(candidates, next.observerOptions);
      const results = observed.signals.length === 0
        ? []
        : await applyMeetingSignals(client, observed.signals, {
          ...base.applyOptions,
          ...next.applyOptions,
        });
      return {
        ...observed,
        results,
      };
    },
    getState() {
      return observer.getState();
    },
    reset(nextState = null) {
      return observer.reset(nextState);
    },
  };
}
