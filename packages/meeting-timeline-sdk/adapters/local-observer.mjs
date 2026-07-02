import { MeetingTimelineSdkError, compactObject, normalizeAbsoluteMs } from '../index.mjs';
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
    getState() {
      return observer.getState();
    },
    reset(nextState = null) {
      return observer.reset(nextState);
    },
  };
}
