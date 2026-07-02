import { compactObject, normalizeAbsoluteMs } from '../index.mjs';
import { createActiveSpeakerObserver, createActiveSpeakerTimelineObserver, normalizeActiveSpeakerSample } from './active-speaker.mjs';
import { selectMeetingSnapshot } from './local-observer.mjs';
import {
  createMeetingSessionDiscoveryObserver,
  createMeetingSessionTimelineDiscovery,
  normalizeMeetingSessionCandidate,
} from './meeting-session-discovery.mjs';

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

function observedAtMs(input = {}, options = {}) {
  const value = firstNonEmpty(
    input.observed_at_ms,
    input.observedAtMs,
    input.detected_at_ms,
    input.detectedAtMs,
    input.timestamp_ms,
    input.timestampMs,
    input.timestamp,
    input.time,
    input.ts,
    options.observedAtMs,
    options.receivedAtMs,
  );
  return value == null ? undefined : normalizeAbsoluteMs(value, 'native_meeting_observed_at_ms');
}

function nativeRows(input = {}, inherited = {}) {
  if (Array.isArray(input)) return input.flatMap((item) => nativeRows(item, inherited));
  if (!input || typeof input !== 'object') return [];
  const rows = [];
  const base = compactObject({
    ...inherited,
    application: input.application ?? input.app ?? inherited.application,
    app: input.app ?? input.application ?? inherited.app,
    process: input.process ?? inherited.process,
    accessibility: input.accessibility ?? input.ax ?? inherited.accessibility,
    audio: input.audio ?? inherited.audio,
  });

  const collections = [
    ['applications', input.applications],
    ['apps', input.apps],
    ['processes', input.processes],
    ['windows', input.windows],
    ['candidates', input.candidates],
    ['snapshots', input.snapshots],
    ['items', input.items],
  ].filter(([, value]) => Array.isArray(value));

  for (const [kind, collection] of collections) {
    for (const item of collection) {
      if (kind === 'applications' || kind === 'apps') {
        rows.push(...nativeRows(item, {
          ...base,
          application: item,
          app: item,
          process: item.process ?? base.process,
        }));
      } else if (kind === 'processes') {
        rows.push(...nativeRows(item, {
          ...base,
          process: item,
          application: item.application ?? base.application,
          app: item.app ?? base.app,
        }));
      } else if (kind === 'windows') {
        rows.push(...nativeRows(item, {
          ...base,
          window: item,
          accessibility: item.accessibility ?? item.ax ?? base.accessibility,
          audio: item.audio ?? base.audio,
        }));
      } else {
        rows.push(...nativeRows(item, base));
      }
    }
  }

  if (collections.length === 0) rows.push(compactObject({ ...base, ...input }));
  return rows;
}

function urlInput(input = {}) {
  return firstPath(input, [
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
    'deep_link',
    'deepLink',
    'window.url',
    'application.url',
    'app.url',
    'process.url',
  ]);
}

function titleInput(input = {}) {
  return firstPath(input, [
    'meeting.title',
    'meeting.topic',
    'meeting.name',
    'title',
    'topic',
    'name',
    'window.title',
    'window.name',
    'accessibility.title',
    'accessibility.window_title',
    'accessibility.windowTitle',
    'application.title',
    'app.title',
    'process.title',
  ]);
}

function appName(input = {}) {
  return firstPath(input, [
    'application.name',
    'app.name',
    'process.name',
    'processName',
    'process_name',
    'executable',
    'window.app_name',
    'window.appName',
    'window.ownerName',
  ]);
}

function bundleId(input = {}) {
  return firstPath(input, [
    'bundle_id',
    'bundleId',
    'application.bundle_id',
    'application.bundleId',
    'app.bundle_id',
    'app.bundleId',
    'process.bundle_id',
    'process.bundleId',
    'window.bundle_id',
    'window.bundleId',
  ]);
}

function explicitPlatform(input = {}) {
  return firstPath(input, [
    'meeting.platform',
    'meeting_platform',
    'meetingPlatform',
    'detected_platform',
    'detectedPlatform',
    'platform',
    'provider',
    'application.platform',
    'app.platform',
  ]);
}

function explicitMeetingId(input = {}) {
  return firstPath(input, [
    'meeting.meeting_id',
    'meeting.meetingId',
    'meeting.id',
    'meeting_id',
    'meetingId',
    'external_meeting_id',
    'externalMeetingId',
    'session_id',
    'sessionId',
    'window.meeting_id',
    'window.meetingId',
    'accessibility.meeting_id',
    'accessibility.meetingId',
  ]);
}

function inMeetingFlag(input = {}) {
  return firstBoolean(input, [
    'in_meeting',
    'inMeeting',
    'meeting.active',
    'meeting.in_meeting',
    'meeting.inMeeting',
    'window.in_meeting',
    'window.inMeeting',
    'accessibility.in_meeting',
    'accessibility.inMeeting',
    'accessibility.call_active',
    'accessibility.callActive',
    'audio.call_active',
    'audio.callActive',
    'call.active',
    'callActive',
    'call_active',
  ]);
}

function visibleFlag(input = {}) {
  return firstBoolean(input, ['visible', 'window.visible', 'application.visible', 'app.visible', 'accessibility.visible']);
}

function activeFlag(input = {}) {
  return firstBoolean(input, ['active', 'focused', 'selected', 'window.active', 'window.focused', 'accessibility.focused']);
}

export function normalizeNativeMeetingCandidate(input = {}, options = {}) {
  return normalizeMeetingSessionCandidate(compactObject({
    ...input,
    url: urlInput(input),
    title: titleInput(input),
    platform: explicitPlatform(input),
    meeting_id: explicitMeetingId(input),
    inMeeting: inMeetingFlag(input),
    active: activeFlag(input),
    visible: visibleFlag(input),
    observedAtMs: observedAtMs(input, options),
    application: input.application,
    app: input.app,
    process: input.process,
    window: input.window,
    accessibility: input.accessibility,
    audio: input.audio,
    processName: firstPath(input, ['processName', 'process_name']) ?? appName(input),
    bundleId: bundleId(input),
  }), options);
}

export function normalizeNativeMeetingCandidates(input = {}, options = {}) {
  return nativeRows(input).map((row) => normalizeNativeMeetingCandidate(row, options)).filter(Boolean);
}

export function selectNativeMeetingCandidate(input = {}, options = {}) {
  const normalizedCandidates = normalizeNativeMeetingCandidates(input, options);
  return {
    ...selectMeetingSnapshot(normalizedCandidates, options),
    normalizedCandidates,
  };
}

function speakerObject(input = {}) {
  const direct = firstPath(input, [
    'activeSpeaker',
    'active_speaker',
    'speaker',
    'accessibility.activeSpeaker',
    'accessibility.active_speaker',
    'accessibility.speaker',
    'audio.activeSpeaker',
    'audio.active_speaker',
    'audio.speaker',
    'window.activeSpeaker',
    'window.active_speaker',
    'meeting.activeSpeaker',
  ]);
  if (direct) return direct;
  const participants = firstPath(input, [
    'participants',
    'accessibility.participants',
    'audio.participants',
    'meeting.participants',
  ]);
  if (Array.isArray(participants)) {
    return participants.find((participant) => firstBoolean(participant, [
      'speaking',
      'is_speaking',
      'isSpeaking',
      'active_speaker',
      'activeSpeaker',
      'active',
    ]));
  }
  return null;
}

function speakerSpeaking(input = {}, speaker = null) {
  return firstBoolean(input, [
    'activeSpeaker.speaking',
    'active_speaker.speaking',
    'speaker.speaking',
    'accessibility.activeSpeaker.speaking',
    'accessibility.active_speaker.speaking',
    'audio.activeSpeaker.speaking',
    'audio.active_speaker.speaking',
    'window.activeSpeaker.speaking',
    'speaking',
    'is_speaking',
    'isSpeaking',
  ]) ?? firstBoolean(speaker ?? {}, [
    'speaking',
    'is_speaking',
    'isSpeaking',
    'active_speaker',
    'activeSpeaker',
    'active',
  ]);
}

function meetingForSpeaker(input = {}, candidate = null) {
  return candidate?.meeting ?? compactObject({
    platform: explicitPlatform(input),
    meeting_id: explicitMeetingId(input),
    meeting_url: urlInput(input),
    title: titleInput(input),
  });
}

export function normalizeNativeActiveSpeakerSample(input = {}, options = {}) {
  const candidate = options.candidate ?? normalizeNativeMeetingCandidate(input, options);
  const meeting = meetingForSpeaker(input, candidate);
  if (!meeting?.meeting_id && !meeting?.meeting_url) return null;
  const speaker = speakerObject(input);
  const speaking = speakerSpeaking(input, speaker);
  return normalizeActiveSpeakerSample(compactObject({
    meeting,
    url: meeting.meeting_url ?? urlInput(input),
    title: meeting.title ?? titleInput(input),
    detected_platform: meeting.platform,
    activeSpeaker: speaker ? {
      id: firstPath(speaker, ['id', 'speaker_id', 'speakerId', 'participant_id', 'participantId', 'user_id', 'userId']),
      name: firstPath(speaker, ['name', 'display_name', 'displayName', 'user_name', 'userName', 'participant_name', 'participantName']),
      speaking,
    } : undefined,
    speaking: speaker ? speaking : false,
    observedAtMs: observedAtMs(input, options),
    source_event_id: firstPath(input, ['source_event_id', 'sourceEventId', 'event_id', 'eventId', 'id'])
      ? `${firstPath(input, ['source_event_id', 'sourceEventId', 'event_id', 'eventId', 'id'])}:active-speaker`
      : undefined,
    raw: input,
  }), options);
}

function selectedRaw(selection = {}) {
  return selection.selectedSnapshot?.candidate_raw
    ?? selection.selectedSnapshot
    ?? null;
}

function inputForNoCandidate(input = {}, options = {}) {
  return compactObject({
    observedAtMs: observedAtMs(input, options) ?? options.observedAtMs ?? options.receivedAtMs,
    raw: input,
  });
}

export function observeNativeMeetingSample(state = null, input = {}, options = {}) {
  const sessionObserver = createMeetingSessionDiscoveryObserver({
    ...(options.sessionOptions ?? {}),
    initialState: state?.sessionState,
    source: options.sessionSource ?? options.source ?? 'native_meeting_observer',
  });
  const speakerObserver = createActiveSpeakerObserver({
    ...(options.speakerOptions ?? {}),
    initialState: state?.speakerState,
    source: options.speakerSource ?? options.source ?? 'native_meeting_observer',
  });
  const candidates = normalizeNativeMeetingCandidates(input, options);
  const session = sessionObserver.observeEnvironment(candidates, options.sessionObserveOptions ?? options);
  const raw = selectedRaw(session.selection);
  const speakerInput = raw ?? inputForNoCandidate(input, options);
  const speakerSample = raw
    ? normalizeNativeActiveSpeakerSample(speakerInput, {
      ...options,
      candidate: session.selection.selectedSnapshot,
    })
    : null;
  const speaker = speakerObserver.observe(speakerSample ?? inputForNoCandidate(input, options), options.speakerObserveOptions ?? options);
  return {
    state: {
      sessionState: session.state,
      speakerState: speaker.state,
    },
    signals: [...session.signals, ...speaker.signals],
    session,
    speaker,
    normalizedCandidates: candidates,
  };
}

export function createNativeMeetingObserver(options = {}) {
  const sessionObserver = createMeetingSessionDiscoveryObserver({
    ...(options.sessionOptions ?? {}),
    source: options.sessionSource ?? options.source ?? 'native_meeting_observer',
  });
  const speakerObserver = createActiveSpeakerObserver({
    ...(options.speakerOptions ?? {}),
    source: options.speakerSource ?? options.source ?? 'native_meeting_observer',
  });
  return {
    observe(input = {}, observeOptions = {}) {
      const candidates = normalizeNativeMeetingCandidates(input, {
        ...options,
        ...observeOptions,
      });
      const session = sessionObserver.observeEnvironment(candidates, observeOptions.sessionObserveOptions ?? observeOptions);
      const raw = selectedRaw(session.selection);
      const speakerInput = raw ?? inputForNoCandidate(input, observeOptions);
      const speakerSample = raw
        ? normalizeNativeActiveSpeakerSample(speakerInput, {
          ...options,
          ...observeOptions,
          candidate: session.selection.selectedSnapshot,
        })
        : null;
      const speaker = speakerObserver.observe(speakerSample ?? inputForNoCandidate(input, observeOptions), observeOptions.speakerObserveOptions ?? observeOptions);
      return {
        state: {
          sessionState: session.state,
          speakerState: speaker.state,
        },
        signals: [...session.signals, ...speaker.signals],
        session,
        speaker,
        normalizedCandidates: candidates,
      };
    },
    getState() {
      return {
        sessionState: sessionObserver.getState(),
        speakerState: speakerObserver.getState(),
      };
    },
    reset(nextState = null) {
      return {
        sessionState: sessionObserver.reset(nextState?.sessionState),
        speakerState: speakerObserver.reset(nextState?.speakerState),
      };
    },
  };
}

export function createNativeMeetingTimelineObserver(client, options = {}) {
  const sessionObserver = createMeetingSessionTimelineDiscovery(client, {
    ...(options.sessionOptions ?? {}),
    source: options.sessionSource ?? options.source ?? 'native_meeting_observer',
    applyOptions: options.sessionApplyOptions ?? options.applyOptions,
  });
  const speakerObserver = createActiveSpeakerTimelineObserver(client, {
    ...(options.speakerOptions ?? {}),
    source: options.speakerSource ?? options.source ?? 'native_meeting_observer',
    applyOptions: options.speakerApplyOptions ?? options.applyOptions,
  });
  return {
    async observe(input = {}, observeOptions = {}) {
      const candidates = normalizeNativeMeetingCandidates(input, {
        ...options,
        ...observeOptions,
      });
      const session = await sessionObserver.observeEnvironment(candidates, observeOptions.sessionObserveOptions ?? observeOptions);
      const raw = selectedRaw(session.selection);
      const speakerInput = raw ?? inputForNoCandidate(input, observeOptions);
      const speakerSample = raw
        ? normalizeNativeActiveSpeakerSample(speakerInput, {
          ...options,
          ...observeOptions,
          candidate: session.selection.selectedSnapshot,
        })
        : null;
      const speaker = await speakerObserver.observe(speakerSample ?? inputForNoCandidate(input, observeOptions), observeOptions.speakerObserveOptions ?? observeOptions);
      return {
        state: {
          sessionState: session.state,
          speakerState: speaker.state,
        },
        signals: [...session.signals, ...speaker.signals],
        results: [...session.results, ...speaker.results],
        session,
        speaker,
        normalizedCandidates: candidates,
      };
    },
    getState() {
      return {
        sessionState: sessionObserver.getState(),
        speakerState: speakerObserver.getState(),
      };
    },
    reset(nextState = null) {
      return {
        sessionState: sessionObserver.reset(nextState?.sessionState),
        speakerState: speakerObserver.reset(nextState?.speakerState),
      };
    },
  };
}
