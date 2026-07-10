import { compactObject, normalizeAbsoluteMs } from './internal-utils.mjs';
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
  return value == null ? undefined : normalizeAbsoluteMs(value, 'browser_meeting_observed_at_ms');
}

function browserRows(input = {}, inherited = {}) {
  if (Array.isArray(input)) return input.flatMap((item) => browserRows(item, inherited));
  if (!input || typeof input !== 'object') return [];
  const rows = [];
  const base = compactObject({
    ...inherited,
    browser: input.browser ?? inherited.browser,
    page: input.page ?? inherited.page,
    dom: input.dom ?? inherited.dom,
    application: input.application ?? input.app ?? inherited.application,
    app: input.app ?? input.application ?? inherited.app,
  });

  if (Array.isArray(input.windows)) {
    for (const windowSnapshot of input.windows) {
      rows.push(...browserRows(windowSnapshot, {
        ...base,
        window: windowSnapshot,
        browser: windowSnapshot.browser ?? base.browser,
      }));
    }
  }
  if (Array.isArray(input.tabs)) {
    for (const tab of input.tabs) {
      rows.push(...browserRows(tab, {
        ...base,
        tab,
        window: input.window ?? base.window,
      }));
    }
  }
  if (Array.isArray(input.pages)) {
    for (const page of input.pages) rows.push(...browserRows(page, { ...base, page }));
  }
  if (Array.isArray(input.frames)) {
    for (const frame of input.frames) rows.push(...browserRows(frame, { ...base, frame }));
  }
  if (Array.isArray(input.candidates)) {
    for (const candidate of input.candidates) rows.push(...browserRows(candidate, base));
  }
  if (!Array.isArray(input.windows) && !Array.isArray(input.tabs) && !Array.isArray(input.pages) && !Array.isArray(input.frames) && !Array.isArray(input.candidates)) {
    rows.push(compactObject({ ...base, ...input }));
  }
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
    'href',
    'location.href',
    'page.url',
    'page.href',
    'dom.url',
    'window.url',
    'browser.url',
    'tab.url',
    'frame.url',
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
    'document_title',
    'documentTitle',
    'page.title',
    'dom.title',
    'window.title',
    'tab.title',
    'frame.title',
  ]);
}

function inMeetingFlag(input = {}) {
  return firstBoolean(input, [
    'in_meeting',
    'inMeeting',
    'meeting.active',
    'meeting.in_meeting',
    'meeting.inMeeting',
    'page.in_meeting',
    'page.inMeeting',
    'page.has_joined',
    'page.hasJoined',
    'page.call_active',
    'page.callActive',
    'dom.in_meeting',
    'dom.inMeeting',
    'dom.has_joined',
    'dom.hasJoined',
    'call.active',
    'callActive',
    'call_active',
  ]);
}

function visibleFlag(input = {}) {
  return firstBoolean(input, [
    'visible',
    'tab.visible',
    'window.visible',
    'page.visible',
    'page.document_visible',
    'page.documentVisible',
    'dom.visible',
  ]);
}

function activeFlag(input = {}) {
  return firstBoolean(input, [
    'active',
    'focused',
    'selected',
    'tab.active',
    'tab.highlighted',
    'window.focused',
    'window.active',
    'page.focused',
  ]);
}

function audibleFlag(input = {}) {
  return firstBoolean(input, ['audible', 'tab.audible', 'page.audible', 'dom.audible', 'audio.active']);
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
    'page.platform',
    'dom.platform',
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
    'page.meeting_id',
    'page.meetingId',
    'dom.meeting_id',
    'dom.meetingId',
  ]);
}

export function normalizeBrowserMeetingCandidate(input = {}, options = {}) {
  const atMs = observedAtMs(input, options);
  return normalizeMeetingSessionCandidate(compactObject({
    ...input,
    url: urlInput(input),
    title: titleInput(input),
    active: activeFlag(input),
    visible: visibleFlag(input),
    audible: audibleFlag(input),
    inMeeting: inMeetingFlag(input),
    platform: explicitPlatform(input),
    meeting_id: explicitMeetingId(input),
    observedAtMs: atMs,
    browser: input.browser,
    window: input.window,
    tab: input.tab,
    page: input.page,
    dom: input.dom,
  }), options);
}

export function normalizeBrowserMeetingCandidates(input = {}, options = {}) {
  return browserRows(input).map((row) => normalizeBrowserMeetingCandidate(row, options)).filter(Boolean);
}

export function selectBrowserMeetingCandidate(input = {}, options = {}) {
  const normalizedCandidates = normalizeBrowserMeetingCandidates(input, options);
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
    'page.activeSpeaker',
    'page.active_speaker',
    'page.speaker',
    'dom.activeSpeaker',
    'dom.active_speaker',
    'dom.speaker',
    'tab.activeSpeaker',
    'meeting.activeSpeaker',
  ]);
  if (direct) return direct;
  const participants = firstPath(input, ['participants', 'page.participants', 'dom.participants']);
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
    'page.activeSpeaker.speaking',
    'page.active_speaker.speaking',
    'dom.activeSpeaker.speaking',
    'dom.active_speaker.speaking',
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

export function normalizeBrowserActiveSpeakerSample(input = {}, options = {}) {
  const candidate = options.candidate ?? normalizeBrowserMeetingCandidate(input, options);
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

export function observeBrowserMeetingSample(state = null, input = {}, options = {}) {
  const sessionObserver = createMeetingSessionDiscoveryObserver({
    ...(options.sessionOptions ?? {}),
    initialState: state?.sessionState,
    source: options.sessionSource ?? options.source ?? 'browser_meeting_observer',
  });
  const speakerObserver = createActiveSpeakerObserver({
    ...(options.speakerOptions ?? {}),
    initialState: state?.speakerState,
    source: options.speakerSource ?? options.source ?? 'browser_meeting_observer',
  });
  const candidates = normalizeBrowserMeetingCandidates(input, options);
  const session = sessionObserver.observeEnvironment(candidates, options.sessionObserveOptions ?? options);
  const raw = selectedRaw(session.selection);
  const speakerInput = raw ?? inputForNoCandidate(input, options);
  const speakerSample = raw
    ? normalizeBrowserActiveSpeakerSample(speakerInput, {
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

export function createBrowserMeetingObserver(options = {}) {
  const sessionObserver = createMeetingSessionDiscoveryObserver({
    ...(options.sessionOptions ?? {}),
    source: options.sessionSource ?? options.source ?? 'browser_meeting_observer',
  });
  const speakerObserver = createActiveSpeakerObserver({
    ...(options.speakerOptions ?? {}),
    source: options.speakerSource ?? options.source ?? 'browser_meeting_observer',
  });
  return {
    observe(input = {}, observeOptions = {}) {
      const candidates = normalizeBrowserMeetingCandidates(input, {
        ...options,
        ...observeOptions,
      });
      const session = sessionObserver.observeEnvironment(candidates, observeOptions.sessionObserveOptions ?? observeOptions);
      const raw = selectedRaw(session.selection);
      const speakerInput = raw ?? inputForNoCandidate(input, observeOptions);
      const speakerSample = raw
        ? normalizeBrowserActiveSpeakerSample(speakerInput, {
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

export function createBrowserMeetingTimelineObserver(client, options = {}) {
  const sessionObserver = createMeetingSessionTimelineDiscovery(client, {
    ...(options.sessionOptions ?? {}),
    source: options.sessionSource ?? options.source ?? 'browser_meeting_observer',
    applyOptions: options.sessionApplyOptions ?? options.applyOptions,
  });
  const speakerObserver = createActiveSpeakerTimelineObserver(client, {
    ...(options.speakerOptions ?? {}),
    source: options.speakerSource ?? options.source ?? 'browser_meeting_observer',
    applyOptions: options.speakerApplyOptions ?? options.applyOptions,
  });
  return {
    async observe(input = {}, observeOptions = {}) {
      const candidates = normalizeBrowserMeetingCandidates(input, {
        ...options,
        ...observeOptions,
      });
      const session = await sessionObserver.observeEnvironment(candidates, observeOptions.sessionObserveOptions ?? observeOptions);
      const raw = selectedRaw(session.selection);
      const speakerInput = raw ?? inputForNoCandidate(input, observeOptions);
      const speakerSample = raw
        ? normalizeBrowserActiveSpeakerSample(speakerInput, {
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
