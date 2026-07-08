import {
  MeetingTimelineSdkError,
  compactObject,
  normalizeAbsoluteMs,
} from '../index.mjs';
import { detectMeetingFromUrl } from './meeting-url.mjs';
import {
  buildMeetingPlatformAnnotationRuntimeEvent,
  buildMeetingPlatformCandidateObservationRuntimeEvent,
  buildMeetingPlatformObserveRuntimeEvent,
  buildMeetingPlatformParticipantTrackRuntimeEvent,
  buildMeetingPlatformProviderRuntimeEvent,
  buildMeetingPlatformSpeakerTrackRuntimeEvent,
} from './platform-runtime-event.mjs';
import {
  buildMeetingPlatformSpeakerTrack,
} from './platform-speaker-track.mjs';
import {
  MEETING_PLATFORM_KEYS,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_RAW_SIGNAL_SCHEMA = 'meeting_platform_raw_signal';
export const MEETING_PLATFORM_RAW_SIGNAL_BATCH_SCHEMA = 'meeting_platform_raw_signal_batch';
export const MEETING_PLATFORM_RAW_SIGNAL_EXAMPLE_BATCH_SCHEMA = 'meeting_platform_raw_signal_example_batch';
export const MEETING_PLATFORM_RAW_SIGNAL_SCHEMA_VERSION = 1;

export const MEETING_PLATFORM_RAW_SIGNAL_KINDS = Object.freeze([
  'meeting_app_snapshot',
  'platform_candidates',
  'provider_event',
  'annotation',
  'speaker_track',
  'participant_track',
]);

const RAW_SIGNAL_KIND_ALIASES = new Map([
  ['snapshot', 'meeting_app_snapshot'],
  ['dom_snapshot', 'meeting_app_snapshot'],
  ['browser_snapshot', 'meeting_app_snapshot'],
  ['meeting_app_snapshot', 'meeting_app_snapshot'],
  ['observe_meeting_app', 'meeting_app_snapshot'],
  ['observe', 'meeting_app_snapshot'],
  ['candidate', 'platform_candidates'],
  ['candidates', 'platform_candidates'],
  ['platform_candidate', 'platform_candidates'],
  ['platform_candidates', 'platform_candidates'],
  ['observe_candidates', 'platform_candidates'],
  ['observe_platform_candidates', 'platform_candidates'],
  ['window_candidates', 'platform_candidates'],
  ['tabs', 'platform_candidates'],
  ['windows', 'platform_candidates'],
  ['provider', 'provider_event'],
  ['provider_event', 'provider_event'],
  ['webhook', 'provider_event'],
  ['lifecycle', 'provider_event'],
  ['annotation', 'annotation'],
  ['mark', 'annotation'],
  ['insert_annotation', 'annotation'],
  ['insert_mark', 'annotation'],
  ['speaker', 'speaker_track'],
  ['speaker_signal', 'speaker_track'],
  ['speaker_track', 'speaker_track'],
  ['active_speaker', 'speaker_track'],
  ['participant', 'participant_track'],
  ['participant_signal', 'participant_track'],
  ['participant_track', 'participant_track'],
  ['roster', 'participant_track'],
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

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date);
}

function getPath(raw, path) {
  const parts = path.split('.');
  let node = raw;
  for (const part of parts) node = node?.[part];
  return node;
}

function firstPath(raw, paths = []) {
  return firstNonEmpty(...paths.map((path) => getPath(raw, path)));
}

function boolOption(options = {}, keys = [], fallback = false) {
  const value = firstNonEmpty(...keys.map((key) => options[key]));
  if (value == null || value === '') return fallback;
  if (typeof value === 'boolean') return value;
  return ['1', 'true', 'yes', 'on'].includes(String(value).trim().toLowerCase());
}

function maybeAbsoluteMs(value, fieldName = 'occurred_at_ms') {
  if (value == null || value === '') return undefined;
  return normalizeAbsoluteMs(value, fieldName);
}

function nowMs(options = {}) {
  const now = firstNonEmpty(options.now, options.clock);
  if (typeof now === 'function') return maybeAbsoluteMs(now(), 'observed_at_ms');
  return Date.now();
}

function normalizedPlatform(value) {
  if (value == null || value === '') return undefined;
  try {
    return normalizeMeetingPlatform(value);
  } catch {
    return undefined;
  }
}

function meetingUrl(input = {}) {
  return firstPath(input, [
    'meeting_url',
    'meetingUrl',
    'join_url',
    'joinUrl',
    'url',
    'href',
    'snapshot.url',
    'snapshot.href',
    'tab.url',
    'window.url',
    'browser.url',
    'meeting.meeting_url',
    'meeting.meetingUrl',
    'current_meeting.meeting_url',
    'currentMeeting.meetingUrl',
  ]);
}

function meetingTitle(input = {}) {
  return firstPath(input, [
    'meeting_title',
    'meetingTitle',
    'title',
    'topic',
    'name',
    'snapshot.title',
    'tab.title',
    'window.title',
    'meeting.title',
    'current_meeting.title',
    'currentMeeting.title',
  ]);
}

function platformFromInput(input = {}, options = {}) {
  const explicit = firstNonEmpty(
    options.platform,
    input.platform,
    input.provider,
    input.adapter,
    input.meeting_platform,
    input.meetingPlatform,
    input.snapshot?.platform,
    input.meeting?.platform,
    input.current_meeting?.platform,
    input.currentMeeting?.platform,
    input.provider_event?.platform,
    input.providerEvent?.platform,
  );
  const normalized = normalizedPlatform(explicit);
  if (normalized) return normalized;
  const detected = detectMeetingFromUrl({ url: meetingUrl(input), title: meetingTitle(input) });
  return detected?.platform;
}

function timestampFor(input = {}, options = {}) {
  return maybeAbsoluteMs(firstNonEmpty(
    input.occurred_at_ms,
    input.occurredAtMs,
    input.observed_at_ms,
    input.observedAtMs,
    input.captured_at_ms,
    input.capturedAtMs,
    input.sent_at_ms,
    input.sentAtMs,
    input.timestamp_ms,
    input.timestampMs,
    input.time_ms,
    input.timeMs,
    input.event_ts,
    input.eventTs,
    input.timestamp,
    input.time,
    input.ts,
    input.snapshot?.observed_at_ms,
    input.snapshot?.observedAtMs,
    input.provider_event?.event_ts,
    input.providerEvent?.eventTs,
    options.observed_at_ms,
    options.observedAtMs,
    options.captured_at_ms,
    options.capturedAtMs,
    options.sent_at_ms,
    options.sentAtMs,
    nowMs(options),
  ), 'raw_signal_time');
}

function normalizeKindValue(kind) {
  const normalized = String(kind ?? '').trim().toLowerCase().replace(/[-.\s]+/g, '_');
  return RAW_SIGNAL_KIND_ALIASES.get(normalized) ?? normalized;
}

function hasAny(input = {}, keys = []) {
  return keys.some((key) => getPath(input, key) != null);
}

function inferRawSignalKind(input = {}, options = {}) {
  const explicit = normalizeKindValue(firstNonEmpty(
    options.kind,
    options.action,
    input.kind,
    input.action,
    input.signal_kind,
    input.signalKind,
    input.runtime_action,
    input.runtimeAction,
    input.type,
  ));
  if (MEETING_PLATFORM_RAW_SIGNAL_KINDS.includes(explicit)) return explicit;
  if (hasAny(input, ['annotation', 'mark', 'label', 'ink', 'strokes', 'captured_at_ms', 'capturedAtMs'])) return 'annotation';
  if (hasAny(input, ['snapshot', 'dom', 'url', 'href', 'tab.url', 'window.url', 'browser.url', 'in_meeting', 'inMeeting'])) return 'meeting_app_snapshot';
  if (hasAny(input, ['speaker', 'active_speaker', 'activeSpeaker', 'speaker_signal', 'speakerSignal'])) return 'speaker_track';
  if (hasAny(input, ['participant', 'participants', 'roster', 'participant_signal', 'participantSignal'])) return 'participant_track';
  if (hasAny(input, ['provider_event', 'providerEvent', 'webhook', 'event_type', 'eventType', 'event.event_type', 'payload.event_type'])) return 'provider_event';
  if (hasAny(input, ['windows', 'tabs', 'applications', 'apps', 'candidates', 'items'])) return 'platform_candidates';
  throw new MeetingTimelineSdkError('Cannot infer meeting platform raw signal kind', {
    supported_kinds: MEETING_PLATFORM_RAW_SIGNAL_KINDS,
    input_keys: Object.keys(input ?? {}),
  });
}

export function normalizeMeetingPlatformRawSignalKind(kind) {
  const normalized = normalizeKindValue(kind);
  if (!MEETING_PLATFORM_RAW_SIGNAL_KINDS.includes(normalized)) {
    throw new MeetingTimelineSdkError('Unsupported meeting platform raw signal kind', {
      kind,
      normalized,
      supported_kinds: MEETING_PLATFORM_RAW_SIGNAL_KINDS,
    });
  }
  return normalized;
}

function sourceFor(kind, input = {}, options = {}) {
  return firstNonEmpty(
    options.source,
    input.source,
    input.detector_source,
    input.detectorSource,
    {
      meeting_app_snapshot: 'browser_or_native_observer',
      platform_candidates: 'local_detector',
      provider_event: 'provider_webhook',
      annotation: 'annotation_device',
      speaker_track: 'local_observer',
      participant_track: 'local_observer',
    }[kind],
  );
}

function eventIdFor(input = {}, options = {}) {
  return firstNonEmpty(
    options.event_id,
    options.eventId,
    input.event_id,
    input.eventId,
    input.source_event_id,
    input.sourceEventId,
    input.id,
  );
}

function currentMeeting(input = {}, platform) {
  const meeting = firstNonEmpty(input.current_meeting, input.currentMeeting, input.meeting, {});
  const url = meetingUrl(input);
  const detected = detectMeetingFromUrl({ url, title: meetingTitle(input) }) ?? {};
  const meetingId = firstNonEmpty(
    meeting.meeting_id,
    meeting.meetingId,
    input.meeting_id,
    input.meetingId,
    input.session_id,
    input.sessionId,
    detected.meeting_id,
  );
  return compactObject({
    platform: firstNonEmpty(normalizedPlatform(meeting.platform), platform),
    meeting_id: meetingId,
    external_meeting_id: firstNonEmpty(meeting.external_meeting_id, meeting.externalMeetingId, input.external_meeting_id, input.externalMeetingId),
    meeting_url: firstNonEmpty(meeting.meeting_url, meeting.meetingUrl, url),
    title: firstNonEmpty(meeting.title, meetingTitle(input)),
    start_time_ms: firstNonEmpty(meeting.start_time_ms, meeting.startTimeMs, input.start_time_ms, input.startTimeMs),
  });
}

function primaryPayload(kind, input = {}) {
  if (kind === 'meeting_app_snapshot') {
    return firstNonEmpty(input.snapshot, input.meeting_app_snapshot, input.meetingAppSnapshot, input.dom, input);
  }
  if (kind === 'platform_candidates') {
    return {
      windows: firstNonEmpty(input.windows, input.payload?.windows),
      tabs: firstNonEmpty(input.tabs, input.payload?.tabs),
      applications: firstNonEmpty(input.applications, input.apps, input.payload?.applications, input.payload?.apps),
      candidates: firstNonEmpty(input.candidates, input.items, input.payload?.candidates, input.payload?.items),
    };
  }
  if (kind === 'provider_event') {
    return firstNonEmpty(input.provider_event, input.providerEvent, input.webhook, input.event, input.payload, input.raw, input);
  }
  if (kind === 'annotation') {
    return firstNonEmpty(input.annotation, input.mark, input.payload?.annotation, input.payload?.mark, input);
  }
  if (kind === 'speaker_track') {
    return firstNonEmpty(input.signals, input.signal, input.speaker_signal, input.speakerSignal, input.speaker, input.active_speaker, input.activeSpeaker, input);
  }
  if (kind === 'participant_track') {
    return firstNonEmpty(input.signals, input.signal, input.participant_signal, input.participantSignal, input.participant, input.participants, input.roster, input);
  }
  return input;
}

function signalArray(kind, input = {}, payload = undefined, atMs = undefined, platform = undefined, source = undefined) {
  const rawSignals = firstNonEmpty(input.signals, input.signal, payload);
  if (Array.isArray(rawSignals)) return rawSignals;
  if (kind === 'speaker_track') {
    const speaker = firstNonEmpty(input.speaker, input.active_speaker, input.activeSpeaker, payload);
    return [compactObject({
      type: firstNonEmpty(input.signal_type, input.signalType, input.type, 'speaker_started'),
      platform,
      speaker: isPlainObject(speaker) ? speaker : { display_name: String(speaker ?? 'unknown') },
      occurred_at_ms: atMs,
      source,
      confidence: input.confidence,
    })];
  }
  if (kind === 'participant_track') {
    const participant = firstNonEmpty(input.participant, payload);
    if (Array.isArray(input.participants)) {
      return input.participants.map((item) => compactObject({
        type: firstNonEmpty(input.signal_type, input.signalType, 'participant_joined'),
        platform,
        participant: item,
        occurred_at_ms: atMs,
        source,
      }));
    }
    return [compactObject({
      type: firstNonEmpty(input.signal_type, input.signalType, input.type, 'participant_joined'),
      platform,
      participant: isPlainObject(participant) ? participant : { display_name: String(participant ?? 'unknown') },
      occurred_at_ms: atMs,
      source,
      confidence: input.confidence,
    })];
  }
  return [];
}

function runtimeEventsFor(kind, input = {}, context = {}, options = {}) {
  const {
    platform,
    source,
    atMs,
    payload,
    current_meeting: meeting,
  } = context;
  const commonOptions = compactObject({
    ...options,
    source,
    sent_at_ms: atMs,
    event_id: context.event_id,
    correlation_id: firstNonEmpty(input.correlation_id, input.correlationId, meeting?.meeting_id),
  });
  if (kind === 'meeting_app_snapshot') {
    const snapshot = compactObject({
      ...payload,
      platform,
      observed_at_ms: firstNonEmpty(payload?.observed_at_ms, payload?.observedAtMs, atMs),
      url: firstNonEmpty(payload?.url, meetingUrl(input)),
      title: firstNonEmpty(payload?.title, meetingTitle(input)),
    });
    const events = [buildMeetingPlatformObserveRuntimeEvent(platform, snapshot, commonOptions)];
    const deriveSpeakerTrack = boolOption(options, ['deriveSpeakerTrack', 'derive_speaker_track'], true);
    const activeSpeaker = firstNonEmpty(snapshot.active_speaker, snapshot.activeSpeaker);
    if (deriveSpeakerTrack && activeSpeaker) {
      events.push(buildMeetingPlatformSpeakerTrackRuntimeEvent(platform, {
        current_meeting: meeting,
        signals: signalArray('speaker_track', {
          ...input,
          speaker: activeSpeaker,
          signal_type: 'speaker_started',
        }, activeSpeaker, atMs, platform, source),
      }, commonOptions));
    }
    return events;
  }
  if (kind === 'platform_candidates') {
    return [buildMeetingPlatformCandidateObservationRuntimeEvent(compactObject({
      ...payload,
      platform,
      observed_at_ms: atMs,
    }), commonOptions)];
  }
  if (kind === 'provider_event') {
    return [buildMeetingPlatformProviderRuntimeEvent(platform, payload, commonOptions)];
  }
  if (kind === 'annotation') {
    return [buildMeetingPlatformAnnotationRuntimeEvent(platform, {
      annotation: compactObject({
        ...payload,
        platform,
        captured_at_ms: firstNonEmpty(payload?.captured_at_ms, payload?.capturedAtMs, atMs),
      }),
      current_meeting: meeting,
    }, commonOptions)];
  }
  if (kind === 'speaker_track') {
    return [buildMeetingPlatformSpeakerTrackRuntimeEvent(platform, {
      current_meeting: meeting,
      signals: signalArray(kind, input, payload, atMs, platform, source),
    }, commonOptions)];
  }
  if (kind === 'participant_track') {
    return [buildMeetingPlatformParticipantTrackRuntimeEvent(platform, {
      current_meeting: meeting,
      signals: signalArray(kind, input, payload, atMs, platform, source),
    }, commonOptions)];
  }
  return [];
}

export function buildMeetingPlatformRawSignal(input = {}, options = {}) {
  const normalizedInput = typeof input === 'string' || input instanceof URL ? { url: String(input) } : (input ?? {});
  const kind = normalizeMeetingPlatformRawSignalKind(inferRawSignalKind(normalizedInput, options));
  const platform = platformFromInput(normalizedInput, options);
  if (!platform && kind !== 'platform_candidates') {
    throw new MeetingTimelineSdkError('Meeting platform raw signal requires a platform or detectable meeting URL', {
      kind,
      required_field: 'platform_or_meeting_url',
    });
  }
  const atMs = timestampFor(normalizedInput, options);
  const source = sourceFor(kind, normalizedInput, options);
  const payload = primaryPayload(kind, normalizedInput);
  const meeting = currentMeeting(normalizedInput, platform);
  const eventId = eventIdFor(normalizedInput, options);
  const runtimeEvents = runtimeEventsFor(kind, normalizedInput, {
    platform,
    source,
    atMs,
    payload,
    current_meeting: meeting,
    event_id: eventId,
  }, options);
  return compactObject({
    type: MEETING_PLATFORM_RAW_SIGNAL_SCHEMA,
    schema: MEETING_PLATFORM_RAW_SIGNAL_SCHEMA,
    schema_version: MEETING_PLATFORM_RAW_SIGNAL_SCHEMA_VERSION,
    kind,
    platform,
    source,
    event_id: eventId,
    observed_at_ms: atMs,
    current_meeting: meeting,
    payload,
    runtime_action: runtimeEvents[0]?.action,
    runtime_event_count: runtimeEvents.length,
    runtime_events: runtimeEvents,
  });
}

export function buildMeetingPlatformRuntimeEventsFromRawSignal(input = {}, options = {}) {
  return buildMeetingPlatformRawSignal(input, options).runtime_events;
}

function rawSignalInputs(input = {}) {
  if (Array.isArray(input)) return input;
  const explicitBatch = firstNonEmpty(
    input.raw_signals,
    input.rawSignals,
    input.events,
    input.samples,
    input.records,
  );
  return explicitBatch == null ? [input] : asArray(explicitBatch);
}

function shouldFilterActiveSpeakerSamples(options = {}) {
  return boolOption(options, ['filterActiveSpeakerSamples', 'filter_active_speaker_samples'], false);
}

function hasActiveSpeakerSample(input = {}) {
  return firstNonEmpty(
    input.active_speaker,
    input.activeSpeaker,
    input.speaker,
    input.snapshot?.active_speaker,
    input.snapshot?.activeSpeaker,
  ) != null;
}

function speakerSampleInput(input = {}) {
  const snapshot = isPlainObject(input.snapshot) ? input.snapshot : {};
  return compactObject({
    ...snapshot,
    ...input,
    active_speaker: firstNonEmpty(input.active_speaker, input.activeSpeaker, snapshot.active_speaker, snapshot.activeSpeaker, input.speaker),
    observed_at_ms: firstNonEmpty(input.observed_at_ms, input.observedAtMs, snapshot.observed_at_ms, snapshot.observedAtMs, input.occurred_at_ms, input.occurredAtMs),
    meeting: firstNonEmpty(input.current_meeting, input.currentMeeting, input.meeting, snapshot.meeting),
    url: firstNonEmpty(input.url, input.href, snapshot.url, snapshot.href, meetingUrl(input)),
    title: firstNonEmpty(input.title, snapshot.title, meetingTitle(input)),
    platform: firstNonEmpty(input.platform, snapshot.platform, platformFromInput(input)),
  });
}

function groupSpeakerSamplesByPlatform(inputs = [], options = {}) {
  const rows = new Map();
  for (const input of inputs) {
    const kind = inferRawSignalKind(input, options);
    if (kind !== 'meeting_app_snapshot' && kind !== 'speaker_track') continue;
    if (!hasActiveSpeakerSample(input)) continue;
    const platform = platformFromInput(input, options);
    if (!platform) continue;
    const list = rows.get(platform) ?? [];
    list.push(speakerSampleInput(input));
    rows.set(platform, list);
  }
  return rows;
}

function filteredSpeakerRuntimeEvents(inputs = [], options = {}) {
  const events = [];
  for (const [platform, samples] of groupSpeakerSamplesByPlatform(inputs, options).entries()) {
    const track = buildMeetingPlatformSpeakerTrack(platform, {
      samples,
    }, options);
    if ((track.signals ?? []).length === 0) continue;
    const meeting = firstNonEmpty(
      track.signals.find((signal) => signal.meeting)?.meeting,
      track.segments?.find?.((segment) => segment.meeting)?.meeting,
    );
    events.push(buildMeetingPlatformSpeakerTrackRuntimeEvent(platform, {
      current_meeting: meeting,
      signals: track.signals,
      speaker_track: track,
      filter_policy: track.filter_policy,
      diagnostics: track.diagnostics,
    }, {
      ...options,
      source: firstNonEmpty(options.source, options.detectorSource, options.detector_source, 'filtered_active_speaker_observer'),
      sent_at_ms: track.signals[0]?.occurred_at_ms,
    }));
  }
  return events;
}

export function buildMeetingPlatformRawSignalBatch(input = {}, options = {}) {
  const inputs = rawSignalInputs(input);
  const filterSpeakers = shouldFilterActiveSpeakerSamples(options);
  const signalOptions = filterSpeakers
    ? { ...options, deriveSpeakerTrack: false, derive_speaker_track: false }
    : options;
  const signals = inputs.map((signal) => buildMeetingPlatformRawSignal(signal, signalOptions));
  const filteredSpeakerEvents = filterSpeakers ? filteredSpeakerRuntimeEvents(inputs, options) : [];
  const runtimeEvents = [
    ...signals.flatMap((signal) => signal.runtime_events ?? []),
    ...filteredSpeakerEvents,
  ];
  const rows = signals.map((signal) => ({
    kind: signal.kind,
    platform: signal.platform,
    source: signal.source,
    runtime_event_count: signal.runtime_event_count,
    runtime_actions: unique((signal.runtime_events ?? []).map((event) => event.action)),
  }));
  return {
    type: MEETING_PLATFORM_RAW_SIGNAL_BATCH_SCHEMA,
    schema: MEETING_PLATFORM_RAW_SIGNAL_BATCH_SCHEMA,
    schema_version: MEETING_PLATFORM_RAW_SIGNAL_SCHEMA_VERSION,
    signal_count: signals.length,
    runtime_event_count: runtimeEvents.length,
    filtered_speaker_event_count: filteredSpeakerEvents.length,
    platforms: unique(signals.map((signal) => signal.platform)),
    kinds: unique(signals.map((signal) => signal.kind)),
    rows,
    signals,
    runtime_events: runtimeEvents,
  };
}

export function buildMeetingPlatformRawSignalExamples(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const baseMs = maybeAbsoluteMs(firstNonEmpty(options.sample_at_ms, options.sampleAtMs, 1_782_614_400_000), 'sample_at_ms');
  const url = {
    google_meet: 'https://meet.google.com/abc-defg-hij',
    microsoft_teams: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
    zoom: 'https://zoom.us/j/987654321',
    webex: 'https://example.webex.com/meet/sample',
    lark: 'https://vc.feishu.cn/j/123456789',
    local_detector: 'local://meeting-window/sample',
  }[key];
  const meeting = {
    platform: key,
    meeting_id: `${key}-sample-meeting`,
    external_meeting_id: `${key}-external-sample`,
    meeting_url: url,
    title: `${key} sample meeting`,
    start_time_ms: baseMs,
  };
  return [
    {
      kind: 'meeting_app_snapshot',
      platform: key,
      source: key === 'google_meet' ? 'browser_dom_observer' : 'native_or_browser_observer',
      url,
      title: meeting.title,
      observed_at_ms: baseMs,
      in_meeting: true,
      active_speaker: { id: 'speaker-001', display_name: 'Alex' },
      current_meeting: meeting,
    },
    {
      kind: 'annotation',
      platform: key,
      source: 'annotation_device',
      label: 'why?',
      captured_at_ms: baseMs + 15_000,
      current_meeting: meeting,
    },
    {
      kind: 'speaker_track',
      platform: key,
      source: 'local_observer',
      speaker: { id: 'speaker-001', display_name: 'Alex' },
      occurred_at_ms: baseMs + 20_000,
      current_meeting: meeting,
    },
  ];
}

export function buildMeetingPlatformRawSignalExampleBatch(options = {}) {
  const platforms = unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_PLATFORM_KEYS))
    .map((platform) => normalizeMeetingPlatform(platform)));
  const rawSignals = platforms.flatMap((platform) => buildMeetingPlatformRawSignalExamples(platform, options));
  return {
    ...buildMeetingPlatformRawSignalBatch(rawSignals, options),
    type: MEETING_PLATFORM_RAW_SIGNAL_EXAMPLE_BATCH_SCHEMA,
    schema: MEETING_PLATFORM_RAW_SIGNAL_EXAMPLE_BATCH_SCHEMA,
    platform_count: platforms.length,
  };
}
