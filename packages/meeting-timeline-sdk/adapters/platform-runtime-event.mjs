import {
  MeetingTimelineApiError,
  MeetingTimelineSdkError,
  compactObject,
  normalizeAbsoluteMs,
} from '../index.mjs';
import {
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA = 'meeting_platform_runtime_event';
export const MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA_VERSION = 1;
export const MEETING_PLATFORM_RUNTIME_EVENT_ENDPOINT = '/api/meeting-platform/runtime-events';

export const MEETING_PLATFORM_RUNTIME_EVENT_ACTIONS = Object.freeze([
  'observe_meeting_app',
  'provider_event',
  'insert_annotation',
  'speaker_track',
  'participant_track',
  'timeline_view',
  'runtime_bundles',
  'registry',
  'manifest',
  'readiness',
  'handoff_readiness',
]);

const ACTION_ALIASES = new Map([
  ['observe', 'observe_meeting_app'],
  ['observe_app', 'observe_meeting_app'],
  ['observe_meeting_app', 'observe_meeting_app'],
  ['meeting_app_snapshot', 'observe_meeting_app'],
  ['snapshot', 'observe_meeting_app'],
  ['provider', 'provider_event'],
  ['provider_event', 'provider_event'],
  ['ingest_provider', 'provider_event'],
  ['webhook', 'provider_event'],
  ['annotation', 'insert_annotation'],
  ['mark', 'insert_annotation'],
  ['insert_annotation', 'insert_annotation'],
  ['insert_mark', 'insert_annotation'],
  ['speaker', 'speaker_track'],
  ['speaker_track', 'speaker_track'],
  ['participant', 'participant_track'],
  ['participant_track', 'participant_track'],
  ['view', 'timeline_view'],
  ['timeline_view', 'timeline_view'],
  ['runtime_bundle_matrix', 'runtime_bundles'],
  ['runtime_bundles', 'runtime_bundles'],
  ['platform_registry', 'registry'],
  ['registry', 'registry'],
  ['runtime_manifest', 'manifest'],
  ['integration_manifest', 'manifest'],
  ['manifest', 'manifest'],
  ['live_readiness', 'readiness'],
  ['readiness', 'readiness'],
  ['handoff_readiness', 'handoff_readiness'],
]);

const PLATFORM_REQUIRED_ACTIONS = new Set([
  'observe_meeting_app',
  'provider_event',
  'insert_annotation',
  'speaker_track',
  'participant_track',
  'timeline_view',
]);

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date);
}

function maybeAbsoluteMs(value, fieldName) {
  if (value == null || value === '') return undefined;
  return normalizeAbsoluteMs(value, fieldName);
}

function nowMs(options = {}) {
  const now = firstNonEmpty(options.now, options.clock);
  if (typeof now === 'function') return maybeAbsoluteMs(now(), 'sent_at_ms');
  return Date.now();
}

function valuePlatform(value) {
  if (value == null || value === '') return undefined;
  return normalizeMeetingPlatform(value);
}

function normalizedAction(value) {
  const key = String(value ?? '').trim().toLowerCase().replace(/[-.\s]+/g, '_');
  return ACTION_ALIASES.get(key) ?? key;
}

function payloadObject(payload) {
  return isPlainObject(payload) ? payload : {};
}

function eventPayload(input = {}, options = {}) {
  return firstNonEmpty(
    options.payload,
    input.payload,
    input.provider_event,
    input.providerEvent,
    input.snapshot,
    input.meeting_app_snapshot,
    input.meetingAppSnapshot,
    input.annotation,
    input.mark,
    input.raw,
  );
}

function eventIdFor(input = {}, options = {}) {
  return firstNonEmpty(
    options.event_id,
    options.eventId,
    input.event_id,
    input.eventId,
    input.id,
    input.source_event_id,
    input.sourceEventId,
  );
}

function correlationIdFor(input = {}, options = {}) {
  return firstNonEmpty(
    options.correlation_id,
    options.correlationId,
    input.correlation_id,
    input.correlationId,
    input.meeting_id,
    input.meetingId,
    input.meeting?.meeting_id,
    input.meeting?.meetingId,
    input.current_meeting?.meeting_id,
    input.currentMeeting?.meetingId,
  );
}

function sentAtFor(input = {}, options = {}) {
  return maybeAbsoluteMs(firstNonEmpty(
    options.sent_at_ms,
    options.sentAtMs,
    options.sent_at,
    options.sentAt,
    input.sent_at_ms,
    input.sentAtMs,
    input.sent_at,
    input.sentAt,
  ), 'sent_at_ms') ?? nowMs(options);
}

function platformFor(input = {}, options = {}, payload = undefined) {
  const value = firstNonEmpty(
    options.platform,
    input.platform,
    input.provider,
    input.adapter,
    payload?.platform,
    payload?.provider,
    payload?.meeting?.platform,
    payload?.current_meeting?.platform,
    payload?.currentMeeting?.platform,
  );
  return value == null ? undefined : valuePlatform(value);
}

function actionFor(input = {}, options = {}) {
  return normalizedAction(firstNonEmpty(
    options.action,
    options.kind,
    input.action,
    input.kind,
    input.event_kind,
    input.eventKind,
    input.runtime_action,
    input.runtimeAction,
    input.type,
  ));
}

function actionDefaults(event = {}, payload = undefined, input = {}) {
  const objectPayload = payloadObject(payload);
  if (event.action === 'observe_meeting_app' && !event.snapshot && !event.meeting_app_snapshot) {
    event.snapshot = firstNonEmpty(input.snapshot, input.meeting_app_snapshot, input.meetingAppSnapshot, input.dom, payload);
  }
  if (event.action === 'provider_event' && !event.provider_event && !event.providerEvent && !event.raw) {
    event.provider_event = payload ?? input.raw ?? objectPayload;
  }
  if (event.action === 'insert_annotation' && !event.annotation && !event.mark) {
    event.annotation = firstNonEmpty(input.annotation, input.mark, objectPayload.annotation, objectPayload.mark, payload);
  }
  if (['speaker_track', 'participant_track'].includes(event.action) && !event.signals) {
    event.signals = firstNonEmpty(input.signals, objectPayload.signals);
  }
  if (event.action === 'timeline_view') {
    event.meeting = firstNonEmpty(input.meeting, objectPayload.meeting, event.meeting);
    event.annotations = firstNonEmpty(input.annotations, objectPayload.annotations, event.annotations);
    event.speakerTrack = firstNonEmpty(input.speakerTrack, input.speaker_track, objectPayload.speakerTrack, objectPayload.speaker_track);
    event.participantTrack = firstNonEmpty(input.participantTrack, input.participant_track, objectPayload.participantTrack, objectPayload.participant_track);
  }
  return event;
}

export function normalizeMeetingPlatformRuntimeEventAction(action) {
  const normalized = normalizedAction(action);
  if (!MEETING_PLATFORM_RUNTIME_EVENT_ACTIONS.includes(normalized)) {
    throw new MeetingTimelineSdkError('Unsupported meeting platform runtime event action', {
      action,
      normalized,
      supported_actions: MEETING_PLATFORM_RUNTIME_EVENT_ACTIONS,
    });
  }
  return normalized;
}

export function buildMeetingPlatformRuntimeEvent(input = {}, options = {}) {
  const payload = eventPayload(input, options);
  const objectPayload = payloadObject(payload);
  const action = normalizeMeetingPlatformRuntimeEventAction(actionFor(input, options));
  const platform = platformFor(input, options, objectPayload);
  const event = actionDefaults(compactObject({
    ...objectPayload,
    ...input,
    type: MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA,
    schema: MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA,
    schema_version: MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA_VERSION,
    action,
    platform,
    source: firstNonEmpty(options.source, input.source),
    event_id: eventIdFor(input, options),
    correlation_id: correlationIdFor(input, options),
    sent_at_ms: sentAtFor(input, options),
    payload,
  }), payload, input);
  return assertMeetingPlatformRuntimeEvent(event, options);
}

export function assertMeetingPlatformRuntimeEvent(event = {}, options = {}) {
  const action = normalizeMeetingPlatformRuntimeEventAction(firstNonEmpty(event.action, options.action));
  if (PLATFORM_REQUIRED_ACTIONS.has(action) && !event.platform) {
    throw new MeetingTimelineSdkError('Meeting platform runtime event requires platform', {
      action,
      required_field: 'platform',
    });
  }
  return {
    ...event,
    action,
  };
}

export function buildMeetingPlatformObserveRuntimeEvent(platform, snapshot = {}, options = {}) {
  return buildMeetingPlatformRuntimeEvent({
    action: 'observe_meeting_app',
    platform,
    snapshot,
    payload: snapshot,
  }, options);
}

export function buildMeetingPlatformProviderRuntimeEvent(platform, payload = {}, options = {}) {
  return buildMeetingPlatformRuntimeEvent({
    action: 'provider_event',
    platform,
    provider_event: payload,
    payload,
  }, options);
}

export function buildMeetingPlatformAnnotationRuntimeEvent(platform, annotationInput = {}, options = {}) {
  const annotation = firstNonEmpty(annotationInput.annotation, annotationInput.mark, annotationInput);
  return buildMeetingPlatformRuntimeEvent({
    ...annotationInput,
    action: 'insert_annotation',
    platform,
    annotation,
    payload: annotationInput,
  }, options);
}

export function buildMeetingPlatformSpeakerTrackRuntimeEvent(platform, input = {}, options = {}) {
  return buildMeetingPlatformRuntimeEvent({
    ...input,
    action: 'speaker_track',
    platform,
    payload: input,
  }, options);
}

export function buildMeetingPlatformParticipantTrackRuntimeEvent(platform, input = {}, options = {}) {
  return buildMeetingPlatformRuntimeEvent({
    ...input,
    action: 'participant_track',
    platform,
    payload: input,
  }, options);
}

export function buildMeetingPlatformTimelineViewRuntimeEvent(platform, input = {}, options = {}) {
  return buildMeetingPlatformRuntimeEvent({
    ...input,
    action: 'timeline_view',
    platform,
    payload: input,
  }, options);
}

export function meetingPlatformRuntimeEventEndpoint(options = {}) {
  const path = firstNonEmpty(options.path, options.endpoint, options.runtimeEventEndpoint, options.runtime_event_endpoint, MEETING_PLATFORM_RUNTIME_EVENT_ENDPOINT);
  const baseUrl = firstNonEmpty(options.baseUrl, options.base_url);
  if (!baseUrl) return path;
  return new URL(path, baseUrl).toString();
}

async function parseResponse(response) {
  const text = await response.text();
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

export function createMeetingPlatformRuntimeEventClient(options = {}) {
  const fetchImpl = options.fetch ?? options.fetchImpl ?? globalThis.fetch;
  if (typeof fetchImpl !== 'function') {
    throw new MeetingTimelineSdkError('fetch is required for meeting platform runtime event client');
  }
  const endpoint = meetingPlatformRuntimeEventEndpoint(options);
  const defaultHeaders = {
    'content-type': 'application/json',
    ...(options.headers ?? {}),
  };

  async function send(eventInput = {}, sendOptions = {}) {
    const event = eventInput?.schema === MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA
      ? assertMeetingPlatformRuntimeEvent(eventInput, sendOptions)
      : buildMeetingPlatformRuntimeEvent(eventInput, {
        ...options,
        ...sendOptions,
      });
    const response = await fetchImpl(sendOptions.endpoint ?? endpoint, {
      method: 'POST',
      ...(sendOptions.fetchOptions ?? sendOptions.fetch_options ?? {}),
      headers: {
        ...defaultHeaders,
        ...(sendOptions.headers ?? {}),
      },
      body: JSON.stringify(event),
      signal: sendOptions.signal,
    });
    const body = await parseResponse(response);
    if (!response.ok) {
      throw new MeetingTimelineApiError('Meeting platform runtime event request failed', {
        status: response.status,
        body,
        event,
      });
    }
    return body;
  }

  return {
    type: 'meeting_platform_runtime_event_client',
    schema: MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA,
    schema_version: MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA_VERSION,
    endpoint,
    send,
    observeMeetingApp(platform, snapshot = {}, observeOptions = {}) {
      return send(buildMeetingPlatformObserveRuntimeEvent(platform, snapshot, observeOptions), observeOptions);
    },
    ingestProvider(platform, payload = {}, ingestOptions = {}) {
      return send(buildMeetingPlatformProviderRuntimeEvent(platform, payload, ingestOptions), ingestOptions);
    },
    insertAnnotation(platform, annotationInput = {}, markOptions = {}) {
      return send(buildMeetingPlatformAnnotationRuntimeEvent(platform, annotationInput, markOptions), markOptions);
    },
    insertMark(platform, annotationInput = {}, markOptions = {}) {
      return this.insertAnnotation(platform, annotationInput, markOptions);
    },
    speakerTrack(platform, input = {}, trackOptions = {}) {
      return send(buildMeetingPlatformSpeakerTrackRuntimeEvent(platform, input, trackOptions), trackOptions);
    },
    participantTrack(platform, input = {}, trackOptions = {}) {
      return send(buildMeetingPlatformParticipantTrackRuntimeEvent(platform, input, trackOptions), trackOptions);
    },
    timelineView(platform, input = {}, viewOptions = {}) {
      return send(buildMeetingPlatformTimelineViewRuntimeEvent(platform, input, viewOptions), viewOptions);
    },
    runtimeBundles(bundleOptions = {}) {
      return send({ action: 'runtime_bundles' }, bundleOptions);
    },
    registry(registryOptions = {}) {
      return send({ action: 'registry' }, registryOptions);
    },
    manifest(manifestOptions = {}) {
      return send({ action: 'manifest' }, manifestOptions);
    },
    readiness(readinessOptions = {}) {
      return send({ action: 'readiness' }, readinessOptions);
    },
    handoffReadiness(readinessOptions = {}) {
      return send({ action: 'handoff_readiness' }, readinessOptions);
    },
  };
}
