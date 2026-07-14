import {
  MeetingTimelineSdkError,
  compactObject,
  normalizeAbsoluteMs,
} from './internal-utils.mjs';

export const MEETING_SIGNAL_TYPES = Object.freeze([
  'meeting_started',
  'meeting_ended',
  'participant_joined',
  'participant_left',
  'speaker_started',
  'speaker_ended',
  'artifact_ready',
  'subscription_lifecycle',
]);

const meetingSignalTypes = new Set(MEETING_SIGNAL_TYPES);

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function requireFunction(target, name) {
  if (typeof target?.[name] !== 'function') {
    throw new MeetingTimelineSdkError(`Meeting timeline client is missing ${name}()`, { method: name });
  }
}

function normalizeSource(value) {
  return String(value || 'adapter').trim() || 'adapter';
}

function normalizePlatform(value) {
  return String(value || 'unknown').trim() || 'unknown';
}

function normalizeLifecycleType(value) {
  const text = String(value || 'unknown').trim() || 'unknown';
  return text
    .replace(/([a-z0-9])([A-Z])/g, '$1_$2')
    .replace(/[\s.-]+/g, '_')
    .toLowerCase();
}

export function normalizeMeetingIdentity(input = {}, defaults = {}) {
  const platform = normalizePlatform(firstNonEmpty(input.platform, defaults.platform));
  const meetingId = firstNonEmpty(
    input.meeting_id,
    input.meetingId,
    input.id,
    input.external_meeting_id,
    input.externalMeetingId,
    defaults.meeting_id,
    defaults.meetingId,
  );
  if (!meetingId) {
    throw new MeetingTimelineSdkError('meeting.meeting_id is required for meeting platform signals', {
      fieldName: 'meeting.meeting_id',
      input,
    });
  }
  return compactObject({
    platform,
    meeting_id: String(meetingId),
    external_meeting_id: firstNonEmpty(input.external_meeting_id, input.externalMeetingId, input.externalId),
    meeting_url: firstNonEmpty(input.meeting_url, input.meetingUrl, input.url, input.join_url, input.joinUrl),
    minute_token: firstNonEmpty(input.minute_token, input.minuteToken),
    title: firstNonEmpty(input.title, input.topic, input.name),
    organizer_id: firstNonEmpty(input.organizer_id, input.organizerId, input.host_id, input.hostId),
    organizer_name: firstNonEmpty(input.organizer_name, input.organizerName, input.host_name, input.hostName),
  });
}

export function normalizeMeetingSignal(input = {}, defaults = {}) {
  const type = String(firstNonEmpty(input.type, input.signal_type, input.signalType, defaults.type, '')).trim();
  if (!meetingSignalTypes.has(type)) {
    throw new MeetingTimelineSdkError(`Unsupported meeting signal type: ${type || '(empty)'}`, {
      type,
      supportedTypes: MEETING_SIGNAL_TYPES,
    });
  }
  const occurredAtInput = firstNonEmpty(
    input.occurred_at_ms,
    input.occurredAtMs,
    input.occurred_at,
    input.occurredAt,
    input.timestamp_ms,
    input.timestampMs,
    input.timestamp,
    input.time,
    input.ts,
    defaults.occurred_at_ms,
    defaults.occurredAtMs,
  );
  if (occurredAtInput == null) {
    throw new MeetingTimelineSdkError('occurred_at_ms is required for meeting platform signals', {
      fieldName: 'occurred_at_ms',
      input,
    });
  }
  if (type === 'subscription_lifecycle') {
    const meetingInput = input.meeting ?? defaults.meeting;
    const meeting = meetingInput ? normalizeMeetingIdentity(meetingInput, defaults.meeting ?? defaults) : undefined;
    const expiresAtInput = firstNonEmpty(
      input.expires_at_ms,
      input.expiresAtMs,
      input.expires_at,
      input.expiresAt,
      input.expire_time,
      input.expireTime,
      input.subscription_expiration_date_time,
      input.subscriptionExpirationDateTime,
    );
    return compactObject({
      type,
      meeting,
      platform: normalizePlatform(firstNonEmpty(input.platform, defaults.platform, meeting?.platform)),
      occurred_at_ms: normalizeAbsoluteMs(occurredAtInput, 'occurred_at_ms'),
      source_event_id: firstNonEmpty(input.source_event_id, input.sourceEventId, input.event_id, input.eventId, input.id),
      source: normalizeSource(firstNonEmpty(input.source, defaults.source)),
      lifecycle_type: normalizeLifecycleType(firstNonEmpty(input.lifecycle_type, input.lifecycleType, input.lifecycleEvent, input.event)),
      subscription_id: firstNonEmpty(input.subscription_id, input.subscriptionId, input.id),
      subscription_name: firstNonEmpty(input.subscription_name, input.subscriptionName, input.name),
      expires_at_ms: expiresAtInput == null ? undefined : normalizeAbsoluteMs(expiresAtInput, 'expires_at_ms'),
      resource: firstNonEmpty(input.resource, input.target_resource, input.targetResource),
      tenant_id: firstNonEmpty(input.tenant_id, input.tenantId),
      client_state: firstNonEmpty(input.client_state, input.clientState),
      raw: input.raw,
    });
  }
  const meeting = normalizeMeetingIdentity(input.meeting ?? input, defaults.meeting ?? defaults);
  return compactObject({
    type,
    meeting,
    occurred_at_ms: normalizeAbsoluteMs(occurredAtInput, 'occurred_at_ms'),
    source_event_id: firstNonEmpty(input.source_event_id, input.sourceEventId, input.event_id, input.eventId, input.id),
    source: normalizeSource(firstNonEmpty(input.source, defaults.source)),
    participant_id: firstNonEmpty(input.participant_id, input.participantId),
    participant_name: firstNonEmpty(input.participant_name, input.participantName),
    speaker_id: firstNonEmpty(input.speaker_id, input.speakerId, input.participant_id, input.participantId),
    speaker_name: firstNonEmpty(input.speaker_name, input.speakerName, input.participant_name, input.participantName),
    artifact_kind: firstNonEmpty(input.artifact_kind, input.artifactKind),
    artifact_id: firstNonEmpty(input.artifact_id, input.artifactId),
    artifact_url: firstNonEmpty(input.artifact_url, input.artifactUrl, input.url),
    raw: input.raw,
  });
}

function detectorSourceFor(signal) {
  const source = signal.source || 'adapter';
  return `${signal.meeting.platform}_${source}`;
}

function signalToStartPayload(signal) {
  return compactObject({
    platform: signal.meeting.platform,
    meeting_id: signal.meeting.meeting_id,
    external_meeting_id: signal.meeting.external_meeting_id,
    meeting_url: signal.meeting.meeting_url,
    minute_token: signal.meeting.minute_token,
    title: signal.meeting.title,
    start_time_ms: signal.occurred_at_ms,
    detector_source: detectorSourceFor(signal),
    note: signal.source_event_id ? `source_event_id=${signal.source_event_id}` : undefined,
  });
}

function signalToEndPayload(signal) {
  return compactObject({
    meeting_id: signal.meeting.meeting_id,
    end_time_ms: signal.occurred_at_ms,
    detector_source: detectorSourceFor(signal),
  });
}

function signalToParticipantMark(signal) {
  const action = signal.type === 'participant_left' ? 'left' : 'joined';
  const name = signal.participant_name || signal.participant_id || 'participant';
  return compactObject({
    id: signal.source_event_id,
    source: `${signal.meeting.platform}_participant`,
    captured_at_ms: signal.occurred_at_ms,
    kind: signal.type,
    label: `${name} ${action}`,
    intent: 'participant_track',
    payload: {
      meeting: signal.meeting,
      participant_id: signal.participant_id,
      participant_name: signal.participant_name,
      source_event_id: signal.source_event_id,
      raw: signal.raw,
    },
  });
}

function signalToSpeakerMark(signal) {
  const action = signal.type === 'speaker_ended' ? 'stopped speaking' : 'speaking';
  const name = signal.speaker_name || signal.speaker_id || signal.participant_name || signal.participant_id || 'speaker';
  return compactObject({
    id: signal.source_event_id,
    source: `${signal.meeting.platform}_speaker`,
    captured_at_ms: signal.occurred_at_ms,
    kind: signal.type,
    label: `${name} ${action}`,
    intent: 'speaker_track',
    payload: {
      meeting: signal.meeting,
      speaker_id: signal.speaker_id,
      speaker_name: signal.speaker_name,
      participant_id: signal.participant_id,
      participant_name: signal.participant_name,
      source_event_id: signal.source_event_id,
      raw: signal.raw,
    },
  });
}

export async function applyMeetingSignal(client, signalInput, options = {}) {
  const signal = normalizeMeetingSignal(signalInput, options.defaults);
  if (signal.type === 'meeting_started') {
    requireFunction(client, 'startMeeting');
    const response = await client.startMeeting(signalToStartPayload(signal));
    return { applied: true, action: 'startMeeting', signal, response };
  }
  if (signal.type === 'meeting_ended') {
    requireFunction(client, 'endMeeting');
    const response = await client.endMeeting(signalToEndPayload(signal));
    return { applied: true, action: 'endMeeting', signal, response };
  }
  if (signal.type === 'participant_joined' || signal.type === 'participant_left') {
    if (typeof options.onParticipantSignal === 'function') {
      const response = await options.onParticipantSignal(signal, client);
      return { applied: true, action: 'onParticipantSignal', signal, response };
    }
    if (options.participantAsAnnotation === true || options.participantAsMark === true) {
      requireFunction(client, 'insertMark');
      const response = await client.insertMark(signalToParticipantMark(signal), { requireCapturedAt: true });
      return { applied: true, action: 'insertParticipantMark', signal, response };
    }
    return { applied: false, action: 'skipParticipantSignal', reason: 'participant_track_not_configured', signal };
  }
  if (signal.type === 'speaker_started' || signal.type === 'speaker_ended') {
    if (typeof options.onSpeakerSignal === 'function') {
      const response = await options.onSpeakerSignal(signal, client);
      return { applied: true, action: 'onSpeakerSignal', signal, response };
    }
    if (options.speakerAsAnnotation === true || options.speakerAsMark === true) {
      requireFunction(client, 'insertMark');
      const response = await client.insertMark(signalToSpeakerMark(signal), { requireCapturedAt: true });
      return { applied: true, action: 'insertSpeakerMark', signal, response };
    }
    return { applied: false, action: 'skipSpeakerSignal', reason: 'speaker_track_not_configured', signal };
  }
  if (signal.type === 'artifact_ready') {
    if (typeof options.onArtifactSignal === 'function') {
      const response = await options.onArtifactSignal(signal, client);
      return { applied: true, action: 'onArtifactSignal', signal, response };
    }
    return { applied: false, action: 'skipArtifactSignal', reason: 'artifact_import_not_configured', signal };
  }
  if (signal.type === 'subscription_lifecycle') {
    if (typeof options.onSubscriptionLifecycleSignal === 'function') {
      const response = await options.onSubscriptionLifecycleSignal(signal, client);
      return { applied: true, action: 'onSubscriptionLifecycleSignal', signal, response };
    }
    return {
      applied: false,
      action: 'skipSubscriptionLifecycleSignal',
      reason: 'subscription_lifecycle_not_configured',
      signal,
    };
  }
  return { applied: false, action: 'skipUnknownSignal', reason: 'unsupported_signal_type', signal };
}

export async function applyMeetingSignals(client, signals = [], options = {}) {
  const rows = Array.isArray(signals) ? signals : [signals];
  const results = [];
  for (const signal of rows) {
    results.push(await applyMeetingSignal(client, signal, options));
  }
  return results;
}
