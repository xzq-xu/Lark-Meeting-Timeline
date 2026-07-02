import { compactObject, normalizeAbsoluteMs } from '../index.mjs';

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

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function decodePubSubData(data) {
  if (typeof data !== 'string' || !data) return null;
  const text = Buffer.from(data, 'base64').toString('utf8');
  if (!text.trim()) return null;
  try {
    return JSON.parse(text);
  } catch {
    return { data: text };
  }
}

export function unwrapGooglePubSubEvent(raw = {}) {
  if (!isPlainObject(raw) || !isPlainObject(raw.message) || typeof raw.message.data !== 'string') return raw;
  const decoded = decodePubSubData(raw.message.data);
  if (!decoded) return raw;
  const event = isPlainObject(decoded) ? { ...decoded } : { data: decoded };
  if (!event.id) event.id = raw.message.messageId ?? raw.message.message_id;
  if (!event.pubsub) {
    event.pubsub = {
      subscription: raw.subscription,
      message_id: raw.message.messageId ?? raw.message.message_id,
      publish_time: raw.message.publishTime ?? raw.message.publish_time,
      attributes: raw.message.attributes,
    };
  }
  return compactObject(event);
}

function resourceId(name, resourceName) {
  const text = String(name || '');
  const match = text.match(new RegExp(`${resourceName}/([^/]+)`));
  return match?.[1];
}

function lastResourceId(name) {
  const parts = String(name || '').split('/').filter(Boolean);
  return parts.at(-1);
}

function eventTypeOf(raw = {}) {
  return String(firstPath(raw, [
    'type',
    'event_type',
    'eventType',
    'header.event_type',
    'data.type',
    'data.event_type',
    'data.eventType',
  ]) || '');
}

function dataOf(raw = {}) {
  return raw.data ?? raw.payload ?? raw.event ?? raw;
}

function resourceNameOf(data = {}) {
  return firstPath(data, [
    'conferenceRecord.name',
    'conference_record.name',
    'conferenceRecordName',
    'participantSession.name',
    'participant_session.name',
    'recording.name',
    'transcript.name',
    'smartNote.name',
    'smart_note.name',
  ]);
}

function conferenceRecordNameOf(data = {}) {
  const explicit = firstPath(data, [
    'conferenceRecord.name',
    'conference_record.name',
    'conferenceRecordName',
  ]);
  if (explicit) return explicit;
  const nested = resourceNameOf(data);
  const recordId = resourceId(nested, 'conferenceRecords');
  return recordId ? `conferenceRecords/${recordId}` : undefined;
}

function eventTimestampMs(raw = {}, options = {}) {
  const input = firstPath(raw, [
    'time',
    'event_time',
    'eventTime',
    'timestamp',
    'data.time',
    'data.event_time',
    'data.eventTime',
  ]) ?? options.receivedAtMs ?? Date.now();
  return normalizeAbsoluteMs(input, 'google_meet_event_time');
}

function meetingIdentity(raw = {}, data = {}) {
  const recordName = conferenceRecordNameOf(data);
  const recordId = resourceId(recordName, 'conferenceRecords')
    ?? firstPath(raw, ['meeting_id', 'meetingId', 'conferenceRecordId'])
    ?? firstPath(data, ['meeting_id', 'meetingId', 'conferenceRecordId']);
  if (!recordId) return null;
  return compactObject({
    platform: 'google_meet',
    meeting_id: String(recordId),
    external_meeting_id: String(recordId),
    meeting_url: firstPath(data, ['meetingUri', 'meeting_url', 'meetingUrl', 'space.meetingUri']),
    title: firstPath(data, ['title', 'topic', 'space.config.title']),
    organizer_id: firstPath(data, ['organizer.user', 'organizer.email', 'space.config.accessType']),
  });
}

function participantSignal(raw, data, signalType, options) {
  const sessionName = firstPath(data, ['participantSession.name', 'participant_session.name']);
  const meeting = meetingIdentity(raw, data);
  if (!meeting) return [];
  return [compactObject({
    type: signalType,
    meeting,
    occurred_at_ms: eventTimestampMs(raw, options),
    source_event_id: firstNonEmpty(raw.id, raw.event_id, raw.eventId),
    source: 'webhook',
    participant_id: resourceId(sessionName, 'participants')
      ?? firstPath(data, ['participantSession.participant', 'participant.id', 'participant_id', 'participantId']),
    participant_name: firstPath(data, [
      'participantSession.participant.displayName',
      'participant.displayName',
      'participant.name',
      'participant_name',
      'participantName',
    ]),
    raw,
  })];
}

function artifactSignal(raw, data, artifactKind, options) {
  const meeting = meetingIdentity(raw, data);
  const artifactName = resourceNameOf(data);
  if (!meeting) return [];
  return [compactObject({
    type: 'artifact_ready',
    meeting,
    occurred_at_ms: eventTimestampMs(raw, options),
    source_event_id: firstNonEmpty(raw.id, raw.event_id, raw.eventId),
    source: 'webhook',
    artifact_kind: artifactKind,
    artifact_id: lastResourceId(artifactName),
    artifact_url: firstPath(data, ['recording.driveDestination.exportUri', 'recording.uri', 'transcript.docsDestination.document', 'smartNote.docsDestination.document']),
    raw,
  })];
}

export function normalizeGoogleMeetEvent(raw = {}, options = {}) {
  if (Array.isArray(raw)) {
    return raw.flatMap((item) => normalizeGoogleMeetEvent(item, options));
  }
  const event = unwrapGooglePubSubEvent(raw);
  const type = eventTypeOf(event);
  const lowerType = type.toLowerCase();
  const data = dataOf(event);
  const meeting = meetingIdentity(event, data);
  const base = meeting ? {
    meeting,
    occurred_at_ms: eventTimestampMs(event, options),
    source_event_id: firstNonEmpty(event.id, event.event_id, event.eventId),
    source: 'webhook',
    raw: event,
  } : null;

  if (lowerType.endsWith('conference.v2.started')) return base ? [{ ...base, type: 'meeting_started' }] : [];
  if (lowerType.endsWith('conference.v2.ended')) return base ? [{ ...base, type: 'meeting_ended' }] : [];
  if (lowerType.endsWith('participant.v2.joined')) return participantSignal(event, data, 'participant_joined', options);
  if (lowerType.endsWith('participant.v2.left')) return participantSignal(event, data, 'participant_left', options);
  if (lowerType.endsWith('transcript.v2.filegenerated')) return artifactSignal(event, data, 'transcript', options);
  if (lowerType.endsWith('recording.v2.filegenerated')) return artifactSignal(event, data, 'recording', options);
  if (lowerType.endsWith('smartnote.v2.filegenerated')) return artifactSignal(event, data, 'smart_notes', options);
  return [];
}
