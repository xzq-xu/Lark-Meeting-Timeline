import { compactObject, normalizeAbsoluteMs } from './internal-utils.mjs';

export const LARK_MEETING_EVENT_TYPES = Object.freeze([
  'vc.meeting.all_meeting_started_v1',
  'vc.meeting.all_meeting_ended_v1',
  'vc.meeting.meeting_started_v1',
  'vc.meeting.meeting_ended_v1',
  'vc.meeting.join_meeting_v1',
  'vc.meeting.leave_meeting_v1',
]);

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

function eventEnvelope(raw = {}) {
  return raw?.event ?? raw?.data ?? raw;
}

function eventPayload(raw = {}) {
  const envelope = eventEnvelope(raw);
  return envelope?.data?.event
    ?? envelope?.event
    ?? raw?.data?.event
    ?? raw?.data?.payload?.event
    ?? envelope;
}

function eventHeader(raw = {}) {
  const envelope = eventEnvelope(raw);
  return raw?.header
    ?? raw?.data?.header
    ?? envelope?.header
    ?? envelope?.data?.header
    ?? raw?.raw_ws?.header
    ?? raw?.raw_ws?.data?.header
    ?? {};
}

function eventTypeOf(raw = {}, event = eventPayload(raw), header = eventHeader(raw)) {
  const envelope = eventEnvelope(raw);
  return String(firstNonEmpty(
    header.event_type,
    raw?.event_type,
    raw?.type,
    raw?.data?.event_type,
    raw?.data?.type,
    envelope?.event_type,
    envelope?.type,
    envelope?.data?.event_type,
    envelope?.data?.type,
    envelope?.data?.header?.event_type,
    event?.event_type,
    event?.type,
  ) || '');
}

function meetingObject(event = {}) {
  return event.meeting ?? event.meeting_info ?? event.meetingInfo ?? {};
}

function eventTimestampMs(raw = {}, event = eventPayload(raw), type = eventTypeOf(raw), options = {}) {
  const header = eventHeader(raw);
  const meeting = meetingObject(event);
  const lowerType = String(type || '').toLowerCase();
  const input = firstNonEmpty(
    lowerType.includes('ended') || lowerType.includes('end') ? meeting.end_time : undefined,
    lowerType.includes('ended') || lowerType.includes('end') ? meeting.end_at : undefined,
    lowerType.includes('ended') || lowerType.includes('end') ? event.end_time : undefined,
    lowerType.includes('ended') || lowerType.includes('end') ? event.end_at : undefined,
    lowerType.includes('started') || lowerType.includes('start') ? meeting.start_time : undefined,
    lowerType.includes('started') || lowerType.includes('start') ? meeting.start_at : undefined,
    lowerType.includes('started') || lowerType.includes('start') ? meeting.begin_time : undefined,
    lowerType.includes('started') || lowerType.includes('start') ? event.start_time : undefined,
    lowerType.includes('started') || lowerType.includes('start') ? event.start_at : undefined,
    lowerType.includes('started') || lowerType.includes('start') ? event.begin_time : undefined,
    header.create_time,
    raw.event_ts,
    raw.ts,
    event.event_ts,
    event.time,
    event.timestamp,
    options.receivedAtMs,
    Date.now(),
  );
  return normalizeAbsoluteMs(input, 'lark_event_time');
}

function minuteToken(raw = {}, event = eventPayload(raw)) {
  return firstPath({ raw, event }, [
    'event.minute.token',
    'event.minute.minute_token',
    'event.minute_token',
    'event.minutes_token',
    'raw.minute_token',
    'raw.minutes_token',
  ]);
}

function meetingIdentity(raw = {}, event = eventPayload(raw)) {
  const meeting = meetingObject(event);
  const meetingId = firstPath({ raw, event, meeting }, [
    'meeting.meeting_id',
    'meeting.id',
    'meeting.open_meeting_id',
    'meeting.meeting_no',
    'event.meeting_id',
    'event.open_meeting_id',
    'event.meeting_no',
    'event.vc_meeting_id',
    'event.id',
    'raw.meeting_id',
  ]);
  if (!meetingId) return null;
  return compactObject({
    platform: 'lark',
    meeting_id: String(meetingId),
    external_meeting_id: firstPath({ event, meeting }, [
      'event.open_meeting_id',
      'event.meeting_no',
      'event.vc_meeting_id',
      'meeting.open_meeting_id',
      'meeting.meeting_no',
      'meeting.vc_meeting_id',
    ]),
    meeting_url: firstPath({ raw, event, meeting }, [
      'meeting.meeting_url',
      'meeting.url',
      'meeting.join_url',
      'meeting.share_url',
      'event.meeting_url',
      'event.url',
      'event.join_url',
      'event.share_url',
      'raw.meeting_url',
    ]),
    minute_token: minuteToken(raw, event),
    title: firstPath({ raw, event, meeting }, [
      'meeting.topic',
      'meeting.title',
      'meeting.name',
      'event.topic',
      'event.title',
      'event.summary',
      'raw.title',
    ]),
    organizer_id: firstPath({ event, meeting }, [
      'meeting.host_user_id',
      'meeting.host_id',
      'meeting.owner_id',
      'event.host_user_id',
      'event.host_id',
      'event.owner_id',
    ]),
    organizer_name: firstPath({ event, meeting }, [
      'meeting.host_name',
      'meeting.owner_name',
      'event.host_name',
      'event.owner_name',
    ]),
  });
}

function participantId(event = {}) {
  return firstPath({ event }, [
    'event.participant.id',
    'event.participant.user_id',
    'event.participant.open_id',
    'event.participant.union_id',
    'event.user.id',
    'event.user.user_id',
    'event.user.open_id',
    'event.user.union_id',
    'event.user_id',
    'event.open_id',
    'event.union_id',
    'event.operator_id',
  ]);
}

function participantName(event = {}) {
  return firstPath({ event }, [
    'event.participant.name',
    'event.participant.display_name',
    'event.participant.user_name',
    'event.user.name',
    'event.user.display_name',
    'event.user.user_name',
    'event.user_name',
    'event.operator_name',
    'event.name',
  ]);
}

function artifactUrl(event = {}) {
  return firstPath({ event }, [
    'event.minute.url',
    'event.minute.link',
    'event.minute.docs_url',
    'event.minute.docsUrl',
    'event.minute_url',
    'event.minutes_url',
    'event.recording.url',
    'event.recording.download_url',
    'event.recording.file_url',
  ]);
}

function normalizeOne(raw = {}, options = {}) {
  const event = eventPayload(raw);
  const eventType = eventTypeOf(raw, event);
  const lowerType = eventType.toLowerCase();
  const meeting = meetingIdentity(raw, event);
  if (!meeting) return [];
  const base = compactObject({
    meeting,
    occurred_at_ms: eventTimestampMs(raw, event, eventType, options),
    source_event_id: firstNonEmpty(eventHeader(raw).event_id, raw.event_id, raw.uuid, event.event_id, event.id),
    source: 'webhook',
    raw,
  });

  if (lowerType.includes('all_meeting_started') || lowerType.includes('meeting_started')) {
    return [{ ...base, type: 'meeting_started' }];
  }
  if (lowerType.includes('all_meeting_ended') || lowerType.includes('meeting_ended')) {
    return [{ ...base, type: 'meeting_ended' }];
  }
  if (lowerType.includes('join_meeting') || lowerType.includes('participant_join') || lowerType.includes('joined')) {
    return [compactObject({
      ...base,
      type: 'participant_joined',
      participant_id: participantId(event),
      participant_name: participantName(event),
    })];
  }
  if (lowerType.includes('leave_meeting') || lowerType.includes('participant_leave') || lowerType.includes('left')) {
    return [compactObject({
      ...base,
      type: 'participant_left',
      participant_id: participantId(event),
      participant_name: participantName(event),
    })];
  }
  if (lowerType.includes('minute') || lowerType.includes('transcript')) {
    return [compactObject({
      ...base,
      type: 'artifact_ready',
      artifact_kind: 'transcript',
      artifact_id: minuteToken(raw, event),
      artifact_url: artifactUrl(event),
    })];
  }
  if (lowerType.includes('record')) {
    return [compactObject({
      ...base,
      type: 'artifact_ready',
      artifact_kind: 'recording',
      artifact_id: firstPath({ event }, ['event.recording.id', 'event.recording.file_token', 'event.recording_token']),
      artifact_url: artifactUrl(event),
    })];
  }
  return [];
}

export function normalizeLarkEvent(raw = {}, options = {}) {
  if (Array.isArray(raw)) {
    return raw.flatMap((item) => normalizeLarkEvent(item, options));
  }
  return normalizeOne(raw, options);
}
