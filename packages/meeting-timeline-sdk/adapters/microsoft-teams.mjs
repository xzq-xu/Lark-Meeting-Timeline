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

function dataOf(raw = {}) {
  return raw.resourceData ?? raw.payload?.resourceData ?? raw.payload ?? raw.value ?? raw;
}

function eventTypeOf(raw = {}, data = dataOf(raw)) {
  return String(firstNonEmpty(
    data.eventType,
    data.callEventType,
    data.type,
    data.name,
    raw.eventType,
    raw.callEventType,
    raw.type,
    raw.name,
  ) || '');
}

function eventTimestampMs(raw = {}, data = dataOf(raw), options = {}) {
  const input = firstNonEmpty(
    data.eventDateTime,
    data.startTime,
    data.endTime,
    data.time,
    data.timestamp,
    raw.eventDateTime,
    raw.time,
    raw.timestamp,
    options.receivedAtMs,
    Date.now(),
  );
  return normalizeAbsoluteMs(input, 'microsoft_teams_event_time');
}

function participantTimestampMs(raw, data, participant, options) {
  const input = firstNonEmpty(
    participant.eventDateTime,
    participant.joinDateTime,
    participant.leaveDateTime,
    participant.time,
    data.eventDateTime,
    raw.eventDateTime,
    options.receivedAtMs,
    Date.now(),
  );
  return normalizeAbsoluteMs(input, 'microsoft_teams_participant_time');
}

function meetingIdentity(raw = {}, data = dataOf(raw)) {
  const meetingId = firstNonEmpty(
    data.onlineMeetingId,
    data.meetingId,
    data.meeting?.id,
    data.id,
    raw.onlineMeetingId,
    raw.meetingId,
    raw.resource,
  );
  if (!meetingId) return null;
  return compactObject({
    platform: 'microsoft_teams',
    meeting_id: String(meetingId),
    external_meeting_id: String(firstNonEmpty(data.onlineMeetingId, data.meetingId, data.id, raw.resource, meetingId)),
    meeting_url: firstNonEmpty(data.joinWebUrl, data.join_url, data.meeting?.joinWebUrl, raw.joinWebUrl),
    title: firstNonEmpty(data.subject, data.title, data.meeting?.subject, raw.subject),
    organizer_id: firstPath(data, [
      'organizer.id',
      'organizer.user.id',
      'meeting.organizer.id',
      'meeting.organizer.user.id',
    ]),
    organizer_name: firstPath(data, [
      'organizer.displayName',
      'organizer.user.displayName',
      'meeting.organizer.displayName',
      'meeting.organizer.user.displayName',
    ]),
  });
}

function participantId(participant = {}) {
  return firstNonEmpty(
    participant.id,
    participant.participantId,
    participant.userId,
    participant.user?.id,
    participant.identity?.user?.id,
    participant.identity?.application?.id,
  );
}

function participantName(participant = {}) {
  return firstNonEmpty(
    participant.displayName,
    participant.name,
    participant.user?.displayName,
    participant.identity?.user?.displayName,
    participant.identity?.application?.displayName,
  );
}

function participantDeltas(raw = {}, data = dataOf(raw)) {
  const direct = data['participants@delta']
    ?? data.participantsDelta
    ?? data.participants
    ?? data.roster
    ?? raw['participants@delta']
    ?? raw.participants;
  if (Array.isArray(direct)) return direct;
  if (direct && typeof direct === 'object') return [direct];
  const single = data.participant ?? raw.participant;
  return single ? [single] : [];
}

function participantSignalType(participant = {}, fallbackType = '') {
  const removed = participant.removedState != null
    || participant['@removed'] != null
    || participant.leaveDateTime != null
    || String(participant.action || '').toLowerCase().includes('left')
    || fallbackType.toLowerCase().includes('leave')
    || fallbackType.toLowerCase().includes('left');
  return removed ? 'participant_left' : 'participant_joined';
}

function normalizeOne(raw = {}, options = {}) {
  const data = dataOf(raw);
  const eventType = eventTypeOf(raw, data);
  const lowerType = eventType.toLowerCase();
  const meeting = meetingIdentity(raw, data);
  if (!meeting) return [];
  const base = compactObject({
    meeting,
    occurred_at_ms: eventTimestampMs(raw, data, options),
    source_event_id: firstNonEmpty(raw.id, raw.subscriptionId, data.id, data.eventId),
    source: 'webhook',
    raw,
  });

  if (lowerType === 'callstarted' || lowerType.includes('meetingstart')) {
    return [{ ...base, type: 'meeting_started' }];
  }
  if (lowerType === 'callended' || lowerType.includes('meetingend')) {
    return [{ ...base, type: 'meeting_ended' }];
  }
  if (lowerType === 'rosterupdated' || lowerType.includes('participant')) {
    return participantDeltas(raw, data).map((participant) => compactObject({
      type: participantSignalType(participant, lowerType),
      meeting,
      occurred_at_ms: participantTimestampMs(raw, data, participant, options),
      source_event_id: firstNonEmpty(raw.id, raw.subscriptionId, data.id, participant.id),
      source: 'webhook',
      participant_id: participantId(participant),
      participant_name: participantName(participant),
      raw,
    }));
  }
  if (lowerType.includes('transcript') && (lowerType.includes('created') || lowerType.includes('available'))) {
    return [{ ...base, type: 'artifact_ready', artifact_kind: 'transcript', artifact_id: data.id }];
  }
  if (lowerType.includes('recording') && (lowerType.includes('created') || lowerType.includes('available'))) {
    return [{ ...base, type: 'artifact_ready', artifact_kind: 'recording', artifact_id: data.id }];
  }
  return [];
}

export function normalizeMicrosoftTeamsEvent(raw = {}, options = {}) {
  if (Array.isArray(raw)) {
    return raw.flatMap((item) => normalizeMicrosoftTeamsEvent(item, options));
  }
  if (Array.isArray(raw.value) && !raw.resourceData) {
    return raw.value.flatMap((item) => normalizeMicrosoftTeamsEvent(item, options));
  }
  return normalizeOne(raw, options);
}
