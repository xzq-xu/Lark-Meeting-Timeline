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

function zoomObject(raw = {}) {
  return raw.payload?.object ?? raw.object ?? raw.payload ?? raw;
}

function eventTimestampMs(raw = {}, object = zoomObject(raw), eventName = '', options = {}) {
  const lower = eventName.toLowerCase();
  const participant = object.participant ?? {};
  const input = firstNonEmpty(
    lower === 'meeting.started' ? object.start_time : undefined,
    lower === 'meeting.ended' ? object.end_time : undefined,
    lower === 'meeting.participant_joined' ? participant.join_time : undefined,
    lower === 'meeting.participant_left' ? participant.leave_time : undefined,
    raw.event_ts,
    raw.eventTs,
    object.start_time,
    object.end_time,
    options.receivedAtMs,
    Date.now(),
  );
  return normalizeAbsoluteMs(input, 'zoom_event_time');
}

function meetingIdentity(raw = {}, object = zoomObject(raw)) {
  const meetingId = firstNonEmpty(object.uuid, object.id, raw.meeting_id, raw.meetingId);
  if (!meetingId) return null;
  return compactObject({
    platform: 'zoom',
    meeting_id: String(meetingId),
    external_meeting_id: String(firstNonEmpty(object.id, meetingId)),
    meeting_url: firstNonEmpty(object.join_url, object.joinUrl, object.start_url, object.startUrl),
    title: firstNonEmpty(object.topic, object.title, object.name),
    organizer_id: firstNonEmpty(object.host_id, object.hostId),
    organizer_name: firstNonEmpty(object.host_name, object.hostName),
  });
}

function participantOf(object = {}) {
  return object.participant ?? object.participant_info ?? {};
}

function recordingUrl(object = {}) {
  const files = Array.isArray(object.recording_files) ? object.recording_files : [];
  return firstNonEmpty(object.share_url, files[0]?.download_url, files[0]?.play_url);
}

function normalizeOne(raw = {}, options = {}) {
  const eventName = String(raw.event || raw.event_type || raw.type || '');
  const lower = eventName.toLowerCase();
  const object = zoomObject(raw);
  const meeting = meetingIdentity(raw, object);
  if (!meeting) return [];
  const base = compactObject({
    meeting,
    occurred_at_ms: eventTimestampMs(raw, object, eventName, options),
    source_event_id: firstNonEmpty(raw.event_id, raw.eventId, raw.id, raw.event_ts),
    source: 'webhook',
    raw,
  });

  if (lower === 'meeting.started') return [{ ...base, type: 'meeting_started' }];
  if (lower === 'meeting.ended') return [{ ...base, type: 'meeting_ended' }];
  if (lower === 'meeting.participant_joined') {
    const participant = participantOf(object);
    return [compactObject({
      ...base,
      type: 'participant_joined',
      participant_id: firstNonEmpty(participant.id, participant.user_id, participant.userId, participant.customer_key),
      participant_name: firstNonEmpty(participant.user_name, participant.userName, participant.name),
    })];
  }
  if (lower === 'meeting.participant_left') {
    const participant = participantOf(object);
    return [compactObject({
      ...base,
      type: 'participant_left',
      participant_id: firstNonEmpty(participant.id, participant.user_id, participant.userId, participant.customer_key),
      participant_name: firstNonEmpty(participant.user_name, participant.userName, participant.name),
    })];
  }
  if (lower === 'recording.completed') {
    return [compactObject({
      ...base,
      type: 'artifact_ready',
      artifact_kind: 'recording',
      artifact_id: firstNonEmpty(object.uuid, object.id),
      artifact_url: recordingUrl(object),
    })];
  }
  if (lower.includes('transcript') && (lower.includes('completed') || lower.includes('ready'))) {
    return [compactObject({
      ...base,
      type: 'artifact_ready',
      artifact_kind: 'transcript',
      artifact_id: firstNonEmpty(object.uuid, object.id),
      artifact_url: firstPath(object, ['transcript.download_url', 'transcript.url']),
    })];
  }
  return [];
}

export function normalizeZoomEvent(raw = {}, options = {}) {
  if (Array.isArray(raw)) {
    return raw.flatMap((item) => normalizeZoomEvent(item, options));
  }
  return normalizeOne(raw, options);
}
