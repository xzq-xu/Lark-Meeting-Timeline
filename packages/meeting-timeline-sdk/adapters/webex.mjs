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

function webexData(raw = {}) {
  return raw.data ?? raw.payload?.data ?? raw.payload ?? raw.resourceData ?? raw;
}

function resourceOf(raw = {}) {
  return String(firstNonEmpty(raw.resource, raw.webhook?.resource, raw.payload?.resource, '')).toLowerCase();
}

function eventOf(raw = {}) {
  return String(firstNonEmpty(raw.event, raw.webhook?.event, raw.payload?.event, raw.type, '')).toLowerCase();
}

function eventTimestampMs(raw = {}, data = webexData(raw), resource = '', event = '', options = {}) {
  const input = firstNonEmpty(
    resource === 'meetings' && event === 'started' ? firstNonEmpty(data.startTime, data.start_time, data.started, data.start) : undefined,
    resource === 'meetings' && event === 'ended' ? firstNonEmpty(data.endTime, data.end_time, data.ended, data.end) : undefined,
    resource === 'meetingparticipants' && event === 'joined' ? firstNonEmpty(data.joinTime, data.join_time, data.joined, data.created) : undefined,
    resource === 'meetingparticipants' && event === 'left' ? firstNonEmpty(data.leaveTime, data.leave_time, data.left, data.updated) : undefined,
    data.created,
    data.createdAt,
    data.created_at,
    data.updated,
    data.updatedAt,
    data.updated_at,
    raw.created,
    raw.createdAt,
    raw.created_at,
    raw.timestamp,
    raw.time,
    options.receivedAtMs,
    Date.now(),
  );
  return normalizeAbsoluteMs(input, 'webex_event_time');
}

function meetingIdentity(raw = {}, data = webexData(raw), resource = resourceOf(raw)) {
  const meetingId = firstNonEmpty(
    data.meetingId,
    data.meeting_id,
    data.meeting?.id,
    data.meetingInstanceId,
    data.meeting_instance_id,
    data.scheduledMeetingId,
    data.scheduled_meeting_id,
    data.meetingSeriesId,
    data.meeting_series_id,
    resource === 'meetings' ? data.id : undefined,
    raw.meetingId,
    raw.meeting_id,
  );
  if (!meetingId) return null;
  return compactObject({
    platform: 'webex',
    meeting_id: String(meetingId),
    external_meeting_id: String(firstNonEmpty(data.meetingNumber, data.meeting_number, data.sipAddress, data.sip_address, meetingId)),
    meeting_url: firstNonEmpty(data.webLink, data.web_link, data.meetingLink, data.meeting_link, data.joinUrl, data.join_url),
    title: firstNonEmpty(data.title, data.topic, data.subject, data.meeting?.title),
    organizer_id: firstNonEmpty(data.hostUserId, data.host_user_id, data.hostPersonId, data.host_person_id, data.host?.id),
    organizer_name: firstNonEmpty(data.hostDisplayName, data.host_display_name, data.hostName, data.host_name, data.host?.displayName),
  });
}

function artifactUrl(data = {}) {
  return firstPath(data, [
    'txtDownloadLink',
    'txt_download_link',
    'temporaryDirectDownloadLinks.recordingDownloadLink',
    'temporary_direct_download_links.recording_download_link',
    'downloadUrl',
    'download_url',
    'playbackUrl',
    'playback_url',
    'webLink',
    'web_link',
  ]);
}

function normalizeOne(raw = {}, options = {}) {
  const data = webexData(raw);
  const resource = resourceOf(raw);
  const event = eventOf(raw);
  const meeting = meetingIdentity(raw, data, resource);
  if (!meeting) return [];
  const base = compactObject({
    meeting,
    occurred_at_ms: eventTimestampMs(raw, data, resource, event, options),
    source_event_id: firstNonEmpty(raw.eventId, raw.event_id, raw.id, data.eventId, data.event_id),
    source: 'webhook',
    raw,
  });

  if (resource === 'meetings' && event === 'started') return [{ ...base, type: 'meeting_started' }];
  if (resource === 'meetings' && event === 'ended') return [{ ...base, type: 'meeting_ended' }];

  if (resource === 'meetingparticipants' && event === 'joined') {
    return [compactObject({
      ...base,
      type: 'participant_joined',
      participant_id: firstNonEmpty(data.personId, data.person_id, data.participantId, data.participant_id, data.id, data.email),
      participant_name: firstNonEmpty(data.displayName, data.display_name, data.name, data.personEmail, data.person_email, data.email),
    })];
  }
  if (resource === 'meetingparticipants' && event === 'left') {
    return [compactObject({
      ...base,
      type: 'participant_left',
      participant_id: firstNonEmpty(data.personId, data.person_id, data.participantId, data.participant_id, data.id, data.email),
      participant_name: firstNonEmpty(data.displayName, data.display_name, data.name, data.personEmail, data.person_email, data.email),
    })];
  }
  if (resource === 'meetingtranscripts' && event === 'created') {
    return [compactObject({
      ...base,
      type: 'artifact_ready',
      artifact_kind: 'transcript',
      artifact_id: firstNonEmpty(data.id, data.transcriptId, data.transcript_id),
      artifact_url: artifactUrl(data),
    })];
  }
  if ((resource === 'recordings' || resource === 'convergedrecordings') && ['created', 'updated'].includes(event)) {
    return [compactObject({
      ...base,
      type: 'artifact_ready',
      artifact_kind: 'recording',
      artifact_id: firstNonEmpty(data.recordingId, data.recording_id, data.id),
      artifact_url: artifactUrl(data),
    })];
  }
  return [];
}

export function normalizeWebexEvent(raw = {}, options = {}) {
  if (Array.isArray(raw)) {
    return raw.flatMap((item) => normalizeWebexEvent(item, options));
  }
  if (Array.isArray(raw.items)) {
    return raw.items.flatMap((item) => normalizeWebexEvent(item, options));
  }
  if (Array.isArray(raw.events)) {
    return raw.events.flatMap((item) => normalizeWebexEvent(item, options));
  }
  return normalizeOne(raw, options);
}
