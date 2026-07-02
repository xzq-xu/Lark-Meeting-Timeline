import { compactObject, normalizeAbsoluteMs } from '../index.mjs';

export const LOCAL_DETECTOR_EVENT_TYPES = Object.freeze([
  'meeting_started',
  'meeting_ended',
  'participant_joined',
  'participant_left',
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

function normalizeEventType(value) {
  const text = String(value || '')
    .trim()
    .replace(/([a-z0-9])([A-Z])/g, '$1_$2')
    .replace(/[\s.-]+/g, '_')
    .toLowerCase();
  if (['meeting_started', 'meeting_start', 'start', 'started', 'session_started', 'session_start', 'open', 'opened', 'detected'].includes(text)) {
    return 'meeting_started';
  }
  if (['meeting_ended', 'meeting_end', 'end', 'ended', 'session_ended', 'session_end', 'close', 'closed', 'left_meeting'].includes(text)) {
    return 'meeting_ended';
  }
  if (['participant_joined', 'participant_join', 'join', 'joined'].includes(text)) return 'participant_joined';
  if (['participant_left', 'participant_leave', 'leave', 'left'].includes(text)) return 'participant_left';
  return text;
}

function eventTypeOf(raw = {}) {
  return normalizeEventType(firstPath(raw, [
    'signal_type',
    'signalType',
    'type',
    'event_type',
    'eventType',
    'event',
    'action',
    'kind',
    'state',
    'status',
  ]));
}

function eventTimestampMs(raw = {}, type = eventTypeOf(raw), options = {}) {
  const input = firstNonEmpty(
    type === 'meeting_started' ? firstPath(raw, ['meeting.start_time_ms', 'meeting.startTimeMs', 'meeting.start_time', 'meeting.startTime']) : undefined,
    type === 'meeting_started' ? firstPath(raw, ['start_time_ms', 'startTimeMs', 'start_time', 'startTime', 'meeting_start_time_ms', 'meetingStartTimeMs']) : undefined,
    type === 'meeting_ended' ? firstPath(raw, ['meeting.end_time_ms', 'meeting.endTimeMs', 'meeting.end_time', 'meeting.endTime']) : undefined,
    type === 'meeting_ended' ? firstPath(raw, ['end_time_ms', 'endTimeMs', 'end_time', 'endTime', 'meeting_end_time_ms', 'meetingEndTimeMs']) : undefined,
    firstPath(raw, [
      'occurred_at_ms',
      'occurredAtMs',
      'occurred_at',
      'occurredAt',
      'detected_at_ms',
      'detectedAtMs',
      'detected_at',
      'detectedAt',
      'captured_at_ms',
      'capturedAtMs',
      'timestamp_ms',
      'timestampMs',
      'timestamp',
      'time',
      'ts',
    ]),
    options.receivedAtMs,
    Date.now(),
  );
  return normalizeAbsoluteMs(input, 'local_detector_event_time');
}

function stableIdFromUrl(url) {
  if (!url) return null;
  try {
    const parsed = new URL(String(url));
    return `${parsed.hostname}${parsed.pathname}`.replace(/[^a-zA-Z0-9._-]+/g, '-').replace(/^-+|-+$/g, '');
  } catch {
    return String(url).replace(/[^a-zA-Z0-9._-]+/g, '-').replace(/^-+|-+$/g, '') || null;
  }
}

function meetingIdentity(raw = {}) {
  const meetingUrl = firstPath(raw, [
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
    'window.url',
    'browser.url',
  ]);
  const platform = firstPath(raw, [
    'meeting.platform',
    'meeting_platform',
    'meetingPlatform',
    'detected_platform',
    'detectedPlatform',
    'target_platform',
    'targetPlatform',
    'provider_platform',
    'providerPlatform',
    'app_platform',
    'appPlatform',
    'platform',
    'provider',
    'app',
  ]) ?? 'local_detector';
  const meetingId = firstPath(raw, [
    'meeting.meeting_id',
    'meeting.meetingId',
    'meeting.id',
    'meeting.external_meeting_id',
    'meeting.externalMeetingId',
    'meeting.meeting_no',
    'meeting.meetingNo',
    'meeting_id',
    'meetingId',
    'external_meeting_id',
    'externalMeetingId',
    'meeting_no',
    'meetingNo',
    'id',
  ]) ?? stableIdFromUrl(meetingUrl);
  if (!meetingId) return null;
  return compactObject({
    platform: String(platform),
    meeting_id: String(meetingId),
    external_meeting_id: firstPath(raw, [
      'meeting.external_meeting_id',
      'meeting.externalMeetingId',
      'meeting.meeting_no',
      'meeting.meetingNo',
      'external_meeting_id',
      'externalMeetingId',
      'meeting_no',
      'meetingNo',
    ]),
    meeting_url: meetingUrl,
    title: firstPath(raw, ['meeting.title', 'meeting.topic', 'meeting.name', 'title', 'topic', 'name', 'window.title']),
    organizer_id: firstPath(raw, ['meeting.organizer_id', 'meeting.organizerId', 'organizer_id', 'organizerId']),
    organizer_name: firstPath(raw, ['meeting.organizer_name', 'meeting.organizerName', 'organizer_name', 'organizerName']),
  });
}

function participantId(raw = {}) {
  return firstPath(raw, [
    'participant.id',
    'participant.participant_id',
    'participant.participantId',
    'participant.user_id',
    'participant.userId',
    'participant_id',
    'participantId',
    'user_id',
    'userId',
  ]);
}

function participantName(raw = {}) {
  return firstPath(raw, [
    'participant.name',
    'participant.display_name',
    'participant.displayName',
    'participant.user_name',
    'participant.userName',
    'participant_name',
    'participantName',
    'user_name',
    'userName',
  ]);
}

function normalizeOne(raw = {}, options = {}) {
  const type = eventTypeOf(raw);
  const meeting = meetingIdentity(raw);
  if (!meeting) return [];
  const base = compactObject({
    meeting,
    occurred_at_ms: eventTimestampMs(raw, type, options),
    source_event_id: firstNonEmpty(raw.source_event_id, raw.sourceEventId, raw.event_id, raw.eventId, raw.id),
    source: 'local_detector',
    raw,
  });
  if (type === 'meeting_started' || type === 'meeting_ended') return [{ ...base, type }];
  if (type === 'participant_joined' || type === 'participant_left') {
    return [compactObject({
      ...base,
      type,
      participant_id: participantId(raw),
      participant_name: participantName(raw),
    })];
  }
  return [];
}

export function normalizeLocalDetectorEvent(raw = {}, options = {}) {
  if (Array.isArray(raw)) {
    return raw.flatMap((item) => normalizeLocalDetectorEvent(item, options));
  }
  return normalizeOne(raw, options);
}
