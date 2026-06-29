import { demoEvents, demoMeeting, demoSequence, demoTranscript } from './demoData.mjs';

const SECOND = 1000;
const MINUTE = 60 * SECOND;

export function toIso(value) {
  if (!value) return null;
  const date = value instanceof Date ? value : new Date(value);
  return Number.isNaN(date.getTime()) ? null : date.toISOString();
}

export function parseAbsoluteMs(value) {
  if (value == null || value === '') return null;
  if (typeof value === 'number') {
    if (value > 10_000_000_000_000) return Math.round(value / 1000);
    if (value > 10_000_000_000) return Math.round(value);
    if (value > 1_000_000_000) return Math.round(value * 1000);
    return null;
  }
  const asNumber = Number(value);
  if (Number.isFinite(asNumber)) return parseAbsoluteMs(asNumber);
  const parsed = Date.parse(value);
  return Number.isNaN(parsed) ? null : parsed;
}

export function parseOffsetMs(value) {
  if (value == null || value === '') return null;
  if (typeof value === 'number') {
    if (value > 10_000_000_000) return null;
    return Math.round(value);
  }
  const text = String(value).trim();
  if (/^\d+(\.\d+)?$/.test(text)) return Math.round(Number(text));
  const match = text.match(/^(?:(\d{1,2}):)?(\d{1,2}):(\d{1,2})(?:[.,](\d{1,3}))?$/);
  if (!match) return null;
  const hours = Number(match[1] ?? 0);
  const minutes = Number(match[2]);
  const seconds = Number(match[3]);
  const millis = Number((match[4] ?? '').padEnd(3, '0') || 0);
  return hours * 60 * MINUTE + minutes * MINUTE + seconds * SECOND + millis;
}

export function parseRelativeOrAbsoluteMs(value, meetingStartMs) {
  const asNumber = Number(value);
  const likelyUnixTimestamp = Number.isFinite(asNumber) && asNumber > 1_000_000_000;
  if (likelyUnixTimestamp) {
    const absolute = parseAbsoluteMs(value);
    if (absolute != null && meetingStartMs != null) return Math.max(0, absolute - meetingStartMs);
  }
  const offset = parseOffsetMs(value);
  if (offset != null) return offset;
  const absolute = parseAbsoluteMs(value);
  if (absolute != null && meetingStartMs != null) return Math.max(0, absolute - meetingStartMs);
  return null;
}

const annotationAbsoluteFields = [
  'captured_at_ms',
  'captured_at',
  'ink_end_at_ms',
  'ink_end_time_ms',
  'stroke_end_at_ms',
  'stroke_end_time_ms',
  'timestamp_ms',
  'timestamp',
  'ts',
  'created_at_ms',
  'device_time_ms',
];

const annotationRelativeFields = [
  'time_ms',
  'offset_ms',
  'relative_ms',
  'meeting_time_ms',
];

function firstField(input = {}, fields = []) {
  const roots = [
    { node: input, prefix: '' },
    { node: input.timing, prefix: 'timing.' },
    { node: input.payload, prefix: 'payload.' },
    { node: input.payload?.timing, prefix: 'payload.timing.' },
  ];
  for (const { node, prefix } of roots) {
    if (!node || typeof node !== 'object') continue;
    for (const field of fields) {
      if (node[field] != null && node[field] !== '') {
        return { value: node[field], field: `${prefix}${field}` };
      }
    }
  }
  return null;
}

function annotationCapturedAt(input = {}) {
  return annotationCapturedAtCandidate(input)?.value;
}

function annotationRelativeAt(input = {}) {
  return firstField(input, annotationRelativeFields)?.value;
}

function collectAbsoluteTimesFromStrokes(input = {}) {
  const roots = [
    input.strokes,
    input.stroke_points,
    input.ink?.strokes,
    input.ink?.points,
    input.payload?.strokes,
    input.payload?.stroke_points,
    input.payload?.ink?.strokes,
    input.payload?.ink?.points,
  ];
  const values = [];
  const visit = (node, depth = 0) => {
    if (node == null || depth > 8) return;
    if (Array.isArray(node)) {
      for (const child of node) visit(child, depth + 1);
      return;
    }
    if (typeof node !== 'object') return;
    for (const field of [
      'captured_at_ms',
      'captured_at',
      'timestamp_ms',
      'timestamp',
      'ts',
      'time',
      'time_ms',
      't',
      'created_at_ms',
      'end_at_ms',
      'end_time_ms',
    ]) {
      const parsed = parseAbsoluteMs(node[field]);
      if (parsed != null) values.push(parsed);
    }
    for (const [key, child] of Object.entries(node)) {
      if (['x', 'y', 'pressure', 'p', 'width', 'height'].includes(key)) continue;
      visit(child, depth + 1);
    }
  };
  for (const root of roots) visit(root);
  return values.length ? Math.max(...values) : null;
}

function annotationCapturedAtCandidate(input = {}) {
  const direct = firstField(input, annotationAbsoluteFields);
  if (direct) return { ...direct, source: 'captured_at' };
  const strokeEnd = collectAbsoluteTimesFromStrokes(input);
  if (strokeEnd != null) {
    return { value: strokeEnd, field: 'strokes[*].time', source: 'stroke_point_time' };
  }
  return null;
}

export function annotationCapturedAbsoluteMs(input = {}, fallbackNow = Date.now()) {
  const captured = annotationCapturedAtCandidate(input);
  return parseAbsoluteMs(captured?.value) ?? fallbackNow;
}

export function annotationCapturedAbsoluteMsStrict(input = {}) {
  const captured = annotationCapturedAtCandidate(input);
  return parseAbsoluteMs(captured?.value);
}

function annotationTiming(input = {}, meeting = {}, fallbackNow = Date.now()) {
  const meetingStartMs = parseAbsoluteMs(meeting.start_time);
  const meetingEndMs = parseAbsoluteMs(meeting.end_time);
  const explicitRelative = firstField(input, annotationRelativeFields);
  const captured = annotationCapturedAtCandidate(input);
  const explicitTimeMs = parseRelativeOrAbsoluteMs(explicitRelative?.value, meetingStartMs);
  const capturedTimeMs = parseRelativeOrAbsoluteMs(captured?.value, meetingStartMs);
  const fallbackRelative = meetingStartMs != null ? Math.max(0, fallbackNow - meetingStartMs) : 0;
  const timeMs = explicitTimeMs ?? capturedTimeMs ?? fallbackRelative;
  const capturedAbsoluteMs = parseAbsoluteMs(captured?.value);
  const afterMeetingEndMs = meetingEndMs != null ? timeMs - Math.max(0, meetingEndMs - (meetingStartMs ?? meetingEndMs)) : null;
  const source = explicitTimeMs != null
    ? 'explicit_time'
    : capturedTimeMs != null
      ? (captured?.source ?? 'captured_at')
      : 'server_received_at';
  return {
    time_ms: timeMs,
    source,
    source_field: explicitTimeMs != null ? explicitRelative.field : captured?.field ?? null,
    relative_input: explicitRelative?.value ?? null,
    captured_at: captured?.value ?? null,
    captured_at_ms: capturedAbsoluteMs,
    server_received_at_ms: fallbackNow,
    server_receive_delay_ms: capturedAbsoluteMs != null ? fallbackNow - capturedAbsoluteMs : null,
    after_meeting_end_ms: afterMeetingEndMs != null && afterMeetingEndMs > 0 ? afterMeetingEndMs : 0,
  };
}

function toAbsoluteIso(value) {
  const absolute = parseAbsoluteMs(value);
  return absolute == null ? toIso(value) : toIso(absolute);
}

function firstDefined(...values) {
  return values.find((value) => value != null && value !== '');
}

function getPath(raw, path) {
  const parts = path.split('.');
  let node = raw;
  for (const part of parts) node = node?.[part];
  return node;
}

function firstPath(raw, paths) {
  return firstDefined(...paths.map((path) => getPath(raw, path)));
}

function larkEventEnvelope(payload = {}) {
  return payload?.event ?? payload?.data ?? payload;
}

function larkEventPayload(payload = {}) {
  const envelope = larkEventEnvelope(payload);
  return envelope?.data?.event
    ?? envelope?.event
    ?? payload?.data?.event
    ?? payload?.data?.payload?.event
    ?? envelope;
}

function larkEventHeader(payload = {}) {
  const envelope = larkEventEnvelope(payload);
  return payload?.header
    ?? payload?.data?.header
    ?? envelope?.header
    ?? envelope?.data?.header
    ?? payload?.raw_ws?.header
    ?? payload?.raw_ws?.data?.header
    ?? {};
}

function larkRawEventType(payload = {}, event = larkEventPayload(payload), header = larkEventHeader(payload)) {
  const envelope = larkEventEnvelope(payload);
  return firstDefined(
    header.event_type,
    payload?.event_type,
    payload?.type,
    payload?.data?.event_type,
    payload?.data?.type,
    envelope?.event_type,
    envelope?.type,
    envelope?.data?.event_type,
    envelope?.data?.type,
    envelope?.data?.header?.event_type,
    event?.event_type,
    event?.type,
  );
}

function textOf(item) {
  const direct = firstDefined(item.text, item.content, item.sentence, item.transcript, item.words);
  if (Array.isArray(direct)) return direct.map((x) => typeof x === 'string' ? x : textOf(x)).join('');
  if (direct && typeof direct === 'object') return textOf(direct);
  return String(direct ?? '').trim();
}

function speakerOf(item) {
  const speaker = firstDefined(item.speaker, item.user, item.participant, item.owner, item.creator);
  if (!speaker || typeof speaker !== 'object') {
    return {
      speaker_id: String(firstDefined(item.speaker_id, item.user_id, item.participant_id, '') || ''),
      speaker_name: String(firstDefined(item.speaker_name, item.user_name, item.name, speaker, '未知说话人')),
    };
  }
  return {
    speaker_id: String(firstDefined(speaker.id, speaker.user_id, speaker.open_id, speaker.union_id, '') || ''),
    speaker_name: String(firstDefined(speaker.name, speaker.display_name, speaker.user_name, speaker.en_name, '未知说话人')),
  };
}

function candidateArray(raw, keys) {
  if (Array.isArray(raw)) return raw;
  for (const key of keys) {
    const parts = key.split('.');
    let node = raw;
    for (const part of parts) node = node?.[part];
    if (Array.isArray(node)) return node;
  }
  return [];
}

export function normalizeTranscript(rawTranscript, meeting = {}) {
  const meetingStartMs = parseAbsoluteMs(meeting.start_time);
  const rows = candidateArray(rawTranscript, [
    'data.transcript',
    'data.segments',
    'data.sentences',
    'data.items',
    'transcript',
    'segments',
    'sentences',
    'items',
  ]);

  return rows.map((item, index) => {
    if (item.source === 'lark_minute' && item.start_ms != null && item.end_ms != null && item.text) {
      return {
        id: String(firstDefined(item.id, `seg-${index + 1}`)),
        start_ms: Math.round(item.start_ms),
        end_ms: Math.max(Math.round(item.start_ms) + 1, Math.round(item.end_ms)),
        speaker_id: String(item.speaker_id ?? ''),
        speaker_name: String(item.speaker_name ?? '未知说话人'),
        text: String(item.text),
        language: item.language ?? null,
        source: 'lark_minute',
        raw: item.raw ?? null,
      };
    }
    const startValue = firstDefined(item.start_ms, item.start_time_ms, item.startTimeMs, item.start, item.start_time, item.startTime, item.ts);
    const endValue = firstDefined(item.end_ms, item.end_time_ms, item.endTimeMs, item.end, item.end_time, item.endTime);
    const start_ms = parseRelativeOrAbsoluteMs(startValue, meetingStartMs) ?? 0;
    const end_ms = parseRelativeOrAbsoluteMs(endValue, meetingStartMs) ?? Math.max(start_ms + 1, start_ms + 12_000);
    const speaker = speakerOf(item);
    return {
      id: String(firstDefined(item.id, item.segment_id, item.sentence_id, `seg-${index + 1}`)),
      start_ms,
      end_ms: Math.max(start_ms + 1, end_ms),
      speaker_id: speaker.speaker_id,
      speaker_name: speaker.speaker_name,
      text: textOf(item),
      language: firstDefined(item.language, item.language_code, item.languageCode, null),
      source: 'lark_minute',
      raw: item,
    };
  }).filter((seg) => seg.text).sort((a, b) => a.start_ms - b.start_ms || a.end_ms - b.end_ms);
}

function eventTypeOf(rawType) {
  const text = String(rawType || '').toLowerCase();
  if (text.includes('screen') && text.includes('end')) return 'screen_share_end';
  if (text.includes('screen') && (text.includes('start') || text.includes('begin'))) return 'screen_share_start';
  if (text.includes('record') && (text.includes('complete') || text.includes('generated') || text.includes('finish'))) return 'recording_completed';
  if (text.includes('record') && text.includes('end')) return 'recording_end';
  if (text.includes('record') && (text.includes('start') || text.includes('begin'))) return 'recording_start';
  if (text.includes('join')) return 'participant_join';
  if (text.includes('leave')) return 'participant_leave';
  if (text.includes('minute') && text.includes('generated')) return 'minute_generated';
  if (text.includes('end')) return 'meeting_end';
  if (text.includes('start') || text.includes('begin')) return 'meeting_start';
  return text || 'meeting_event';
}

function labelForEvent(type) {
  const labels = {
    meeting_start: '会议开始',
    meeting_end: '会议结束',
    participant_join: '参会人加入',
    participant_leave: '参会人离开',
    recording_start: '录制开始',
    recording_end: '录制结束',
    recording_completed: '录制完成',
    screen_share_start: '开始共享屏幕',
    screen_share_end: '结束共享屏幕',
    minute_generated: '妙记生成',
  };
  return labels[type] ?? type;
}

export function normalizeLarkEventPayload(payload, meeting = {}) {
  const event = larkEventPayload(payload);
  const header = larkEventHeader(payload);
  const rawType = larkRawEventType(payload, event, header);
  const type = eventTypeOf(rawType);
  const meetingStartMs = parseAbsoluteMs(meeting.start_time);
  const meetingInfo = event?.meeting ?? event?.meeting_info ?? {};
  const eventTimeValue = type === 'meeting_end'
    ? firstDefined(
      meetingInfo.end_time,
      meetingInfo.end_at,
      event?.end_time,
      event?.end_at,
      payload?.end_time,
      header.create_time,
      payload?.event_ts,
      payload?.ts,
      event?.event_ts,
      event?.time,
    )
    : type === 'meeting_start'
      ? firstDefined(
        meetingInfo.start_time,
        meetingInfo.start_at,
        meetingInfo.begin_time,
        event?.start_time,
        event?.start_at,
        event?.begin_time,
        payload?.start_time,
        header.create_time,
        payload?.event_ts,
        payload?.ts,
        event?.event_ts,
        event?.time,
      )
      : firstDefined(header.create_time, payload?.event_ts, payload?.ts, event?.event_ts, event?.time, event?.start_time, event?.end_time);
  const absolute = parseAbsoluteMs(eventTimeValue);
  const explicitOffset = parseOffsetMs(firstDefined(event?.time_ms, event?.offset_ms, event?.relative_time_ms));
  const time_ms = explicitOffset ?? (absolute != null && meetingStartMs != null ? Math.max(0, absolute - meetingStartMs) : 0);
  return {
    id: String(firstDefined(header.event_id, payload?.uuid, event?.id, `evt-${type}-${time_ms}`)),
    time_ms,
    type,
    label: labelForEvent(type),
    source: 'lark_event',
    metadata: {
      raw_type: rawType,
      meeting_id: firstDefined(
        event?.meeting_id,
        event?.meeting?.meeting_id,
        event?.meeting?.id,
        event?.meeting_info?.meeting_id,
        event?.meeting_info?.id,
        event?.meeting_info?.open_meeting_id,
        event?.meeting_info?.meeting_no,
        payload?.meeting_id,
        meeting.meeting_id,
      ),
      minute_token: firstDefined(event?.minute_token, event?.minute?.token, payload?.minute_token, meeting.minute_token),
    },
    raw: payload,
  };
}

export function extractLarkMeetingPatch(payload, currentMeeting = {}) {
  const event = larkEventPayload(payload);
  const header = larkEventHeader(payload);
  const rawType = larkRawEventType(payload, event, header);
  const type = eventTypeOf(rawType);
  const explicitStartCandidate = firstPath({ payload, event, header }, [
    'event.meeting.start_time',
    'event.meeting.start_at',
    'event.meeting.begin_time',
    'event.meeting_info.start_time',
    'event.meeting_info.start_at',
    'event.meeting_info.begin_time',
    'event.start_time',
    'event.start_at',
    'event.begin_time',
    'payload.start_time',
  ]);
  const startCandidate = explicitStartCandidate ?? (type === 'meeting_start' || type === 'participant_join'
    ? firstPath({ payload, event, header }, [
      'header.create_time',
      'payload.event_ts',
      'payload.ts',
      'event.event_ts',
      'event.time',
    ])
    : null);
  const endCandidate = firstPath({ payload, event, header }, [
    'event.meeting.end_time',
    'event.meeting.end_at',
    'event.meeting_info.end_time',
    'event.meeting_info.end_at',
    'event.end_time',
    'event.end_at',
    'payload.end_time',
    'header.create_time',
  ]);
  const startTime = toAbsoluteIso(startCandidate);
  const endTime = toAbsoluteIso(endCandidate);
  const startTimeSource = explicitStartCandidate != null
    ? 'lark_payload_start_time'
    : startCandidate != null
      ? 'event_create_time'
      : currentMeeting.start_time
        ? currentMeeting.start_time_source ?? 'current_meeting'
        : 'server_now_fallback';
  const endTimeSource = type === 'meeting_end'
    ? endCandidate != null
      ? 'lark_payload_end_time'
      : 'server_now_fallback'
    : currentMeeting.end_time
      ? currentMeeting.end_time_source ?? 'current_meeting'
      : null;
  const patch = {
    platform: 'lark',
    meeting_id: String(firstPath({ payload, event }, [
      'event.meeting.meeting_id',
      'event.meeting.id',
      'event.meeting_info.meeting_id',
      'event.meeting_info.id',
      'event.meeting_info.open_meeting_id',
      'event.meeting_info.meeting_no',
      'event.meeting_id',
      'event.meeting_no',
      'event.vc_meeting_id',
      'event.id',
      'payload.meeting_id',
    ]) ?? currentMeeting.meeting_id ?? 'lark-live-meeting'),
    external_meeting_id: firstPath({ payload, event }, [
      'event.open_meeting_id',
      'event.meeting.open_meeting_id',
      'event.meeting.meeting_no',
      'event.meeting_info.open_meeting_id',
      'event.meeting_info.meeting_no',
      'event.meeting_no',
      'event.vc_meeting_id',
    ]) ?? currentMeeting.external_meeting_id ?? null,
    meeting_url: firstPath({ payload, event }, [
      'event.meeting.meeting_url',
      'event.meeting.url',
      'event.meeting.join_url',
      'event.meeting.share_url',
      'event.meeting_info.meeting_url',
      'event.meeting_info.url',
      'event.meeting_info.join_url',
      'event.meeting_info.share_url',
      'event.meeting_url',
      'event.url',
      'event.join_url',
      'event.share_url',
      'payload.meeting_url',
    ]) ?? currentMeeting.meeting_url ?? null,
    minute_token: firstPath({ payload, event }, [
      'event.minute.token',
      'event.minute.minute_token',
      'event.minute_token',
      'event.minutes_token',
      'payload.minute_token',
    ]) ?? currentMeeting.minute_token ?? null,
    title: String(firstPath({ payload, event }, [
      'event.meeting.topic',
      'event.meeting.title',
      'event.meeting.name',
      'event.meeting_info.topic',
      'event.meeting_info.title',
      'event.meeting_info.name',
      'event.topic',
      'event.title',
      'event.summary',
      'payload.title',
    ]) ?? currentMeeting.title ?? '飞书实时会议'),
    start_time: startTime ?? currentMeeting.start_time ?? toIso(new Date()),
    start_time_source: startTimeSource,
    start_time_reliable: startTimeSource !== 'server_now_fallback',
    end_time: type === 'meeting_start'
      ? null
      : type === 'meeting_end'
        ? (endTime ?? toIso(new Date()))
        : currentMeeting.end_time ?? null,
    end_time_source: endTimeSource,
    timezone: currentMeeting.timezone ?? 'Asia/Shanghai',
  };
  return patch;
}

export function normalizeEvents(rawEvents, meeting = {}) {
  const rows = candidateArray(rawEvents, ['data.events', 'events', 'items']);
  return rows.map((event, index) => {
    if (event.type && event.time_ms != null && event.label) return event;
    const normalized = normalizeLarkEventPayload(event, meeting);
    return { ...normalized, id: normalized.id || `evt-${index + 1}` };
  }).sort((a, b) => a.time_ms - b.time_ms);
}

export function normalizeSequence(rawSequence, meeting = {}) {
  const rows = candidateArray(rawSequence, ['data.sequence', 'sequence', 'items', 'marks', 'events']);
  const meetingStartMs = parseAbsoluteMs(meeting.start_time);
  return rows.map((item, index) => {
    if (item.id && item.time_ms != null && item.kind && item.label) {
      return {
        id: String(item.id),
        time_ms: Math.round(item.time_ms),
        kind: String(item.kind),
        label: String(item.label),
        source: item.source ? String(item.source) : undefined,
        time_source: item.time_source ?? item.payload?.timing?.source ?? undefined,
        payload: item.payload ?? {},
        raw: item.raw ?? null,
      };
    }
    const timeValue = firstDefined(
      item.time_ms,
      item.offset_ms,
      item.relative_ms,
      item.meeting_time_ms,
      item.created_at_ms,
      item.device_time_ms,
      item.timestamp,
      item.ts,
      item.time,
    );
    const time_ms = parseRelativeOrAbsoluteMs(timeValue, meetingStartMs) ?? 0;
    return {
      id: String(firstDefined(item.id, item.mark_id, item.event_id, `seq-${index + 1}`)),
      time_ms,
      kind: String(firstDefined(item.kind, item.type, item.action, 'external')),
      label: String(firstDefined(item.label, item.text, item.title, item.intent, item.kind, `序列 ${index + 1}`)),
      source: item.source ? String(item.source) : undefined,
      time_source: item.time_source ?? item.payload?.timing?.source ?? undefined,
      payload: item.payload ?? item,
      raw: item,
    };
  }).sort((a, b) => a.time_ms - b.time_ms);
}

export function normalizeAnnotationEvent(input, meeting = {}, fallbackNow = Date.now()) {
  const timing = annotationTiming(input, meeting, fallbackNow);
  const time_ms = timing.time_ms;
  const textCandidates = input.text_candidates ?? input.candidates ?? input.ocr_candidates ?? input.payload?.text_candidates;
  const primaryText = firstDefined(
    input.text,
    input.reading,
    Array.isArray(textCandidates) ? textCandidates[0] : null,
    input.payload?.text,
    input.payload?.reading,
  );
  const kind = String(firstDefined(input.kind, input.type, input.intent, input.mark?.action, input.action, 'annotation'));
  const label = String(firstDefined(input.label, primaryText, input.intent, input.mark?.target_text, kind, '实时标注'));
  return {
    id: String(firstDefined(input.id, input.annotation_id, input.mark_id, input.event_id, `ann-${fallbackNow}`)),
    time_ms,
    kind,
    label,
    source: String(firstDefined(input.source, input.device?.type, input.device_id, 'external_annotation')),
    time_source: timing.source,
    payload: {
      text: primaryText ?? null,
      text_candidates: Array.isArray(textCandidates) ? textCandidates : undefined,
      intent: input.intent ?? input.payload?.intent ?? null,
      mark: input.mark ?? input.payload?.mark ?? undefined,
      target: input.target ?? input.target_region ?? input.payload?.target ?? undefined,
      strokes: input.strokes ?? input.stroke_points ?? input.payload?.strokes ?? undefined,
      device: input.device ?? (input.device_id ? { id: input.device_id } : undefined),
      timing: {
        ...timing,
        normalized_time_ms: time_ms,
      },
      raw_payload: input.payload ?? undefined,
    },
    raw: input,
  };
}

function distanceToSegment(timeMs, segment) {
  if (timeMs >= segment.start_ms && timeMs <= segment.end_ms) return 0;
  return Math.min(Math.abs(timeMs - segment.start_ms), Math.abs(timeMs - segment.end_ms));
}

export function alignSequence(timeline, options = {}) {
  const windowMs = options.window_ms ?? 90_000;
  const eventWindowMs = options.event_window_ms ?? 120_000;
  return (timeline.sequence ?? []).map((item) => {
    const nearbySegments = (timeline.segments ?? [])
      .map((segment) => ({ segment, distance_ms: distanceToSegment(item.time_ms, segment) }))
      .filter((x) => x.distance_ms <= windowMs)
      .sort((a, b) => a.distance_ms - b.distance_ms || a.segment.start_ms - b.segment.start_ms)
      .slice(0, 4);
    const activeSegment = nearbySegments.find((x) => x.distance_ms === 0)?.segment ?? nearbySegments[0]?.segment ?? null;
    const nearbyEvents = (timeline.events ?? [])
      .map((event) => ({ event, distance_ms: Math.abs(item.time_ms - event.time_ms) }))
      .filter((x) => x.distance_ms <= eventWindowMs)
      .sort((a, b) => a.distance_ms - b.distance_ms || a.event.time_ms - b.event.time_ms)
      .slice(0, 5);
    const nearestDistance = nearbySegments[0]?.distance_ms ?? windowMs;
    const confidence = activeSegment
      ? 0.95
      : Math.max(0.2, Number((1 - nearestDistance / windowMs).toFixed(2)));
    return {
      sequence_id: item.id,
      time_ms: item.time_ms,
      label: item.label,
      kind: item.kind,
      active_segment_id: activeSegment?.id ?? null,
      confidence,
      segments: nearbySegments,
      events: nearbyEvents,
    };
  });
}

export function buildTimeline({ meeting = {}, segments = [], events = [], sequence = [] } = {}) {
  const normalizedMeeting = {
    platform: 'lark',
    meeting_id: meeting.meeting_id ?? meeting.id ?? 'unknown-meeting',
    external_meeting_id: meeting.external_meeting_id ?? meeting.lark_meeting_id ?? null,
    meeting_url: meeting.meeting_url ?? meeting.url ?? null,
    minute_token: meeting.minute_token ?? null,
    title: meeting.title ?? meeting.topic ?? 'Lark meeting',
    start_time: toIso(meeting.start_time) ?? toIso(new Date()),
    end_time: toIso(meeting.end_time) ?? null,
    timezone: meeting.timezone ?? 'Asia/Shanghai',
    pending_binding: Boolean(meeting.pending_binding),
    source: meeting.source ?? null,
    start_time_source: meeting.start_time_source ?? null,
    start_time_reliable: meeting.start_time_reliable ?? null,
    end_time_source: meeting.end_time_source ?? null,
    reserve_id: meeting.reserve_id ?? null,
    reserve_meeting_no: meeting.reserve_meeting_no ?? null,
    app_link: meeting.app_link ?? null,
    live_link: meeting.live_link ?? null,
  };
  const timeline = {
    meeting: normalizedMeeting,
    segments: normalizeTranscript(segments, normalizedMeeting),
    events: normalizeEvents(events, normalizedMeeting),
    sequence: normalizeSequence(sequence, normalizedMeeting),
    duration_ms: 0,
    alignments: [],
    updated_at: new Date().toISOString(),
  };
  const maxEnd = Math.max(
    10 * MINUTE,
    ...timeline.segments.map((x) => x.end_ms),
    ...timeline.events.map((x) => x.time_ms + 30_000),
    ...timeline.sequence.map((x) => x.time_ms + 30_000),
  );
  const explicitEnd = parseAbsoluteMs(normalizedMeeting.end_time);
  const start = parseAbsoluteMs(normalizedMeeting.start_time);
  timeline.duration_ms = explicitEnd != null && start != null ? Math.max(maxEnd, explicitEnd - start) : maxEnd;
  timeline.alignments = alignSequence(timeline);
  return timeline;
}

export function demoTimeline() {
  return buildTimeline({
    meeting: demoMeeting,
    segments: demoTranscript,
    events: demoEvents,
    sequence: demoSequence,
  });
}

export function mergeTimeline(current, patch) {
  const meeting = { ...(current?.meeting ?? demoMeeting), ...(patch.meeting ?? {}) };
  return buildTimeline({
    meeting,
    segments: patch.segments ?? current?.segments ?? [],
    events: patch.events ?? current?.events ?? [],
    sequence: patch.sequence ?? current?.sequence ?? [],
  });
}
