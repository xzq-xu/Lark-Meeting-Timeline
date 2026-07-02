export const SDK_VERSION = '0.1.0';

export const DEFAULT_ENDPOINTS = Object.freeze({
  state: '/api/state',
  ingestInfo: '/api/annotation-ingest-info',
  meetingSessionStatus: '/api/meeting-session/status',
  meetingSessionStart: '/api/meeting-session/start',
  meetingSessionEnd: '/api/meeting-session/end',
  annotation: '/api/annotations',
  annotationBatch: '/api/annotations/batch',
  annotationStatus: '/api/annotations/status',
  stream: '/api/stream',
});

export class MeetingTimelineSdkError extends Error {
  constructor(message, details = {}) {
    super(message);
    this.name = 'MeetingTimelineSdkError';
    this.details = details;
  }
}

export class MeetingTimelineApiError extends MeetingTimelineSdkError {
  constructor(message, details = {}) {
    super(message, details);
    this.name = 'MeetingTimelineApiError';
    this.status = details.status ?? null;
    this.body = details.body ?? null;
  }
}

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date);
}

export function compactObject(value) {
  if (Array.isArray(value)) return value.map((item) => compactObject(item));
  if (!isPlainObject(value)) return value;
  return Object.fromEntries(
    Object.entries(value)
      .filter(([, item]) => item !== undefined)
      .map(([key, item]) => [key, compactObject(item)]),
  );
}

export function normalizeAbsoluteMs(value, fieldName = 'timestamp') {
  if (value == null || value === '') return undefined;
  if (value instanceof Date) {
    const ms = value.getTime();
    if (Number.isFinite(ms)) return ms;
    throw new MeetingTimelineSdkError(`Invalid ${fieldName}: Date is not finite`, { fieldName, value });
  }
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      throw new MeetingTimelineSdkError(`Invalid ${fieldName}: number is not finite`, { fieldName, value });
    }
    if (value > 10_000_000_000_000) return Math.round(value / 1000);
    if (value > 10_000_000_000) return Math.round(value);
    if (value > 1_000_000_000) return Math.round(value * 1000);
    throw new MeetingTimelineSdkError(`Invalid ${fieldName}: expected absolute unix time`, { fieldName, value });
  }
  const numeric = Number(value);
  if (Number.isFinite(numeric)) return normalizeAbsoluteMs(numeric, fieldName);
  const parsed = Date.parse(String(value));
  if (Number.isFinite(parsed)) return parsed;
  throw new MeetingTimelineSdkError(`Invalid ${fieldName}: cannot parse absolute time`, { fieldName, value });
}

function maybeAbsoluteMs(...values) {
  const value = firstNonEmpty(...values);
  return value == null ? undefined : normalizeAbsoluteMs(value);
}

function normalizeCandidates(input = {}) {
  const candidates = firstNonEmpty(
    input.text_candidates,
    input.textCandidates,
    input.candidates,
    input.ocr_candidates,
    input.ocrCandidates,
    input.payload?.text_candidates,
  );
  if (candidates == null) return undefined;
  return Array.isArray(candidates) ? candidates.map((item) => String(item)) : [String(candidates)];
}

export function buildMeetingStartPayload(input = {}, defaults = {}) {
  const startTimeMs = maybeAbsoluteMs(
    input.start_time_ms,
    input.startTimeMs,
    input.started_at_ms,
    input.startedAtMs,
    input.start_time,
    input.startTime,
    input.started_at,
    input.startedAt,
  );
  return compactObject({
    platform: input.platform ?? defaults.platform ?? 'lark',
    meeting_id: firstNonEmpty(input.meeting_id, input.meetingId, input.id, input.session_id, input.sessionId),
    external_meeting_id: firstNonEmpty(
      input.external_meeting_id,
      input.externalMeetingId,
      input.meeting_no,
      input.meetingNo,
      input.lark_meeting_id,
      input.open_meeting_id,
    ),
    meeting_url: firstNonEmpty(input.meeting_url, input.meetingUrl, input.url, input.join_url, input.joinUrl),
    minute_token: firstNonEmpty(input.minute_token, input.minuteToken),
    title: firstNonEmpty(input.title, input.topic, input.name),
    timezone: input.timezone ?? defaults.timezone,
    start_time_ms: startTimeMs,
    end_time: input.end_time ?? input.endTime,
    detector_source: input.detector_source ?? input.detectorSource ?? input.source ?? defaults.detectorSource,
    note: input.note,
    force: input.force ?? input.forceMeetingSession ?? defaults.force,
    keep_segments: input.keep_segments ?? input.keepSegments,
    suppress_auto_annotations: input.suppress_auto_annotations ?? input.suppressAutoAnnotations,
    suppress_demo_annotations: input.suppress_demo_annotations ?? input.suppressDemoAnnotations,
  });
}

export function buildMeetingEndPayload(input = {}, defaults = {}) {
  const endTimeMs = maybeAbsoluteMs(
    input.end_time_ms,
    input.endTimeMs,
    input.ended_at_ms,
    input.endedAtMs,
    input.end_time,
    input.endTime,
    input.ended_at,
    input.endedAt,
  );
  return compactObject({
    meeting_id: firstNonEmpty(input.meeting_id, input.meetingId),
    end_time_ms: endTimeMs,
    time_ms: input.time_ms ?? input.timeMs,
    detector_source: input.detector_source ?? input.detectorSource ?? input.source ?? defaults.detectorSource,
  });
}

export function buildTimelineMark(input = {}, defaults = {}) {
  const requireCapturedAt = defaults.requireCapturedAt ?? true;
  const capturedAtInput = firstNonEmpty(
    input.captured_at_ms,
    input.capturedAtMs,
    input.captured_at,
    input.capturedAt,
    input.ink_end_at_ms,
    input.inkEndAtMs,
    input.timestamp_ms,
    input.timestampMs,
    input.timestamp,
    input.ts,
  );
  if (capturedAtInput == null && requireCapturedAt) {
    throw new MeetingTimelineSdkError('captured_at_ms is required for reliable timeline insertion', {
      fieldName: 'captured_at_ms',
    });
  }
  const capturedAtMs = capturedAtInput == null ? undefined : normalizeAbsoluteMs(capturedAtInput, 'captured_at_ms');
  const textCandidates = normalizeCandidates(input);
  const primaryText = firstNonEmpty(
    input.text,
    input.reading,
    textCandidates?.[0],
    input.payload?.text,
    input.payload?.reading,
  );
  const kind = String(firstNonEmpty(input.kind, input.type, input.intent, input.mark?.action, input.action, defaults.kind, 'annotation'));
  const label = String(firstNonEmpty(input.label, primaryText, input.intent, input.mark?.target_text, kind, '实时标注'));
  const meetingSession = input.meeting_session ?? input.meetingSession ?? defaults.meetingSession;
  return compactObject({
    id: firstNonEmpty(input.id, input.annotation_id, input.annotationId, input.mark_id, input.markId),
    source: input.source ?? defaults.source ?? input.device?.type ?? 'external_annotation',
    device_id: firstNonEmpty(input.device_id, input.deviceId, defaults.deviceId),
    captured_at_ms: capturedAtMs,
    time_ms: input.time_ms ?? input.timeMs,
    kind,
    label,
    text: primaryText,
    text_candidates: textCandidates,
    intent: input.intent ?? input.payload?.intent,
    mark: input.mark ?? input.payload?.mark,
    target: input.target ?? input.target_region ?? input.targetRegion ?? input.payload?.target,
    strokes: input.strokes ?? input.stroke_points ?? input.strokePoints ?? input.payload?.strokes,
    device: input.device ?? (defaults.deviceId ? { id: defaults.deviceId } : undefined),
    payload: input.payload,
    realtime: input.realtime ?? defaults.realtime,
    live: input.live ?? defaults.live,
    meeting_session: meetingSession ? buildMeetingStartPayload(meetingSession, defaults.meetingDefaults ?? defaults) : undefined,
    start_meeting_session: input.start_meeting_session ?? input.startMeetingSession,
    force_meeting_session: input.force_meeting_session ?? input.forceMeetingSession,
  });
}

function normalizeBaseUrl(baseUrl) {
  if (!baseUrl) throw new MeetingTimelineSdkError('baseUrl is required');
  return String(baseUrl).replace(/\/+$/, '');
}

function buildUrl(baseUrl, path, query = {}) {
  const url = new URL(path, `${normalizeBaseUrl(baseUrl)}/`);
  for (const [key, value] of Object.entries(query)) {
    if (value != null && value !== '') url.searchParams.set(key, String(value));
  }
  return url;
}

export class MeetingTimelineClient {
  constructor(options = {}) {
    this.baseUrl = normalizeBaseUrl(options.baseUrl);
    this.fetchImpl = options.fetch ?? options.fetchImpl ?? globalThis.fetch;
    if (typeof this.fetchImpl !== 'function') {
      throw new MeetingTimelineSdkError('fetch implementation is required');
    }
    this.headers = options.headers ?? {};
    this.authToken = options.authToken ?? options.token ?? null;
    this.source = options.source ?? 'external_annotation';
    this.deviceId = options.deviceId ?? options.device_id ?? null;
    this.detectorSource = options.detectorSource ?? options.detector_source ?? this.source;
    this.requireCapturedAt = options.requireCapturedAt ?? true;
    this.timeoutMs = options.timeoutMs ?? 30_000;
  }

  async request(path, options = {}) {
    const url = buildUrl(this.baseUrl, path, options.query);
    const headers = {
      accept: 'application/json',
      ...this.headers,
      ...(options.headers ?? {}),
    };
    const init = {
      method: options.method ?? 'GET',
      headers,
      signal: options.signal,
    };
    if (options.body !== undefined) {
      init.headers = { 'content-type': 'application/json; charset=utf-8', ...headers };
      init.body = JSON.stringify(options.body);
    }
    if (this.authToken && !init.headers.authorization) {
      init.headers.authorization = `Bearer ${this.authToken}`;
    }
    const response = await this.fetchImpl(String(url), init);
    const text = await response.text();
    let body = null;
    if (text) {
      try {
        body = JSON.parse(text);
      } catch {
        body = text;
      }
    }
    if (!response.ok) {
      throw new MeetingTimelineApiError(`Meeting timeline API request failed: ${response.status}`, {
        status: response.status,
        body,
        url: String(url),
      });
    }
    return body;
  }

  getState() {
    return this.request(DEFAULT_ENDPOINTS.state);
  }

  getIngestInfo() {
    return this.request(DEFAULT_ENDPOINTS.ingestInfo);
  }

  getMeetingSessionStatus() {
    return this.request(DEFAULT_ENDPOINTS.meetingSessionStatus);
  }

  startMeeting(input = {}) {
    return this.request(DEFAULT_ENDPOINTS.meetingSessionStart, {
      method: 'POST',
      body: buildMeetingStartPayload(input, {
        detectorSource: this.detectorSource,
      }),
    });
  }

  endMeeting(input = {}) {
    return this.request(DEFAULT_ENDPOINTS.meetingSessionEnd, {
      method: 'POST',
      body: buildMeetingEndPayload(input, {
        detectorSource: this.detectorSource,
      }),
    });
  }

  insertMark(input = {}, options = {}) {
    const body = options.raw === true
      ? input
      : buildTimelineMark(input, {
        source: this.source,
        deviceId: this.deviceId,
        detectorSource: this.detectorSource,
        requireCapturedAt: options.requireCapturedAt ?? this.requireCapturedAt,
        realtime: options.realtime,
        live: options.live,
        meetingSession: options.meetingSession,
      });
    return this.request(DEFAULT_ENDPOINTS.annotation, {
      method: 'POST',
      body,
    });
  }

  addAnnotation(input = {}, options = {}) {
    return this.insertMark(input, options);
  }

  insertAnnotation(input = {}, options = {}) {
    return this.insertMark(input, options);
  }

  insertMarks(inputs = [], options = {}) {
    const rows = Array.isArray(inputs) ? inputs : inputs.annotations ?? inputs.items ?? [];
    const annotations = options.raw === true
      ? rows
      : rows.map((item) => buildTimelineMark(item, {
        source: this.source,
        deviceId: this.deviceId,
        detectorSource: this.detectorSource,
        requireCapturedAt: options.requireCapturedAt ?? this.requireCapturedAt,
        realtime: options.realtime,
        live: options.live,
      }));
    return this.request(DEFAULT_ENDPOINTS.annotationBatch, {
      method: 'POST',
      body: { annotations },
    });
  }

  addAnnotations(inputs = [], options = {}) {
    return this.insertMarks(inputs, options);
  }

  getAnnotationStatus(id) {
    return this.request(DEFAULT_ENDPOINTS.annotationStatus, {
      query: { id },
    });
  }

  subscribeState(options = {}) {
    const EventSourceImpl = options.EventSource ?? options.EventSourceImpl ?? globalThis.EventSource;
    if (typeof EventSourceImpl !== 'function') {
      throw new MeetingTimelineSdkError('EventSource implementation is required for subscribeState');
    }
    const url = buildUrl(this.baseUrl, options.path ?? DEFAULT_ENDPOINTS.stream, options.query);
    const stream = new EventSourceImpl(String(url), options.eventSourceOptions);
    const onState = options.onState;
    const onMessage = (event) => {
      const data = JSON.parse(event.data);
      if (typeof onState === 'function') onState(data, event);
    };
    stream.addEventListener?.('state', onMessage);
    stream.addEventListener?.('message', onMessage);
    if (typeof options.onError === 'function') {
      stream.addEventListener?.('error', options.onError);
      stream.onerror = options.onError;
    }
    return {
      stream,
      close() {
        stream.removeEventListener?.('state', onMessage);
        stream.removeEventListener?.('message', onMessage);
        stream.close();
      },
    };
  }
}

export function createMeetingTimelineClient(options = {}) {
  return new MeetingTimelineClient(options);
}

export const createMeetingTimelineSdk = createMeetingTimelineClient;
