import {
  MeetingTimelineApiError,
  MeetingTimelineSdkError,
} from './errors.mjs';

let fallbackProducerIdSequence = 0;

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date);
}

function compactObject(value) {
  if (Array.isArray(value)) return value.map((item) => compactObject(item));
  if (!isPlainObject(value)) return value;
  return Object.fromEntries(
    Object.entries(value)
      .filter(([, item]) => item !== undefined)
      .map(([key, item]) => [key, compactObject(item)]),
  );
}

function normalizeAbsoluteMs(value, fieldName = 'timestamp') {
  if (value == null || value === '') {
    throw new MeetingTimelineSdkError(`Invalid ${fieldName}: value is required`, { fieldName, value });
  }
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

function producerInputId(input = {}) {
  return firstNonEmpty(input.id, input.annotation_id, input.annotationId, input.mark_id, input.markId);
}

function producerCapturedAt(input = {}) {
  return firstNonEmpty(
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
}

function defaultProducerAnnotationId(producerId, capturedAtMs) {
  const randomId = globalThis.crypto?.randomUUID?.();
  if (randomId) return `${producerId}:${randomId}`;
  fallbackProducerIdSequence += 1;
  return `${producerId}:${capturedAtMs.toString(36)}:${fallbackProducerIdSequence.toString(36)}:${Math.random().toString(36).slice(2, 10)}`;
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

function buildProducerMeetingSession(input = {}) {
  const startTime = firstNonEmpty(
    input.start_time_ms,
    input.startTimeMs,
    input.started_at_ms,
    input.startedAtMs,
    input.start_time,
    input.startTime,
  );
  return compactObject({
    platform: input.platform ?? 'lark',
    meeting_id: firstNonEmpty(input.meeting_id, input.meetingId, input.id, input.session_id, input.sessionId),
    external_meeting_id: firstNonEmpty(input.external_meeting_id, input.externalMeetingId, input.meeting_no, input.meetingNo),
    meeting_url: firstNonEmpty(input.meeting_url, input.meetingUrl, input.url, input.join_url, input.joinUrl),
    title: firstNonEmpty(input.title, input.topic, input.name),
    timezone: input.timezone,
    start_time_ms: startTime == null ? undefined : normalizeAbsoluteMs(startTime, 'meeting_session.start_time_ms'),
    detector_source: input.detector_source ?? input.detectorSource ?? input.source,
    force: input.force,
  });
}

function buildProducerMark(input, defaults = {}) {
  const textCandidates = normalizeCandidates(input);
  const primaryText = firstNonEmpty(
    input.text,
    input.reading,
    textCandidates?.[0],
    input.payload?.text,
    input.payload?.reading,
  );
  const kind = String(firstNonEmpty(input.kind, input.type, input.intent, input.mark?.action, input.action, 'annotation'));
  const label = String(firstNonEmpty(input.label, primaryText, input.intent, input.mark?.target_text, kind, 'annotation'));
  const meetingSession = input.meeting_session ?? input.meetingSession;
  return compactObject({
    id: String(input.id),
    source: input.source ?? defaults.source,
    device_id: firstNonEmpty(input.device_id, input.deviceId),
    captured_at_ms: input.captured_at_ms,
    time_ms: input.time_ms ?? input.timeMs,
    kind,
    label,
    text: primaryText,
    text_candidates: textCandidates,
    intent: input.intent ?? input.payload?.intent,
    mark: input.mark ?? input.payload?.mark,
    target: input.target ?? input.target_region ?? input.targetRegion ?? input.payload?.target,
    strokes: input.strokes ?? input.stroke_points ?? input.strokePoints ?? input.payload?.strokes,
    device: input.device,
    payload: input.payload,
    realtime: input.realtime ?? defaults.realtime,
    live: input.live,
    meeting_session: meetingSession ? buildProducerMeetingSession(meetingSession) : undefined,
    start_meeting_session: input.start_meeting_session ?? input.startMeetingSession,
    force_meeting_session: input.force_meeting_session ?? input.forceMeetingSession,
  });
}

function normalizeBaseUrl(baseUrl) {
  if (!baseUrl) throw new MeetingTimelineSdkError('baseUrl is required');
  return String(baseUrl).replace(/\/+$/, '');
}

class ProducerHttpClient {
  constructor(options = {}) {
    this.baseUrl = normalizeBaseUrl(options.baseUrl);
    const configuredFetch = options.fetch ?? options.fetchImpl;
    if (configuredFetch != null && typeof configuredFetch !== 'function') {
      throw new MeetingTimelineSdkError('fetch implementation is required');
    }
    if (typeof configuredFetch === 'function') {
      this.fetchImpl = (input, init) => configuredFetch(input, init);
    } else if (typeof globalThis.fetch === 'function') {
      // Some browser implementations reject Window.fetch when it is invoked
      // later as a method of another object ("Illegal invocation").
      this.fetchImpl = (input, init) => globalThis.fetch(input, init);
    } else {
      throw new MeetingTimelineSdkError('fetch implementation is required');
    }
    this.headers = options.headers ?? {};
    this.authToken = options.authToken ?? options.token ?? null;
  }

  async insertMark(annotation) {
    const headers = {
      accept: 'application/json',
      'content-type': 'application/json; charset=utf-8',
      ...this.headers,
    };
    if (this.authToken && !headers.authorization) headers.authorization = `Bearer ${this.authToken}`;
    const response = await this.fetchImpl(`${this.baseUrl}/api/annotations`, {
      method: 'POST',
      headers,
      body: JSON.stringify(annotation),
    });
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
        url: `${this.baseUrl}/api/annotations`,
      });
    }
    return body;
  }
}

function positiveInteger(value, fallback) {
  const number = Number(value);
  return Number.isInteger(number) && number > 0 ? number : fallback;
}

function nonNegativeNumber(value, fallback) {
  const number = Number(value);
  return Number.isFinite(number) && number >= 0 ? number : fallback;
}

function isRetryableProducerError(error) {
  const status = Number(error?.status ?? error?.details?.status);
  if (!Number.isFinite(status)) return true;
  return [408, 425, 429].includes(status) || status >= 500;
}

function defaultProducerSleep(delayMs) {
  return new Promise((resolve) => setTimeout(resolve, delayMs));
}

/**
 * Device-neutral source adapter for annotations. It timestamps at capture time,
 * keeps one stable id across retries, and retains failed deliveries for flush().
 */
export class MeetingTimelineAnnotationProducer {
  constructor(options = {}) {
    this.source = options.source ?? 'sdk_annotation_producer';
    this.producerId = String(options.producerId ?? options.producer_id ?? 'standard-producer');
    this.client = options.client ?? new ProducerHttpClient(options);
    if (!this.client || typeof this.client.insertMark !== 'function') {
      throw new MeetingTimelineSdkError('client.insertMark is required for annotation producer');
    }
    this.clock = options.clock ?? Date.now;
    this.idFactory = options.idFactory ?? options.id_factory ?? null;
    this.maxAttempts = positiveInteger(options.maxAttempts ?? options.max_attempts, 3);
    this.retryDelayMs = nonNegativeNumber(options.retryDelayMs ?? options.retry_delay_ms, 100);
    this.maxRetryDelayMs = nonNegativeNumber(options.maxRetryDelayMs ?? options.max_retry_delay_ms, 1_000);
    this.sleep = options.sleep ?? defaultProducerSleep;
    this.pending = new Map();
    this.sequence = 0;
    this.flushPromise = null;
  }

  capture(input = {}) {
    if (!isPlainObject(input)) {
      throw new MeetingTimelineSdkError('annotation input must be an object');
    }
    const capturedAtMs = normalizeAbsoluteMs(producerCapturedAt(input) ?? this.clock(), 'captured_at_ms');
    this.sequence += 1;
    const id = producerInputId(input) ?? (
      typeof this.idFactory === 'function'
        ? this.idFactory({ input, capturedAtMs, sequence: this.sequence, producerId: this.producerId })
        : defaultProducerAnnotationId(this.producerId, capturedAtMs)
    );
    if (id == null || id === '') {
      throw new MeetingTimelineSdkError('annotation producer idFactory returned an empty id');
    }
    return buildProducerMark({
      ...input,
      id: String(id),
      source: input.source ?? this.source,
      captured_at_ms: capturedAtMs,
    }, {
      source: this.source,
      realtime: input.realtime ?? true,
    });
  }

  enqueue(input = {}) {
    const annotation = this.capture(input);
    const existing = this.pending.get(annotation.id);
    if (existing) return existing.annotation;
    this.pending.set(annotation.id, {
      annotation,
      attempts: 0,
      enqueued_at_ms: Number(this.clock()),
      last_error: null,
    });
    return annotation;
  }

  pendingAnnotations() {
    return [...this.pending.values()].map((record) => ({ ...record.annotation }));
  }

  get pendingCount() {
    return this.pending.size;
  }

  discard(id) {
    return this.pending.delete(String(id));
  }

  async deliverPending(id, options = {}) {
    const key = String(id);
    const record = this.pending.get(key);
    if (!record) {
      throw new MeetingTimelineSdkError(`annotation is not pending: ${key}`, { annotation_id: key });
    }
    const maxAttempts = positiveInteger(options.maxAttempts ?? options.max_attempts, this.maxAttempts);
    const retryDelayMs = nonNegativeNumber(options.retryDelayMs ?? options.retry_delay_ms, this.retryDelayMs);
    const maxRetryDelayMs = nonNegativeNumber(options.maxRetryDelayMs ?? options.max_retry_delay_ms, this.maxRetryDelayMs);
    let lastError = null;

    for (let localAttempt = 1; localAttempt <= maxAttempts; localAttempt += 1) {
      record.attempts += 1;
      try {
        const response = await this.client.insertMark(record.annotation, { raw: true });
        const localVisibleAtMs = Number(this.clock());
        const evidenceVisibleAtMs = Number(response?.annotation_evidence?.row?.visible_at_ms);
        const visibleAtMs = Number.isFinite(evidenceVisibleAtMs) ? evidenceVisibleAtMs : localVisibleAtMs;
        this.pending.delete(key);
        return {
          accepted: true,
          annotation: record.annotation,
          response,
          attempts: record.attempts,
          captured_at_ms: record.annotation.captured_at_ms,
          visible_at_ms: visibleAtMs,
          delivery_latency_ms: Math.max(0, visibleAtMs - record.annotation.captured_at_ms),
        };
      } catch (error) {
        lastError = error;
        record.last_error = String(error?.message ?? error);
        if (!isRetryableProducerError(error) || localAttempt >= maxAttempts) break;
        const delayMs = Math.min(maxRetryDelayMs, retryDelayMs * (2 ** (localAttempt - 1)));
        if (delayMs > 0) await this.sleep(delayMs);
      }
    }

    const failure = new MeetingTimelineSdkError('annotation delivery failed; item remains pending', {
      annotation_id: key,
      attempts: record.attempts,
      pending: true,
      cause: lastError,
    });
    failure.cause = lastError;
    throw failure;
  }

  async publish(input = {}, options = {}) {
    const annotation = this.enqueue(input);
    return this.deliverPending(annotation.id, options);
  }

  async publishBatch(inputs = [], options = {}) {
    const rows = Array.isArray(inputs) ? inputs : inputs.annotations ?? inputs.items ?? [];
    const annotations = rows.map((input) => this.enqueue(input));
    const requestedIds = new Set(annotations.map((annotation) => annotation.id));
    const report = await this.flush(options);
    return {
      ...report,
      requested_count: annotations.length,
      requested_ids: [...requestedIds],
      deliveries: report.deliveries.filter((delivery) => requestedIds.has(delivery.annotation?.id)),
      failures: report.failures.filter((failure) => requestedIds.has(failure.annotation_id)),
    };
  }

  async flush(options = {}) {
    if (this.flushPromise) return this.flushPromise;
    this.flushPromise = (async () => {
      const deliveries = [];
      const failures = [];
      const ids = [...this.pending.keys()];
      for (const id of ids) {
        if (!this.pending.has(id)) continue;
        try {
          deliveries.push(await this.deliverPending(id, options));
        } catch (error) {
          failures.push({
            annotation_id: id,
            error: String(error?.message ?? error),
            attempts: this.pending.get(id)?.attempts ?? 0,
          });
        }
      }
      const report = {
        accepted: failures.length === 0,
        attempted_count: ids.length,
        delivered_count: deliveries.length,
        failed_count: failures.length,
        pending_count: this.pending.size,
        deliveries,
        failures,
      };
      if (failures.length > 0 && options.throwOnError === true) {
        throw new MeetingTimelineSdkError('annotation producer flush failed', report);
      }
      return report;
    })();
    try {
      return await this.flushPromise;
    } finally {
      this.flushPromise = null;
    }
  }
}

export function createMeetingTimelineAnnotationProducer(options = {}) {
  return new MeetingTimelineAnnotationProducer(options);
}
