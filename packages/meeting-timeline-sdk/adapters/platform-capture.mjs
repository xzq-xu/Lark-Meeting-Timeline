import { createHash } from 'node:crypto';

import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  buildMeetingPlatformAcceptanceSummary,
  buildPlatformAcceptanceReport,
} from './platform-acceptance.mjs';
import { platformWebhookRequestFromWebRequest } from './platform-http.mjs';
import { matchPlatformWebhookRoute } from './platform-webhook-router.mjs';
import { normalizeMeetingPlatform } from './platform-setup.mjs';

export const PLATFORM_EVENT_CAPTURE_SCHEMA = 'meeting_platform_event_capture';
export const PLATFORM_EVENT_CAPTURE_SCHEMA_VERSION = 1;

const DEFAULT_REDACTED_HEADERS = Object.freeze([
  'authorization',
  'cookie',
  'set-cookie',
  'x-api-key',
  'x-zm-signature',
  'x-spark-signature',
]);

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function maybeJson(value) {
  if (value == null || value === '') return undefined;
  if (typeof value !== 'string') return value;
  try {
    return JSON.parse(value);
  } catch {
    return value;
  }
}

function bodyText(value) {
  if (value == null) return '';
  if (Buffer.isBuffer(value)) return value.toString('utf8');
  if (typeof value === 'string') return value;
  return JSON.stringify(value);
}

function sha256Hex(value) {
  const text = bodyText(value);
  return text ? createHash('sha256').update(text).digest('hex') : undefined;
}

function headersToObject(headers = {}) {
  if (!headers) return {};
  if (typeof headers.forEach === 'function') {
    const out = {};
    headers.forEach((value, key) => {
      out[String(key).toLowerCase()] = value;
    });
    return out;
  }
  return Object.fromEntries(
    Object.entries(headers).map(([key, value]) => [
      String(key).toLowerCase(),
      Array.isArray(value) ? value.join(', ') : value,
    ]),
  );
}

function redactHeaders(headers = {}, options = {}) {
  const source = headersToObject(headers);
  const redacted = new Set([
    ...DEFAULT_REDACTED_HEADERS,
    ...(options.redactedHeaders ?? options.redacted_headers ?? []),
  ].map((item) => String(item).toLowerCase()));
  return Object.fromEntries(Object.entries(source).map(([key, value]) => (
    redacted.has(String(key).toLowerCase()) ? [key, '[redacted]'] : [key, value]
  )));
}

function requestPath(input = {}) {
  if (input.path) return String(input.path);
  if (input.url) {
    try {
      return new URL(String(input.url), 'http://localhost').pathname;
    } catch {
      return String(input.url).split('?')[0] || '/';
    }
  }
  return undefined;
}

function normalizeCaptureArgs(platformOrInput, payload, options = {}) {
  if (platformOrInput && typeof platformOrInput === 'object' && !Array.isArray(platformOrInput)) {
    const input = platformOrInput;
    return {
      input,
      platform: firstNonEmpty(input.platform, input.provider, input.adapter, options.platform),
      body: firstNonEmpty(input.body, input.payload, input.raw_event, input.rawEvent, input.raw, payload),
      rawBody: firstNonEmpty(input.rawBody, input.raw_body, options.rawBody, options.raw_body),
      options,
    };
  }
  return {
    input: {},
    platform: firstNonEmpty(platformOrInput, options.platform),
    body: payload,
    rawBody: firstNonEmpty(options.rawBody, options.raw_body),
    options,
  };
}

function capturePlatform(input = {}, explicitPlatform, options = {}) {
  if (explicitPlatform) return normalizeMeetingPlatform(explicitPlatform);
  const route = matchPlatformWebhookRoute(input, options);
  if (route?.platform) return route.platform;
  throw new MeetingTimelineSdkError('platform is required to capture a meeting platform event', {
    reason: 'missing_platform',
  });
}

function captureLabel(platform, body = {}, capturedAtMs, hash) {
  const explicit = firstNonEmpty(body?.id, body?.event_id, body?.eventId, body?.event, body?.type);
  return `${platform}:${explicit ?? hash?.slice(0, 10) ?? capturedAtMs}`;
}

export function capturePlatformWebhookEvent(platformOrInput, payload, options = {}) {
  const { input, platform, body, rawBody } = normalizeCaptureArgs(platformOrInput, payload, options);
  const key = capturePlatform(input, platform, options);
  const capturedAtMs = Number(firstNonEmpty(
    options.capturedAtMs,
    options.captured_at_ms,
    input.capturedAtMs,
    input.captured_at_ms,
    Date.now(),
  ));
  const receivedAtMs = firstNonEmpty(
    options.receivedAtMs,
    options.received_at_ms,
    input.receivedAtMs,
    input.received_at_ms,
    input.received_at,
    capturedAtMs,
  );
  const parsedBody = maybeJson(body);
  const rawText = rawBody == null ? bodyText(parsedBody) : bodyText(rawBody);
  const rawHash = sha256Hex(rawText);
  const route = matchPlatformWebhookRoute(input, options);
  const label = firstNonEmpty(options.label, input.label, captureLabel(key, parsedBody, capturedAtMs, rawHash));

  return compactObject({
    schema: PLATFORM_EVENT_CAPTURE_SCHEMA,
    schema_version: PLATFORM_EVENT_CAPTURE_SCHEMA_VERSION,
    id: firstNonEmpty(options.id, input.id, `${key}-${capturedAtMs}-${rawHash?.slice(0, 12) ?? 'event'}`),
    platform: key,
    label,
    captured_at_ms: capturedAtMs,
    received_at_ms: receivedAtMs,
    method: input.method,
    url: input.url,
    path: requestPath(input),
    route: route ? {
      kind: route.kind,
      platform: route.platform,
      slug: route.slug,
      path: route.path,
    } : undefined,
    headers: options.includeHeaders === false || options.include_headers === false
      ? undefined
      : redactHeaders(input.headers ?? options.headers, options),
    body: parsedBody,
    raw_body_sha256: rawHash,
    raw_body: options.includeRawBody === true || options.include_raw_body === true ? rawText : undefined,
    metadata: options.metadata,
  });
}

export async function capturePlatformWebRequest(platformOrRequest, requestOrOptions = {}, maybeOptions = {}) {
  const explicitPlatform = typeof platformOrRequest === 'string' ? platformOrRequest : undefined;
  const request = explicitPlatform ? requestOrOptions : platformOrRequest;
  const options = explicitPlatform ? maybeOptions : requestOrOptions;
  const input = await platformWebhookRequestFromWebRequest(request, options);
  return capturePlatformWebhookEvent({
    ...input,
    platform: explicitPlatform ?? options.platform,
  }, undefined, options);
}

function asArray(value) {
  return Array.isArray(value) ? value : value == null ? [] : [value];
}

export function buildPlatformCaptureSamples(records = [], options = {}) {
  const rows = asArray(records);
  const samples = {};
  for (const record of rows) {
    if (!record) continue;
    const platform = normalizeMeetingPlatform(record.platform ?? record.provider ?? record.adapter);
    if (!samples[platform]) samples[platform] = [];
    samples[platform].push(compactObject({
      platform,
      label: record.label ?? record.id,
      body: record.body ?? maybeJson(record.raw_body),
      receivedAtMs: record.received_at_ms ?? record.captured_at_ms,
      capture_id: record.id,
      options: {
        receivedAtMs: record.received_at_ms ?? record.captured_at_ms,
      },
    }));
  }
  return samples;
}

export function buildPlatformCaptureAcceptanceReport(platform, records = [], options = {}) {
  const samples = buildPlatformCaptureSamples(records, options);
  return buildPlatformAcceptanceReport(platform, {
    ...options,
    samples,
  });
}

export function buildMeetingPlatformCaptureAcceptanceSummary(records = [], options = {}) {
  const samples = buildPlatformCaptureSamples(records, options);
  return buildMeetingPlatformAcceptanceSummary({
    ...options,
    samples,
  });
}

export function serializePlatformCaptureRecord(record = {}) {
  return JSON.stringify(record);
}

export function parsePlatformCaptureRecord(line) {
  const parsed = JSON.parse(String(line));
  if (parsed?.schema !== PLATFORM_EVENT_CAPTURE_SCHEMA) {
    throw new MeetingTimelineSdkError('Invalid platform capture record schema', {
      reason: 'invalid_capture_schema',
      schema: parsed?.schema,
    });
  }
  return parsed;
}

export function parsePlatformCaptureJsonl(text = '') {
  return String(text)
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => parsePlatformCaptureRecord(line));
}
