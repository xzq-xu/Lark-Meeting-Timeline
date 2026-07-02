import { createHmac, timingSafeEqual } from 'node:crypto';

import { MeetingTimelineSdkError, compactObject } from '../index.mjs';

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function headerValue(headers = {}, name) {
  if (!headers || !name) return undefined;
  if (typeof headers.get === 'function') return headers.get(name) ?? headers.get(name.toLowerCase()) ?? undefined;
  const lowerName = String(name).toLowerCase();
  const entry = Object.entries(headers).find(([key]) => key.toLowerCase() === lowerName);
  const value = entry?.[1];
  return Array.isArray(value) ? value[0] : value;
}

function bodyText(rawBody) {
  if (rawBody == null) return '';
  if (Buffer.isBuffer(rawBody)) return rawBody.toString('utf8');
  if (typeof rawBody === 'string') return rawBody;
  return JSON.stringify(rawBody);
}

function safeEqual(left, right) {
  const leftBuffer = Buffer.from(String(left ?? ''), 'utf8');
  const rightBuffer = Buffer.from(String(right ?? ''), 'utf8');
  if (leftBuffer.length !== rightBuffer.length) return false;
  return timingSafeEqual(leftBuffer, rightBuffer);
}

function hmacSha256Hex(secret, message) {
  return createHmac('sha256', String(secret)).update(String(message)).digest('hex');
}

export function buildZoomUrlValidationResponse(input = {}, options = {}) {
  const secretToken = firstNonEmpty(options.secretToken, options.zoomSecretToken, process.env.ZOOM_WEBHOOK_SECRET_TOKEN);
  if (!secretToken) {
    throw new MeetingTimelineSdkError('Zoom webhook secret token is required for URL validation', {
      fieldName: 'secretToken',
    });
  }
  const plainToken = firstNonEmpty(input.plainToken, input.plain_token, input.payload?.plainToken);
  if (!plainToken) {
    throw new MeetingTimelineSdkError('Zoom URL validation plainToken is required', {
      fieldName: 'plainToken',
    });
  }
  return {
    plainToken: String(plainToken),
    encryptedToken: hmacSha256Hex(secretToken, plainToken),
  };
}

export function verifyZoomWebhookEvent({ headers = {}, rawBody = '', body, secretToken, toleranceMs = 5 * 60_000 } = {}) {
  const token = firstNonEmpty(secretToken, process.env.ZOOM_WEBHOOK_SECRET_TOKEN);
  if (!token) {
    return { ok: true, skipped: true, reason: 'zoom_secret_token_not_configured' };
  }
  const timestamp = headerValue(headers, 'x-zm-request-timestamp');
  const signature = headerValue(headers, 'x-zm-signature');
  if (!timestamp || !signature) {
    return { ok: false, reason: 'zoom_signature_headers_missing' };
  }
  const timestampMs = Number(timestamp) * 1000;
  if (Number.isFinite(timestampMs) && toleranceMs > 0 && Math.abs(Date.now() - timestampMs) > toleranceMs) {
    return { ok: false, reason: 'zoom_signature_timestamp_out_of_tolerance', timestamp };
  }
  const payloadText = rawBody !== '' && rawBody != null ? bodyText(rawBody) : bodyText(body);
  const expected = `v0=${hmacSha256Hex(token, `v0:${timestamp}:${payloadText}`)}`;
  if (!safeEqual(expected, signature)) {
    return { ok: false, reason: 'zoom_signature_mismatch', timestamp };
  }
  return { ok: true, skipped: false, reason: 'verified', timestamp };
}

export function microsoftGraphValidationResponse(urlOrToken) {
  if (urlOrToken == null) return null;
  if (typeof urlOrToken === 'string') return urlOrToken;
  const token = urlOrToken.searchParams?.get?.('validationToken')
    ?? urlOrToken.validationToken
    ?? urlOrToken.validation_token;
  return token == null ? null : String(token);
}

export function verifyMicrosoftGraphClientState(input = {}, options = {}) {
  const expected = firstNonEmpty(options.clientState, process.env.MICROSOFT_GRAPH_CLIENT_STATE);
  if (!expected) {
    return { ok: true, skipped: true, reason: 'microsoft_graph_client_state_not_configured' };
  }
  const notifications = Array.isArray(input?.value) ? input.value : Array.isArray(input) ? input : [input];
  const mismatches = notifications
    .map((item, index) => ({ index, clientState: item?.clientState ?? item?.client_state }))
    .filter((item) => !safeEqual(item.clientState, expected));
  if (mismatches.length) {
    return {
      ok: false,
      reason: 'microsoft_graph_client_state_mismatch',
      mismatch_count: mismatches.length,
    };
  }
  return {
    ok: true,
    skipped: false,
    reason: 'verified',
    notification_count: notifications.length,
  };
}

export function verifyGooglePubSubBearer({ headers = {}, expectedToken } = {}) {
  const token = firstNonEmpty(expectedToken, process.env.GOOGLE_PUBSUB_BEARER_TOKEN);
  if (!token) {
    return { ok: true, skipped: true, reason: 'google_pubsub_bearer_not_configured' };
  }
  const authorization = String(headerValue(headers, 'authorization') ?? '');
  const bearer = authorization.replace(/^Bearer\s+/i, '');
  if (!safeEqual(bearer, token)) {
    return { ok: false, reason: 'google_pubsub_bearer_mismatch' };
  }
  return { ok: true, skipped: false, reason: 'verified' };
}

export function platformWebhookVerificationStatus(platform, result = {}) {
  return compactObject({
    platform,
    ok: result.ok,
    skipped: result.skipped,
    reason: result.reason,
    timestamp: result.timestamp,
    notification_count: result.notification_count,
    mismatch_count: result.mismatch_count,
  });
}
