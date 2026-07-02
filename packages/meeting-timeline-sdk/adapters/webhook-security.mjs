import { createHmac, createPublicKey, createVerify, timingSafeEqual } from 'node:crypto';

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

function bearerToken(headers = {}) {
  const authorization = String(headerValue(headers, 'authorization') ?? '');
  const match = authorization.match(/^Bearer\s+(.+)$/i);
  return match?.[1] ?? null;
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

function base64UrlDecode(text) {
  const padded = String(text).replace(/-/g, '+').replace(/_/g, '/').padEnd(Math.ceil(String(text).length / 4) * 4, '=');
  return Buffer.from(padded, 'base64');
}

function parseJwtPart(text, label) {
  try {
    return JSON.parse(base64UrlDecode(text).toString('utf8'));
  } catch (error) {
    throw new MeetingTimelineSdkError(`Invalid JWT ${label}`, { label, cause: error.message });
  }
}

function parseJwt(token) {
  const parts = String(token || '').split('.');
  if (parts.length !== 3 || parts.some((part) => !part)) {
    throw new MeetingTimelineSdkError('Malformed JWT', { reason: 'jwt_malformed' });
  }
  return {
    header: parseJwtPart(parts[0], 'header'),
    payload: parseJwtPart(parts[1], 'payload'),
    signingInput: `${parts[0]}.${parts[1]}`,
    signature: base64UrlDecode(parts[2]),
  };
}

function normalizeList(value) {
  if (value == null || value === '') return [];
  return Array.isArray(value) ? value.map(String) : [String(value)];
}

function maxAgeMs(cacheControl = '') {
  const match = String(cacheControl).match(/max-age=(\d+)/i);
  return match ? Number(match[1]) * 1000 : 0;
}

let googleJwksCache = null;

async function googleJwks(options = {}) {
  if (options.jwks) return options.jwks;
  if (options.publicKeys || options.certs) {
    const rows = Object.entries(options.publicKeys ?? options.certs ?? {}).map(([kid, key]) => ({ kid, key }));
    return { keys: rows };
  }
  const now = Date.now();
  const jwksUrl = options.jwksUrl ?? 'https://www.googleapis.com/oauth2/v3/certs';
  if (googleJwksCache?.url === jwksUrl && googleJwksCache.expiresAt > now) return googleJwksCache.jwks;
  const fetchImpl = options.fetchImpl ?? globalThis.fetch;
  if (typeof fetchImpl !== 'function') {
    throw new MeetingTimelineSdkError('fetch implementation is required to load Google JWKS');
  }
  const response = await fetchImpl(jwksUrl);
  if (!response?.ok) {
    throw new MeetingTimelineSdkError('Failed to load Google JWKS', { status: response?.status });
  }
  const jwks = await response.json();
  googleJwksCache = {
    url: jwksUrl,
    jwks,
    expiresAt: now + (maxAgeMs(response.headers?.get?.('cache-control')) || 60 * 60 * 1000),
  };
  return jwks;
}

function publicKeyForJwt(header = {}, jwks = {}) {
  const kid = header.kid;
  const rows = Array.isArray(jwks.keys) ? jwks.keys : [];
  const row = rows.find((item) => !kid || item.kid === kid);
  if (!row) return null;
  if (row.key?.kty) return createPublicKey({ key: row.key, format: 'jwk' });
  if (row.key) return createPublicKey(row.key);
  if (row.kty) return createPublicKey({ key: row, format: 'jwk' });
  if (typeof row === 'string') return createPublicKey(row);
  return null;
}

function verifyJwtSignature({ signingInput, signature }, publicKey) {
  const verifier = createVerify('RSA-SHA256');
  verifier.update(signingInput);
  verifier.end();
  return verifier.verify(publicKey, signature);
}

function verifyGoogleClaims(payload = {}, options = {}) {
  const nowSec = Math.floor((options.nowMs ?? Date.now()) / 1000);
  const toleranceSec = Number(options.clockToleranceSec ?? 300);
  const issuers = normalizeList(options.issuer ?? ['https://accounts.google.com', 'accounts.google.com']);
  if (issuers.length && !issuers.includes(String(payload.iss ?? ''))) {
    return { ok: false, reason: 'google_pubsub_oidc_issuer_mismatch', issuer: payload.iss };
  }
  if (payload.exp == null || Number(payload.exp) + toleranceSec < nowSec) {
    return { ok: false, reason: 'google_pubsub_oidc_token_expired', expires_at: payload.exp };
  }
  if (payload.nbf != null && Number(payload.nbf) - toleranceSec > nowSec) {
    return { ok: false, reason: 'google_pubsub_oidc_token_not_yet_valid' };
  }
  const expectedAudiences = normalizeList(firstNonEmpty(
    options.expectedAudience,
    options.audience,
    process.env.GOOGLE_PUBSUB_OIDC_AUDIENCE,
  ));
  const actualAudiences = normalizeList(payload.aud);
  if (expectedAudiences.length && !actualAudiences.some((item) => expectedAudiences.includes(item))) {
    return {
      ok: false,
      reason: 'google_pubsub_oidc_audience_mismatch',
      audience: payload.aud,
    };
  }
  const expectedEmail = firstNonEmpty(
    options.serviceAccountEmail,
    options.expectedEmail,
    process.env.GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL,
  );
  if (expectedEmail && !safeEqual(payload.email, expectedEmail)) {
    return {
      ok: false,
      reason: 'google_pubsub_oidc_email_mismatch',
      email: payload.email,
    };
  }
  return {
    ok: true,
    reason: 'verified',
    issuer: payload.iss,
    audience: payload.aud,
    email: payload.email,
    expires_at: payload.exp,
  };
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
  const bearer = bearerToken(headers) ?? '';
  if (!safeEqual(bearer, token)) {
    return { ok: false, reason: 'google_pubsub_bearer_mismatch' };
  }
  return { ok: true, skipped: false, reason: 'verified' };
}

export async function verifyGooglePubSubOidcJwt({
  headers = {},
  token,
  expectedAudience,
  audience,
  serviceAccountEmail,
  expectedEmail,
  issuer,
  jwks,
  publicKeys,
  certs,
  jwksUrl,
  fetchImpl,
  nowMs,
  clockToleranceSec,
} = {}) {
  const jwt = firstNonEmpty(token, bearerToken(headers));
  if (!jwt) return { ok: false, reason: 'google_pubsub_oidc_token_missing' };
  let parsed;
  try {
    parsed = parseJwt(jwt);
  } catch (error) {
    return { ok: false, reason: error.details?.reason ?? 'google_pubsub_oidc_token_malformed' };
  }
  if (parsed.header.alg !== 'RS256') {
    return { ok: false, reason: 'google_pubsub_oidc_unsupported_algorithm', algorithm: parsed.header.alg };
  }
  if (!parsed.header.kid) return { ok: false, reason: 'google_pubsub_oidc_kid_missing' };
  let keys;
  try {
    keys = await googleJwks({ jwks, publicKeys, certs, jwksUrl, fetchImpl });
  } catch (error) {
    return { ok: false, reason: 'google_pubsub_oidc_jwks_unavailable', error: error.message };
  }
  const publicKey = publicKeyForJwt(parsed.header, keys);
  if (!publicKey) return { ok: false, reason: 'google_pubsub_oidc_signing_key_not_found', kid: parsed.header.kid };
  if (!verifyJwtSignature(parsed, publicKey)) {
    return { ok: false, reason: 'google_pubsub_oidc_signature_mismatch', kid: parsed.header.kid };
  }
  const claims = verifyGoogleClaims(parsed.payload, {
    expectedAudience,
    audience,
    serviceAccountEmail,
    expectedEmail,
    issuer,
    nowMs,
    clockToleranceSec,
  });
  return {
    ...claims,
    skipped: false,
    kid: parsed.header.kid,
  };
}

export function platformWebhookVerificationStatus(platform, result = {}) {
  return compactObject({
    platform,
    ok: result.ok,
    skipped: result.skipped,
    reason: result.reason,
    timestamp: result.timestamp,
    issuer: result.issuer,
    audience: result.audience,
    email: result.email,
    expires_at: result.expires_at,
    notification_count: result.notification_count,
    mismatch_count: result.mismatch_count,
  });
}
