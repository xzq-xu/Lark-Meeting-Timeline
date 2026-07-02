import { MeetingTimelineSdkError } from '../index.mjs';
import { createMeetingPlatformWebhookRouter } from './platform-webhook-router.mjs';

function isRouter(value) {
  return typeof value === 'function'
    && typeof value.routeTable === 'function'
    && typeof value.status === 'function';
}

function headersToObject(headers = {}) {
  return Object.fromEntries(
    Object.entries(headers ?? {}).map(([key, value]) => [
      String(key).toLowerCase(),
      Array.isArray(value) ? value.join(', ') : value,
    ]),
  );
}

function contentTypeOf(headers = {}) {
  const rows = headersToObject(headers);
  return String(rows['content-type'] ?? '').toLowerCase();
}

function shouldReadBody(method) {
  return !['GET', 'HEAD'].includes(String(method || 'GET').toUpperCase());
}

async function readNodeBody(req, options = {}) {
  if (options.rawBody != null || options.raw_body != null) return String(options.rawBody ?? options.raw_body);
  if (req.body != null && (typeof req.body === 'string' || Buffer.isBuffer(req.body))) return String(req.body);
  if (!req || typeof req[Symbol.asyncIterator] !== 'function') return '';
  const chunks = [];
  let total = 0;
  const maxBodyBytes = Number(options.maxBodyBytes ?? options.max_body_bytes ?? 5 * 1024 * 1024);
  for await (const chunk of req) {
    const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
    total += buffer.length;
    if (total > maxBodyBytes) {
      throw new MeetingTimelineSdkError('Platform webhook request body is too large', {
        reason: 'body_too_large',
        max_body_bytes: maxBodyBytes,
      });
    }
    chunks.push(buffer);
  }
  return Buffer.concat(chunks).toString('utf8');
}

function parseRawBody(rawBody = '', headers = {}, options = {}) {
  if (rawBody == null || rawBody === '') return undefined;
  if (typeof options.parseBody === 'function') return options.parseBody(rawBody, { headers });
  if (options.parseJson === false || options.parse_json === false) return rawBody;
  const contentType = contentTypeOf(headers);
  if (!contentType || contentType.includes('json') || contentType.includes('+json')) {
    try {
      return JSON.parse(rawBody);
    } catch (error) {
      if (contentType.includes('json') || options.strictJson === true || options.strict_json === true) {
        throw new MeetingTimelineSdkError('Failed to parse platform webhook JSON body', {
          reason: 'invalid_json_body',
          cause: error.message,
        });
      }
    }
  }
  return rawBody;
}

function requestUrl(req = {}, options = {}) {
  const rawUrl = String(options.url ?? req.originalUrl ?? req.url ?? '');
  if (/^https?:\/\//i.test(rawUrl)) return rawUrl;
  const baseUrl = options.requestBaseUrl ?? options.request_base_url ?? options.baseUrl ?? options.base_url;
  if (!baseUrl) return rawUrl || '/';
  try {
    return new URL(rawUrl || '/', baseUrl).toString();
  } catch {
    return rawUrl || '/';
  }
}

export async function platformWebhookRequestFromNodeRequest(req, options = {}) {
  if (!req || typeof req !== 'object') {
    throw new MeetingTimelineSdkError('Node request object is required for platform webhook handling');
  }
  const method = String(options.method ?? req.method ?? 'GET').toUpperCase();
  const headers = headersToObject(options.headers ?? req.headers ?? {});
  const rawBody = shouldReadBody(method)
    ? await readNodeBody(req, options)
    : '';
  const body = shouldReadBody(method)
    ? parseRawBody(rawBody, headers, options)
    : undefined;
  return {
    method,
    url: requestUrl(req, options),
    path: options.path ?? req.path,
    headers,
    rawBody,
    body,
    receivedAtMs: options.receivedAtMs ?? options.received_at_ms,
  };
}

function responseBodyAndHeaders(result = {}, options = {}) {
  const headers = {
    ...(result.headers ?? {}),
    ...(options.headers ?? {}),
  };
  const body = result.body;
  if (body == null) return { headers, body: '' };
  if (typeof body === 'string' || Buffer.isBuffer(body)) return { headers, body };
  if (!Object.keys(headers).some((key) => key.toLowerCase() === 'content-type')) {
    headers['content-type'] = 'application/json; charset=utf-8';
  }
  return { headers, body: JSON.stringify(body) };
}

export function writePlatformWebhookNodeResponse(res, result = {}, options = {}) {
  if (!res || typeof res !== 'object') {
    throw new MeetingTimelineSdkError('Node response object is required for platform webhook handling');
  }
  const status = Number(result.status ?? options.status ?? 200);
  const { headers, body } = responseBodyAndHeaders(result, options);
  if (typeof res.status === 'function') res.status(status);
  else res.statusCode = status;
  for (const [key, value] of Object.entries(headers)) {
    if (typeof res.setHeader === 'function') res.setHeader(key, value);
    else if (typeof res.set === 'function') res.set(key, value);
  }
  if (typeof res.send === 'function') {
    res.send(body);
  } else if (typeof res.end === 'function') {
    res.end(body);
  }
  return res;
}

export function createMeetingPlatformNodeHandler(clientOrRouter, defaults = {}) {
  const router = isRouter(clientOrRouter)
    ? clientOrRouter
    : createMeetingPlatformWebhookRouter(clientOrRouter, defaults);

  async function handle(req, res, options = {}) {
    const merged = { ...defaults, ...options };
    let input;
    try {
      input = await platformWebhookRequestFromNodeRequest(req, merged);
    } catch (error) {
      const failure = {
        status: error.status ?? 400,
        body: {
          error: error.message ?? String(error),
          details: error.details,
        },
      };
      return res ? writePlatformWebhookNodeResponse(res, failure, merged) : failure;
    }
    const result = await router(input, merged);
    return res ? writePlatformWebhookNodeResponse(res, result, merged) : result;
  }

  handle.router = router;
  handle.requestFromNodeRequest = (req, options = {}) => (
    platformWebhookRequestFromNodeRequest(req, { ...defaults, ...options })
  );
  handle.writeNodeResponse = writePlatformWebhookNodeResponse;
  return handle;
}

export function createMeetingPlatformExpressMiddleware(clientOrRouter, defaults = {}) {
  const handler = createMeetingPlatformNodeHandler(clientOrRouter, defaults);
  return async function meetingPlatformExpressMiddleware(req, res, next) {
    try {
      return await handler(req, res);
    } catch (error) {
      if (typeof next === 'function') return next(error);
      throw error;
    }
  };
}
