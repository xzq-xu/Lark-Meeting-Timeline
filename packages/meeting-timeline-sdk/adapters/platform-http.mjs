import { MeetingTimelineSdkError } from '../index.mjs';
import { createMeetingPlatformWebhookRouter } from './platform-webhook-router.mjs';

function isRouter(value) {
  return typeof value === 'function'
    && typeof value.routeTable === 'function'
    && typeof value.status === 'function';
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

function contentTypeOf(headers = {}) {
  const rows = headersToObject(headers);
  return String(rows['content-type'] ?? rows['Content-Type'] ?? '').toLowerCase();
}

function shouldReadBody(method) {
  return !['GET', 'HEAD'].includes(String(method || 'GET').toUpperCase());
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

function responseBodyAndHeaders(result = {}, options = {}) {
  const headers = {
    ...(result.headers ?? {}),
    ...(options.headers ?? {}),
  };
  const body = result.body;
  if (body == null) return { headers, body: '' };
  if (typeof body === 'string' || body instanceof Uint8Array || body instanceof ArrayBuffer) {
    return { headers, body };
  }
  if (!Object.keys(headers).some((key) => key.toLowerCase() === 'content-type')) {
    headers['content-type'] = 'application/json; charset=utf-8';
  }
  return { headers, body: JSON.stringify(body) };
}

export async function platformWebhookRequestFromWebRequest(request, options = {}) {
  if (!request || typeof request !== 'object') {
    throw new MeetingTimelineSdkError('Web Request object is required for platform webhook HTTP handling');
  }
  const method = String(request.method ?? options.method ?? 'GET').toUpperCase();
  const headers = headersToObject(request.headers ?? options.headers ?? {});
  const rawBody = shouldReadBody(method) && typeof request.text === 'function'
    ? await request.clone().text()
    : '';
  return {
    method,
    url: String(request.url ?? options.url ?? ''),
    headers,
    rawBody,
    body: shouldReadBody(method)
      ? parseRawBody(rawBody, headers, options)
      : undefined,
    receivedAtMs: options.receivedAtMs ?? options.received_at_ms,
  };
}

export function platformWebhookResponseToWebResponse(result = {}, options = {}) {
  const ResponseCtor = options.ResponseCtor ?? options.responseCtor ?? globalThis.Response;
  if (typeof ResponseCtor !== 'function') {
    throw new MeetingTimelineSdkError('Response constructor is required for platform webhook HTTP handling');
  }
  const { headers, body } = responseBodyAndHeaders(result, options);
  return new ResponseCtor(body, {
    status: Number(result.status ?? options.status ?? 200),
    headers,
  });
}

export function createMeetingPlatformFetchHandler(clientOrRouter, defaults = {}) {
  const router = isRouter(clientOrRouter)
    ? clientOrRouter
    : createMeetingPlatformWebhookRouter(clientOrRouter, defaults);

  async function handle(request, options = {}) {
    const merged = { ...defaults, ...options };
    let input;
    try {
      input = await platformWebhookRequestFromWebRequest(request, merged);
    } catch (error) {
      return platformWebhookResponseToWebResponse({
        status: error.status ?? 400,
        body: {
          error: error.message ?? String(error),
          details: error.details,
        },
      }, merged);
    }
    const result = await router(input, merged);
    return platformWebhookResponseToWebResponse(result, merged);
  }

  handle.router = router;
  handle.requestFromWebRequest = (request, options = {}) => (
    platformWebhookRequestFromWebRequest(request, { ...defaults, ...options })
  );
  handle.responseToWebResponse = platformWebhookResponseToWebResponse;
  return handle;
}
