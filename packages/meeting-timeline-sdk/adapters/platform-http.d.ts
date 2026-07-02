import type { PlatformWebhookRouterOptions, PlatformWebhookRouterRequest } from './platform-webhook-router.mjs';

export interface PlatformWebhookHttpOptions extends PlatformWebhookRouterOptions {
  method?: string;
  url?: string;
  headers?: Headers | Record<string, unknown>;
  parseJson?: boolean;
  parse_json?: boolean;
  strictJson?: boolean;
  strict_json?: boolean;
  parseBody?: (rawBody: string, context: { headers: Record<string, unknown> }) => unknown;
  ResponseCtor?: typeof Response;
  responseCtor?: typeof Response;
  status?: number;
  receivedAtMs?: number | string | Date;
  received_at_ms?: number | string | Date;
  [key: string]: unknown;
}

export interface MeetingPlatformFetchHandler {
  (request: Request, options?: PlatformWebhookHttpOptions): Promise<Response>;
  router: unknown;
  requestFromWebRequest(request: Request, options?: PlatformWebhookHttpOptions): Promise<PlatformWebhookRouterRequest>;
  responseToWebResponse(result?: Record<string, unknown>, options?: PlatformWebhookHttpOptions): Response;
}

export function platformWebhookRequestFromWebRequest(
  request: Request,
  options?: PlatformWebhookHttpOptions,
): Promise<PlatformWebhookRouterRequest>;

export function platformWebhookResponseToWebResponse(
  result?: Record<string, unknown>,
  options?: PlatformWebhookHttpOptions,
): Response;

export function createMeetingPlatformFetchHandler(
  clientOrRouter: unknown,
  defaults?: PlatformWebhookHttpOptions,
): MeetingPlatformFetchHandler;
