import type { IncomingMessage, ServerResponse } from 'node:http';
import type { PlatformWebhookRouterOptions, PlatformWebhookRouterRequest } from './platform-webhook-router.mjs';

export interface PlatformWebhookNodeOptions extends PlatformWebhookRouterOptions {
  method?: string;
  url?: string;
  path?: string;
  headers?: Record<string, unknown>;
  requestBaseUrl?: string;
  request_base_url?: string;
  parseJson?: boolean;
  parse_json?: boolean;
  strictJson?: boolean;
  strict_json?: boolean;
  parseBody?: (rawBody: string, context: { headers: Record<string, unknown> }) => unknown;
  rawBody?: string | Buffer;
  raw_body?: string | Buffer;
  maxBodyBytes?: number;
  max_body_bytes?: number;
  receivedAtMs?: number | string | Date;
  received_at_ms?: number | string | Date;
  status?: number;
  [key: string]: unknown;
}

export interface MeetingPlatformNodeHandler {
  (req: IncomingMessage & Record<string, unknown>, res?: ServerResponse & Record<string, unknown>, options?: PlatformWebhookNodeOptions): Promise<unknown>;
  router: unknown;
  requestFromNodeRequest(req: IncomingMessage & Record<string, unknown>, options?: PlatformWebhookNodeOptions): Promise<PlatformWebhookRouterRequest>;
  writeNodeResponse(res: ServerResponse & Record<string, unknown>, result?: Record<string, unknown>, options?: PlatformWebhookNodeOptions): unknown;
}

export function platformWebhookRequestFromNodeRequest(
  req: IncomingMessage & Record<string, unknown>,
  options?: PlatformWebhookNodeOptions,
): Promise<PlatformWebhookRouterRequest>;

export function writePlatformWebhookNodeResponse(
  res: ServerResponse & Record<string, unknown>,
  result?: Record<string, unknown>,
  options?: PlatformWebhookNodeOptions,
): unknown;

export function createMeetingPlatformNodeHandler(
  clientOrRouter: unknown,
  defaults?: PlatformWebhookNodeOptions,
): MeetingPlatformNodeHandler;

export function createMeetingPlatformExpressMiddleware(
  clientOrRouter: unknown,
  defaults?: PlatformWebhookNodeOptions,
): (req: IncomingMessage & Record<string, unknown>, res: ServerResponse & Record<string, unknown>, next?: (error?: unknown) => unknown) => Promise<unknown>;
