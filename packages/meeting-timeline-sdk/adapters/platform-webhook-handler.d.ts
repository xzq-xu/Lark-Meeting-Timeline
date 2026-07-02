import type { MeetingTimelineClient } from '../index.mjs';
import type { PlatformEventIngestResult } from './platform-ingest.mjs';
import type { WebhookVerificationResult } from './webhook-security.mjs';

export interface PlatformWebhookRequestInput {
  platform?: string;
  provider?: string;
  adapter?: string;
  method?: string;
  url?: string;
  query?: Record<string, string | string[] | number | boolean | null | undefined>;
  searchParams?: URLSearchParams;
  headers?: Headers | Record<string, string | string[] | undefined>;
  body?: unknown;
  rawBody?: string | Buffer;
  raw_body?: string | Buffer;
  receivedAtMs?: number | string | Date;
  received_at_ms?: number | string | Date;
  received_at?: number | string | Date;
}

export interface PlatformWebhookHandlerOptions {
  platform?: string;
  verify?: boolean;
  receivedAtMs?: number | string | Date;
  normalizerOptions?: Record<string, unknown>;
  normalizer_options?: Record<string, unknown>;
  applyOptions?: Record<string, unknown>;
  apply_options?: Record<string, unknown>;
  secretToken?: string;
  secret?: string;
  bearerToken?: string;
  clientState?: string;
  zoomSecretToken?: string;
  microsoftGraphClientState?: string;
  googlePubSubBearerToken?: string;
  googlePubSubOidcAudience?: string;
  googlePubSubAudience?: string;
  googlePubSubServiceAccountEmail?: string;
  googlePubSubJwksUrl?: string;
  zoom?: Record<string, unknown>;
  zoomWebhook?: Record<string, unknown>;
  zoom_webhook?: Record<string, unknown>;
  webex?: Record<string, unknown>;
  webexWebhook?: Record<string, unknown>;
  webex_webhook?: Record<string, unknown>;
  microsoftTeams?: Record<string, unknown>;
  microsoft_teams?: Record<string, unknown>;
  graph?: Record<string, unknown>;
  microsoftGraph?: Record<string, unknown>;
  googleMeet?: Record<string, unknown>;
  google_meet?: Record<string, unknown>;
  google?: Record<string, unknown>;
  googlePubSub?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface PlatformWebhookHandlerResponse {
  status: number;
  headers: Record<string, string>;
  body: unknown;
}

export interface PlatformWebhookIngestBody extends Omit<PlatformEventIngestResult, 'adapter'> {
  ok: true;
  adapter: {
    key: string;
    source: string;
    aliases: readonly string[];
  };
  verification?: WebhookVerificationResult | null;
  signal_count: number;
}

export function verifyPlatformWebhook(
  platform: string,
  request?: PlatformWebhookRequestInput,
  options?: PlatformWebhookHandlerOptions,
): Promise<WebhookVerificationResult>;

export function handlePlatformWebhookRequest(
  client: MeetingTimelineClient,
  input?: PlatformWebhookRequestInput,
  options?: PlatformWebhookHandlerOptions,
): Promise<PlatformWebhookHandlerResponse>;

export function createPlatformWebhookHandler(
  client: MeetingTimelineClient,
  defaults?: PlatformWebhookHandlerOptions,
): (input?: PlatformWebhookRequestInput, options?: PlatformWebhookHandlerOptions) => Promise<PlatformWebhookHandlerResponse>;
