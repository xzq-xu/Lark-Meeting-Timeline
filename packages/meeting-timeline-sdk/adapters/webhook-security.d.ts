export interface WebhookVerificationResult {
  ok: boolean;
  skipped?: boolean;
  reason?: string;
  timestamp?: string | number;
  notification_count?: number;
  mismatch_count?: number;
  [key: string]: unknown;
}

export interface ZoomWebhookVerificationInput {
  headers?: Headers | Record<string, string | string[] | undefined>;
  rawBody?: string | Buffer;
  body?: unknown;
  secretToken?: string;
  toleranceMs?: number;
}

export function buildZoomUrlValidationResponse(
  input?: Record<string, unknown>,
  options?: { secretToken?: string; zoomSecretToken?: string },
): { plainToken: string; encryptedToken: string };

export function verifyZoomWebhookEvent(input?: ZoomWebhookVerificationInput): WebhookVerificationResult;

export function microsoftGraphValidationResponse(
  urlOrToken?: URL | string | { validationToken?: string; validation_token?: string },
): string | null;

export function verifyMicrosoftGraphClientState(
  input?: unknown,
  options?: { clientState?: string },
): WebhookVerificationResult;

export function verifyGooglePubSubBearer(input?: {
  headers?: Headers | Record<string, string | string[] | undefined>;
  expectedToken?: string;
}): WebhookVerificationResult;

export function platformWebhookVerificationStatus(
  platform: string,
  result?: WebhookVerificationResult,
): Record<string, unknown>;
