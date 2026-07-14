export const PLATFORM_EVENT_CAPTURE_SCHEMA: string;
export const PLATFORM_EVENT_CAPTURE_SCHEMA_VERSION: number;

export interface PlatformCaptureOptions {
  platform?: string;
  basePath?: string;
  base_path?: string;
  capturedAtMs?: number;
  captured_at_ms?: number;
  receivedAtMs?: number | string | Date;
  received_at_ms?: number | string | Date;
  label?: string;
  id?: string;
  headers?: Headers | Record<string, unknown>;
  includeHeaders?: boolean;
  include_headers?: boolean;
  redactedHeaders?: string[];
  redacted_headers?: string[];
  includeRawBody?: boolean;
  include_raw_body?: boolean;
  rawBody?: string | Buffer;
  raw_body?: string | Buffer;
  metadata?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface PlatformCaptureRecord {
  schema: string;
  schema_version: number;
  id: string;
  platform: string;
  label?: string;
  captured_at_ms: number;
  received_at_ms?: number | string | Date;
  method?: string;
  url?: string;
  path?: string;
  route?: Record<string, unknown>;
  headers?: Record<string, unknown>;
  body?: unknown;
  raw_body_sha256?: string;
  raw_body?: string;
  metadata?: Record<string, unknown>;
}

export function capturePlatformWebhookEvent(
  platformOrInput: string | Record<string, unknown>,
  payload?: unknown,
  options?: PlatformCaptureOptions,
): PlatformCaptureRecord;

export function capturePlatformWebRequest(
  platformOrRequest: string | Request,
  requestOrOptions?: Request | PlatformCaptureOptions,
  maybeOptions?: PlatformCaptureOptions,
): Promise<PlatformCaptureRecord>;

export function buildPlatformCaptureSamples(
  records?: PlatformCaptureRecord[],
  options?: Record<string, unknown>,
): Record<string, Record<string, unknown>[]>;

export function buildPlatformCaptureAcceptanceReport(
  platform: string,
  records?: PlatformCaptureRecord[],
  options?: Record<string, unknown>,
): Record<string, unknown>;

export function buildMeetingPlatformCaptureAcceptanceSummary(
  records?: PlatformCaptureRecord[],
  options?: Record<string, unknown>,
): Record<string, unknown>;

export function serializePlatformCaptureRecord(record?: PlatformCaptureRecord): string;
export function parsePlatformCaptureRecord(line: string): PlatformCaptureRecord;
export function parsePlatformCaptureJsonl(text?: string): PlatformCaptureRecord[];
