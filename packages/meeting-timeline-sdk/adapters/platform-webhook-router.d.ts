export const MEETING_PLATFORM_ROUTE_SLUGS: Readonly<Record<string, string>>;

export interface PlatformWebhookRouterOptions {
  baseUrl?: string;
  basePath?: string;
  base_path?: string;
  env?: Record<string, unknown>;
  reconcile?: boolean;
  verify?: boolean;
  [key: string]: unknown;
}

export interface PlatformWebhookRouterRequest {
  platform?: string;
  provider?: string;
  adapter?: string;
  method?: string;
  url?: string;
  path?: string;
  headers?: Record<string, unknown>;
  query?: Record<string, unknown>;
  body?: unknown;
  rawBody?: string;
  raw_body?: string;
  [key: string]: unknown;
}

export function platformWebhookRoutePath(platform: string, options?: PlatformWebhookRouterOptions): string;
export function buildPlatformWebhookRouteTable(options?: PlatformWebhookRouterOptions): Record<string, unknown>[];
export function matchPlatformWebhookRoute(input?: PlatformWebhookRouterRequest, options?: PlatformWebhookRouterOptions): Record<string, unknown> | null;
export function buildPlatformWebhookRouterStatus(options?: PlatformWebhookRouterOptions): Record<string, unknown>;
export function buildPlatformWebhookRouterSetup(options?: PlatformWebhookRouterOptions): Record<string, unknown>;

export function createMeetingPlatformWebhookRouter(client: unknown, defaults?: PlatformWebhookRouterOptions): {
  (input?: PlatformWebhookRouterRequest, options?: PlatformWebhookRouterOptions): Promise<Record<string, unknown>>;
  match(input?: PlatformWebhookRouterRequest, options?: PlatformWebhookRouterOptions): Record<string, unknown> | null;
  routeTable(options?: PlatformWebhookRouterOptions): Record<string, unknown>[];
  status(options?: PlatformWebhookRouterOptions): Record<string, unknown>;
  setup(options?: PlatformWebhookRouterOptions): Record<string, unknown>;
  getReconciliationState(): unknown;
  resetReconciliationState(nextState?: Record<string, unknown>): unknown;
};
