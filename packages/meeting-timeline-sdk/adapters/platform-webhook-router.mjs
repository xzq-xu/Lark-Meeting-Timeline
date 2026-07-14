import { compactObject } from '../index.mjs';
import {
  createPlatformWebhookHandler,
} from './platform-webhook-handler.mjs';
import {
  MEETING_PLATFORM_KEYS,
  buildPlatformIntegrationPlan,
  buildPlatformPermissionPlan,
  buildPlatformSetup,
  normalizeMeetingPlatform,
  platformSetupManifest,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_ROUTE_SLUGS = Object.freeze({
  local_detector: 'local-detector',
  lark: 'lark',
  google_meet: 'google-meet',
  microsoft_teams: 'teams',
  zoom: 'zoom',
  webex: 'webex',
});

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function normalizeBasePath(value) {
  const raw = String(value || '/api/platform-events').trim();
  const withLeading = raw.startsWith('/') ? raw : `/${raw}`;
  return withLeading.replace(/\/+$/, '') || '/';
}

function requestPath(input = {}) {
  if (input.path) return String(input.path);
  if (input.url) {
    try {
      return new URL(String(input.url), 'http://localhost').pathname;
    } catch {
      return String(input.url).split('?')[0] || '/';
    }
  }
  return null;
}

function jsonResponse(status, body, headers = {}) {
  return {
    status,
    headers: {
      'content-type': 'application/json; charset=utf-8',
      ...headers,
    },
    body,
  };
}

function routePath(basePath, platform, suffix = '') {
  const key = normalizeMeetingPlatform(platform);
  const slug = MEETING_PLATFORM_ROUTE_SLUGS[key] ?? key;
  return `${normalizeBasePath(basePath)}/${slug}${suffix}`;
}

function routeMethods(platform) {
  return normalizeMeetingPlatform(platform) === 'microsoft_teams' ? ['GET', 'POST'] : ['POST'];
}

function mergedOptions(base = {}, next = {}) {
  return { ...base, ...next };
}

export function platformWebhookRoutePath(platform, options = {}) {
  return routePath(options.basePath ?? options.base_path, platform);
}

export function buildPlatformWebhookRouteTable(options = {}) {
  const basePath = normalizeBasePath(options.basePath ?? options.base_path);
  return MEETING_PLATFORM_KEYS.map((platform) => {
    const manifest = platformSetupManifest(platform, options);
    const permissionPlan = buildPlatformPermissionPlan(platform, options);
    const path = routePath(basePath, platform);
    return compactObject({
      platform,
      display_name: manifest.display_name,
      slug: MEETING_PLATFORM_ROUTE_SLUGS[platform],
      path,
      status_path: `${path}/status`,
      setup_path: `${path}/setup`,
      endpoint: manifest.endpoint,
      methods: routeMethods(platform),
      required_security_env: permissionPlan.required_security_env,
      missing_security_env: permissionPlan.missing_security_env,
      ready: permissionPlan.readiness?.ready,
    });
  });
}

export function matchPlatformWebhookRoute(input = {}, options = {}) {
  const explicitPlatform = firstNonEmpty(input.platform, input.provider, input.adapter);
  const path = requestPath(input);
  if (!path && explicitPlatform) {
    const platform = normalizeMeetingPlatform(explicitPlatform);
    return {
      matched: true,
      kind: 'webhook',
      platform,
      slug: MEETING_PLATFORM_ROUTE_SLUGS[platform],
      path: null,
    };
  }
  const basePath = normalizeBasePath(options.basePath ?? options.base_path);
  if (!path || (path !== basePath && !path.startsWith(`${basePath}/`))) return null;
  const rest = path.slice(basePath.length).replace(/^\/+/, '');
  if (!rest) {
    return { matched: true, kind: 'index', platform: null, slug: null, path };
  }
  if (rest === 'status') {
    return { matched: true, kind: 'status', platform: null, slug: null, path };
  }
  if (rest === 'setup') {
    return { matched: true, kind: 'setup', platform: null, slug: null, path };
  }
  const [slug, tail] = rest.split('/');
  let platform;
  try {
    platform = normalizeMeetingPlatform(slug);
  } catch {
    return null;
  }
  const kind = tail === 'status'
    ? 'status'
    : tail === 'setup'
      ? 'setup'
      : tail
        ? 'unknown'
        : 'webhook';
  return {
    matched: kind !== 'unknown',
    kind,
    platform,
    slug,
    path,
  };
}

export function buildPlatformWebhookRouterStatus(options = {}) {
  const routeTable = buildPlatformWebhookRouteTable(options);
  return {
    ok: routeTable.every((item) => item.ready !== false),
    base_path: normalizeBasePath(options.basePath ?? options.base_path),
    routes: routeTable,
    platforms: routeTable.map((route) => {
      const integrationPlan = buildPlatformIntegrationPlan(route.platform, options);
      const permissionPlan = buildPlatformPermissionPlan(route.platform, options);
      return compactObject({
        platform: route.platform,
        display_name: route.display_name,
        ready: permissionPlan.readiness?.ready,
        route: route.path,
        status_route: route.status_path,
        setup_route: route.setup_path,
        recommended_mode: integrationPlan.recommended_mode,
        provider_events_enabled: integrationPlan.provider_events?.enabled,
        required_security_env: permissionPlan.required_security_env,
        missing_security_env: permissionPlan.missing_security_env,
        readiness: permissionPlan.readiness,
      });
    }),
  };
}

export function buildPlatformWebhookRouterSetup(options = {}) {
  const routeTable = buildPlatformWebhookRouteTable(options);
  return {
    base_path: normalizeBasePath(options.basePath ?? options.base_path),
    routes: routeTable,
    setup: MEETING_PLATFORM_KEYS.map((platform) => buildPlatformSetup(platform, options)),
  };
}

export function createMeetingPlatformWebhookRouter(client, defaults = {}) {
  const handler = createPlatformWebhookHandler(client, defaults);

  function optionsFor(options = {}) {
    return mergedOptions(defaults, options);
  }

  async function handle(input = {}, options = {}) {
    const merged = optionsFor(options);
    const route = matchPlatformWebhookRoute(input, merged);
    if (!route || route.matched === false) {
      return jsonResponse(404, {
        error: 'platform_webhook_route_not_found',
        base_path: normalizeBasePath(merged.basePath ?? merged.base_path),
      });
    }
    if (route.kind === 'index' || (route.kind === 'status' && !route.platform)) {
      return jsonResponse(200, buildPlatformWebhookRouterStatus(merged));
    }
    if (route.kind === 'setup' && !route.platform) {
      return jsonResponse(200, buildPlatformWebhookRouterSetup(merged));
    }
    if (route.kind === 'status') {
      const status = buildPlatformWebhookRouterStatus(merged).platforms.find((item) => item.platform === route.platform);
      return jsonResponse(200, status);
    }
    if (route.kind === 'setup') {
      return jsonResponse(200, buildPlatformSetup(route.platform, merged));
    }
    return handler({
      ...input,
      platform: route.platform,
    }, options);
  }

  handle.match = (input = {}, options = {}) => matchPlatformWebhookRoute(input, optionsFor(options));
  handle.routeTable = (options = {}) => buildPlatformWebhookRouteTable(optionsFor(options));
  handle.status = (options = {}) => buildPlatformWebhookRouterStatus(optionsFor(options));
  handle.setup = (options = {}) => buildPlatformWebhookRouterSetup(optionsFor(options));
  handle.getReconciliationState = () => handler.getReconciliationState();
  handle.resetReconciliationState = (nextState = {}) => handler.resetReconciliationState(nextState);
  return handle;
}
