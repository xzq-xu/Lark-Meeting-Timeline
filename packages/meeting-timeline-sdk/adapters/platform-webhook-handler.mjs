import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  createReconciledPlatformEventIngestor,
  ingestPlatformEvent,
} from './platform-ingest.mjs';
import { meetingPlatformEventAdapterFor } from './platform-registry.mjs';
import {
  buildZoomUrlValidationResponse,
  microsoftGraphValidationResponse,
  platformWebhookVerificationStatus,
  verifyGooglePubSubBearer,
  verifyGooglePubSubOidcJwt,
  verifyMicrosoftGraphClientState,
  verifyWebexWebhookEvent,
  verifyZoomWebhookEvent,
} from './webhook-security.mjs';

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function optionGroup(options = {}, ...names) {
  for (const name of names) {
    const value = options[name];
    if (value && typeof value === 'object') return value;
  }
  return {};
}

function adapterFor(platform) {
  const adapter = meetingPlatformEventAdapterFor(platform);
  if (!adapter) {
    throw new MeetingTimelineSdkError(`Unsupported meeting platform: ${String(platform || '(empty)')}`, {
      platform,
    });
  }
  return adapter;
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

function textResponse(status, body, headers = {}) {
  return {
    status,
    headers: {
      'content-type': 'text/plain; charset=utf-8',
      ...headers,
    },
    body: String(body ?? ''),
  };
}

function requestSearchParams(input = {}) {
  if (input.searchParams && typeof input.searchParams.get === 'function') return input.searchParams;
  if (input.url) {
    try {
      return new URL(String(input.url), 'http://localhost').searchParams;
    } catch {
      // Fall through to query object support.
    }
  }
  const params = new URLSearchParams();
  for (const [key, value] of Object.entries(input.query ?? {})) {
    if (Array.isArray(value)) {
      for (const item of value) params.append(key, String(item));
    } else if (value != null) {
      params.set(key, String(value));
    }
  }
  return params;
}

function platformEventPayload(body = {}) {
  if (Array.isArray(body)) return body;
  if (body?.raw_event !== undefined) return body.raw_event;
  if (body?.rawEvent !== undefined) return body.rawEvent;
  if (body?.raw !== undefined) return body.raw;
  if (Array.isArray(body?.events)) return body.events;
  if (Array.isArray(body?.items)) return body.items;
  return body;
}

function publicAdapter(adapter = {}) {
  return {
    key: adapter.key,
    source: adapter.source,
    aliases: adapter.aliases,
  };
}

function publicIngestResult(result = {}, verification = null) {
  return compactObject({
    ok: true,
    platform: result.platform,
    source: result.source,
    adapter: publicAdapter(result.adapter),
    verification,
    raw_signal_count: result.rawSignals?.length,
    signal_count: result.signals?.length ?? 0,
    signals: result.signals,
    reconciliation: result.reconciliation,
    results: result.results,
  });
}

function reconcileEnabled(options = {}) {
  if (options.reconcile === false || options.reconciled === false) return false;
  return Boolean(
    options.reconcile === true
      || options.reconciled === true
      || options.reconciledIngestor
      || options.reconciled_ingestor
      || options.reconciler
      || options.signalReconciler
      || options.signal_reconciler,
  );
}

function reconciledIngestorFor(client, options = {}) {
  return options.reconciledIngestor
    ?? options.reconciled_ingestor
    ?? (reconcileEnabled(options) ? createReconciledPlatformEventIngestor(client, options) : null);
}

function googleVerificationConfigured(google = {}) {
  return Boolean(
    google.expectedAudience
      || google.audience
      || google.serviceAccountEmail
      || google.expectedEmail
      || google.jwks
      || google.publicKeys
      || google.certs
      || google.jwksUrl
      || process.env.GOOGLE_PUBSUB_OIDC_AUDIENCE
      || process.env.GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL,
  );
}

export async function verifyPlatformWebhook(platform, request = {}, options = {}) {
  const adapter = adapterFor(platform);
  const headers = request.headers ?? {};
  const body = request.body;
  const rawBody = request.rawBody ?? request.raw_body ?? '';

  if (adapter.key === 'zoom') {
    const zoom = optionGroup(options, 'zoom', 'zoomWebhook', 'zoom_webhook');
    return platformWebhookVerificationStatus(adapter.key, verifyZoomWebhookEvent({
      headers,
      rawBody,
      body,
      secretToken: firstNonEmpty(zoom.secretToken, zoom.secret_token, options.zoomSecretToken, options.secretToken),
      toleranceMs: firstNonEmpty(zoom.toleranceMs, zoom.tolerance_ms, options.toleranceMs),
    }));
  }

  if (adapter.key === 'webex') {
    const webex = optionGroup(options, 'webex', 'webexWebhook', 'webex_webhook');
    return platformWebhookVerificationStatus(adapter.key, verifyWebexWebhookEvent({
      headers,
      rawBody,
      body,
      secret: firstNonEmpty(webex.secret, webex.webhookSecret, webex.webhook_secret, options.webexSecret, options.secret),
    }));
  }

  if (adapter.key === 'microsoft_teams') {
    const graph = optionGroup(options, 'microsoftTeams', 'microsoft_teams', 'graph', 'microsoftGraph');
    return platformWebhookVerificationStatus(adapter.key, verifyMicrosoftGraphClientState(body, {
      clientState: firstNonEmpty(graph.clientState, graph.client_state, options.microsoftGraphClientState, options.clientState),
    }));
  }

  if (adapter.key === 'google_meet') {
    const google = optionGroup(options, 'googleMeet', 'google_meet', 'google', 'googlePubSub');
    if (google.preferBearer === true || google.prefer_bearer === true) {
      return platformWebhookVerificationStatus(adapter.key, verifyGooglePubSubBearer({
        headers,
        expectedToken: firstNonEmpty(google.bearerToken, google.bearer_token, options.googlePubSubBearerToken, options.bearerToken),
      }));
    }
    if (googleVerificationConfigured(google)) {
      return platformWebhookVerificationStatus(adapter.key, await verifyGooglePubSubOidcJwt({
        headers,
        token: firstNonEmpty(google.token, options.googlePubSubOidcToken),
        expectedAudience: firstNonEmpty(google.expectedAudience, google.expected_audience, options.googlePubSubOidcAudience),
        audience: firstNonEmpty(google.audience, options.googlePubSubAudience),
        serviceAccountEmail: firstNonEmpty(google.serviceAccountEmail, google.service_account_email, options.googlePubSubServiceAccountEmail),
        expectedEmail: firstNonEmpty(google.expectedEmail, google.expected_email),
        issuer: google.issuer,
        jwks: google.jwks,
        publicKeys: google.publicKeys ?? google.public_keys,
        certs: google.certs,
        jwksUrl: firstNonEmpty(google.jwksUrl, google.jwks_url, options.googlePubSubJwksUrl),
        fetchImpl: google.fetchImpl ?? google.fetch,
        nowMs: firstNonEmpty(google.nowMs, google.now_ms),
        clockToleranceSec: firstNonEmpty(google.clockToleranceSec, google.clock_tolerance_sec),
      }));
    }
    return platformWebhookVerificationStatus(adapter.key, verifyGooglePubSubBearer({
      headers,
      expectedToken: firstNonEmpty(google.bearerToken, google.bearer_token, options.googlePubSubBearerToken, options.bearerToken),
    }));
  }

  return { platform: adapter.key, ok: true, skipped: true, reason: 'platform_verification_not_configured' };
}

export async function handlePlatformWebhookRequest(client, input = {}, options = {}) {
  const platform = firstNonEmpty(input.platform, input.provider, input.adapter, options.platform);
  const adapter = adapterFor(platform);
  const method = String(input.method ?? 'POST').toUpperCase();
  const body = input.body ?? {};
  const searchParams = requestSearchParams(input);

  if (adapter.key === 'microsoft_teams' && method === 'GET' && searchParams.has('validationToken')) {
    return textResponse(200, microsoftGraphValidationResponse({ validationToken: searchParams.get('validationToken') }) ?? '');
  }

  if (method !== 'POST') {
    return jsonResponse(405, {
      error: 'method_not_allowed',
      platform: adapter.key,
      allowed_methods: adapter.key === 'microsoft_teams' ? ['GET', 'POST'] : ['POST'],
    });
  }

  try {
    if (adapter.key === 'zoom' && body?.event === 'endpoint.url_validation') {
      const zoom = optionGroup(options, 'zoom', 'zoomWebhook', 'zoom_webhook');
      return jsonResponse(200, buildZoomUrlValidationResponse(body.payload ?? body, {
        secretToken: firstNonEmpty(zoom.secretToken, zoom.secret_token, options.zoomSecretToken, options.secretToken),
      }));
    }

    const verification = options.verify === false
      ? { platform: adapter.key, ok: true, skipped: true, reason: 'verification_disabled' }
      : await verifyPlatformWebhook(adapter.key, input, options);
    if (verification?.ok === false) {
      return jsonResponse(401, {
        error: verification.reason ?? 'platform_webhook_verification_failed',
        platform: adapter.key,
        verification,
      });
    }

    const ingestInput = {
      platform: adapter.key,
      body: platformEventPayload(body),
      options: {
        receivedAtMs: firstNonEmpty(input.receivedAtMs, input.received_at_ms, input.received_at, options.receivedAtMs),
        normalizerOptions: options.normalizerOptions ?? options.normalizer_options,
        applyOptions: options.applyOptions ?? options.apply_options,
        reconcileOptions: options.reconcileOptions ?? options.reconcile_options,
      },
    };
    const reconciledIngestor = reconciledIngestorFor(client, options);
    const result = reconciledIngestor
      ? await reconciledIngestor.ingest(ingestInput)
      : await ingestPlatformEvent(client, ingestInput);
    return jsonResponse(200, publicIngestResult(result, verification));
  } catch (error) {
    return jsonResponse(error.status ?? 400, {
      error: error.message ?? String(error),
      platform: adapter.key,
      details: error.details,
    });
  }
}

export function createPlatformWebhookHandler(client, defaults = {}) {
  let sharedReconciledIngestor = defaults.reconciledIngestor ?? defaults.reconciled_ingestor;

  function ensureReconciledIngestor(options = {}) {
    if (!reconcileEnabled(options)) return null;
    if (options.reconciledIngestor || options.reconciled_ingestor) {
      sharedReconciledIngestor = options.reconciledIngestor ?? options.reconciled_ingestor;
      return sharedReconciledIngestor;
    }
    if (!sharedReconciledIngestor) {
      sharedReconciledIngestor = createReconciledPlatformEventIngestor(client, options);
    }
    return sharedReconciledIngestor;
  }

  async function handle(input = {}, options = {}) {
    const mergedOptions = {
      ...defaults,
      ...options,
    };
    const reconciledIngestor = ensureReconciledIngestor(mergedOptions);
    return handlePlatformWebhookRequest(client, {
      ...input,
      platform: firstNonEmpty(input.platform, defaults.platform),
    }, {
      ...mergedOptions,
      reconciledIngestor,
    });
  }

  handle.getReconciliationState = () => sharedReconciledIngestor?.getState?.() ?? null;
  handle.resetReconciliationState = (nextState = {}) => sharedReconciledIngestor?.reset?.(nextState) ?? null;
  return handle;
}
