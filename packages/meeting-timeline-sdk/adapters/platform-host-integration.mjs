import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  MEETING_PLATFORM_KEYS,
  buildPlatformIntegrationPlan,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';
import {
  buildMeetingPlatformLiveAdapterHandoffBundle,
} from './platform-live-adapter.mjs';
import {
  buildMeetingAppExtensionInstallPlan,
} from './meeting-app-extension.mjs';
import {
  buildMeetingAppRuntimeObserverPlanMatrix,
} from './meeting-app-profile.mjs';
import {
  buildMeetingPlatformRuntimeBundleMatrix,
} from './platform-runtime-bundle.mjs';
import {
  buildMeetingPlatformRuntimeEventPlanMatrix,
} from './platform-runtime-event.mjs';
import {
  buildMeetingPlatformParticipantTrackMatrix,
} from './platform-participant-track.mjs';
import {
  buildMeetingPlatformSpeakerTrackMatrix,
} from './platform-speaker-track.mjs';
import {
  buildMeetingPlatformAdaptationStrategyMatrix,
} from './platform-strategy.mjs';
import {
  buildMeetingPlatformAdapterRouteMatrix,
} from './platform-adapter-route.mjs';
import {
  buildMeetingPlatformAdapterBlueprintMatrix,
} from './platform-adapter-blueprint.mjs';
import {
  buildMeetingPlatformAdapterStartupPlanMatrix,
} from './platform-adapter-startup.mjs';
import {
  buildMeetingPlatformAdapterPreflightMatrix,
} from './platform-adapter-preflight.mjs';
import {
  buildMeetingPlatformAdapterContractAcceptanceMatrix,
  buildMeetingPlatformAdapterContractMatrix,
} from './platform-adapter-contract.mjs';
import {
  buildMeetingPlatformConformanceReport,
} from './platform-conformance.mjs';

export const MEETING_PLATFORM_HOST_INTEGRATION_SCHEMA = 'meeting_platform_host_integration';
export const MEETING_PLATFORM_HOST_INTEGRATION_SCHEMA_VERSION = 1;
export const MEETING_PLATFORM_HOST_INTEGRATION_SCAFFOLD_SCHEMA = 'meeting_platform_host_integration_scaffold';
export const MEETING_PLATFORM_HOST_INTEGRATION_ACCEPTANCE_SCHEMA = 'meeting_platform_host_integration_acceptance';

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

function unique(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function selectedPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_PLATFORM_KEYS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

function sourceFile(path, content, role, mime = 'text/plain') {
  return {
    path,
    role,
    mime,
    content: content.endsWith('\n') ? content : `${content}\n`,
  };
}

function json(value) {
  return JSON.stringify(value, null, 2);
}

function fileByPath(scaffold = {}, path) {
  return asArray(scaffold.files).find((file) => file?.path === path);
}

function issue(severity, code, message, details = {}) {
  return compactObject({ severity, code, message, ...details });
}

function platformImportNames(platforms = []) {
  return platforms.map((platform) => ({
    platform,
    normalized: normalizeMeetingPlatform(platform),
    variable: normalizeMeetingPlatform(platform).replace(/_([a-z])/g, (_, letter) => letter.toUpperCase()),
  }));
}

function pascalPlatformName(platform) {
  return normalizeMeetingPlatform(platform)
    .split('_')
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join('');
}

function adapterFilePath(platform) {
  return `src/platform-adapters/${normalizeMeetingPlatform(platform)}.mjs`;
}

function rowsByPlatform(matrix = {}) {
  return Object.fromEntries(asArray(matrix.rows).map((row) => [row.platform, row]));
}

function buildAdapterRuntimeContract(platforms = [], adapterRouteMatrix = {}, adapterBlueprintMatrix = {}) {
  const routeRows = rowsByPlatform(adapterRouteMatrix);
  const blueprintRows = rowsByPlatform(adapterBlueprintMatrix);
  const rows = platforms.map((platform) => {
    const route = routeRows[platform] ?? {};
    const blueprint = blueprintRows[platform] ?? {};
    const ready = route.route_ready === true
      && blueprint.ready === true
      && route.provider_blocks_realtime !== true
      && route.transcript_blocks_realtime !== true
      && blueprint.provider_blocks_realtime !== true
      && blueprint.transcript_blocks_realtime !== true;
    const missingItems = [
      route.route_ready === true ? null : 'adapter_route_not_ready',
      blueprint.ready === true ? null : 'adapter_blueprint_not_ready',
      route.provider_blocks_realtime === true || blueprint.provider_blocks_realtime === true ? 'provider_blocks_realtime' : null,
      route.transcript_blocks_realtime === true || blueprint.transcript_blocks_realtime === true ? 'transcript_blocks_realtime' : null,
    ].filter(Boolean);
    return compactObject({
      platform,
      ready,
      adapter_file: adapterFilePath(platform),
      selected_surface: blueprint.primary_surface ?? route.primary_surface,
      first_route: route.first_route,
      surface_order: blueprint.surface_order ?? route.surface_order,
      observe_candidates_method: 'observeCandidates',
      insert_annotation_method: 'insertAnnotation',
      speaker_track_method: 'speakerTrack',
      participant_track_method: 'participantTrack',
      provider_reconcile_method: 'ingestProvider',
      timestamp_field: 'captured_at_ms',
      provider_blocks_realtime: route.provider_blocks_realtime === true || blueprint.provider_blocks_realtime === true,
      transcript_blocks_realtime: route.transcript_blocks_realtime === true || blueprint.transcript_blocks_realtime === true,
      missing_items: missingItems,
    });
  });
  const readyCount = rows.filter((row) => row.ready === true).length;
  return {
    type: 'meeting_platform_adapter_runtime_contract',
    platform_count: platforms.length,
    ready_count: readyCount,
    missing_count: platforms.length - readyCount,
    all_ready: readyCount === platforms.length,
    index_file: 'src/platform-adapters/index.mjs',
    runtime_rule: 'per_platform_adapter_entries_call_observe_candidates_before_insert_annotation_and_preserve_captured_at_ms',
    acceptance_gate: 'all_selected_platforms_must_have_adapter_runtime_entrypoints',
    rows,
  };
}

function timelineClientSource(options = {}) {
  const defaultBaseUrl = firstNonEmpty(options.baseUrl, options.base_url, 'http://localhost:8787');
  return `import { createMeetingTimelineClient } from '@ai-annotation/meeting-timeline-sdk';

export function createHostTimelineClient(options = {}) {
  return createMeetingTimelineClient({
    baseUrl: options.baseUrl ?? options.base_url ?? ${JSON.stringify(defaultBaseUrl)},
    fetch: options.fetch ?? globalThis.fetch,
    headers: options.headers,
  });
}
`;
}

function hostSource(platforms = [], options = {}) {
  const defaultPlatforms = JSON.stringify(platforms);
  const defaultBaseUrl = firstNonEmpty(options.baseUrl, options.base_url, 'http://localhost:8787');
  return `import { createMeetingPlatformIntegrationRuntime } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime';
import { createHostTimelineClient } from './timeline-client.mjs';
import { createMeetingPlatformAdapter, createMeetingPlatformAdapters } from './platform-adapters/index.mjs';

const DEFAULT_PLATFORMS = ${defaultPlatforms};

export function createMeetingPlatformHost(options = {}) {
  const client = options.client ?? createHostTimelineClient({
    baseUrl: options.baseUrl ?? options.base_url ?? ${JSON.stringify(defaultBaseUrl)},
    fetch: options.fetch,
    headers: options.headers,
  });
  const platforms = options.platforms ?? options.platform_keys ?? DEFAULT_PLATFORMS;
  const integrationRuntime = createMeetingPlatformIntegrationRuntime(client, {
    ...options,
    platforms,
  });
  const kit = integrationRuntime.kit;
  const liveAdapters = integrationRuntime.liveAdapters;

  function extensionPlatforms(value = platforms) {
    const rows = Array.isArray(value) ? value : [value].filter(Boolean);
    return rows.filter((platform) => platform !== 'local_detector');
  }

  const host = {
    client,
    integrationRuntime,
    integration_runtime: integrationRuntime,
    kit,
    platforms,
    liveAdapters,
    handoffBundle(handoffOptions = {}) {
      return kit.platformLiveAdapterHandoffBundle({
        ...handoffOptions,
        platforms: handoffOptions.platforms ?? handoffOptions.platform_keys ?? platforms,
      });
    },
    adapterContracts(contractOptions = {}) {
      return kit.platformAdapterContractMatrix({
        ...contractOptions,
        platforms: contractOptions.platforms ?? contractOptions.platform_keys ?? platforms,
      });
    },
    adapterContractAcceptance(acceptanceOptions = {}) {
      return kit.platformAdapterContractAcceptanceMatrix({
        ...acceptanceOptions,
        platforms: acceptanceOptions.platforms ?? acceptanceOptions.platform_keys ?? platforms,
      });
    },
    platformConformance(conformanceOptions = {}) {
      return integrationRuntime.conformance({
        ...conformanceOptions,
        platforms: conformanceOptions.platforms ?? conformanceOptions.platform_keys ?? platforms,
      });
    },
    assertPlatformConformance(conformanceOptions = {}) {
      return integrationRuntime.assertConformance({
        ...conformanceOptions,
        platforms: conformanceOptions.platforms ?? conformanceOptions.platform_keys ?? platforms,
      });
    },
    consumerHandoff(handoffOptions = {}) {
      return kit.platformConsumerHandoff({
        ...handoffOptions,
        platforms: handoffOptions.platforms ?? handoffOptions.platform_keys ?? platforms,
      });
    },
    assertConsumerHandoff(handoffOptions = {}) {
      return kit.assertPlatformConsumerHandoff({
        ...handoffOptions,
        platforms: handoffOptions.platforms ?? handoffOptions.platform_keys ?? platforms,
      });
    },
    runtimeBundles(bundleOptions = {}) {
      return kit.platformRuntimeBundleMatrix({
        ...bundleOptions,
        platforms: bundleOptions.platforms ?? bundleOptions.platform_keys ?? platforms,
      });
    },
    observerPlans(planOptions = {}) {
      return kit.meetingAppRuntimeObserverPlanMatrix({
        ...planOptions,
        platforms: planOptions.platforms ?? planOptions.platform_keys ?? extensionPlatforms(platforms),
      });
    },
    runtimeEventPlans(planOptions = {}) {
      return kit.platformRuntimeEventPlanMatrix({
        ...planOptions,
        platforms: planOptions.platforms ?? planOptions.platform_keys ?? platforms,
      });
    },
    adapterRoutes(routeOptions = {}) {
      return kit.platformAdapterRouteMatrix({
        ...routeOptions,
        platforms: routeOptions.platforms ?? routeOptions.platform_keys ?? platforms,
      });
    },
    adapterBlueprints(blueprintOptions = {}) {
      return kit.platformAdapterBlueprintMatrix({
        ...blueprintOptions,
        platforms: blueprintOptions.platforms ?? blueprintOptions.platform_keys ?? platforms,
      });
    },
    adapterStartupPlans(startupOptions = {}) {
      return kit.platformAdapterStartupPlanMatrix(startupOptions.input ?? {}, {
        ...startupOptions,
        platforms: startupOptions.platforms ?? startupOptions.platform_keys ?? platforms,
      });
    },
    adapterPreflight(input = {}, preflightOptions = {}) {
      return kit.platformAdapterPreflight(input, preflightOptions);
    },
    adapterPreflightMatrix(input = {}, preflightOptions = {}) {
      return kit.platformAdapterPreflightMatrix(input, {
        ...preflightOptions,
        platforms: preflightOptions.platforms ?? preflightOptions.platform_keys ?? platforms,
      });
    },
    adapterCandidatePreflight(input = {}, preflightOptions = {}) {
      return kit.platformAdapterCandidatePreflight(input, {
        ...preflightOptions,
        platforms: preflightOptions.platforms ?? preflightOptions.platform_keys ?? platforms,
      });
    },
    adaptationStrategyMatrix(strategyOptions = {}) {
      return kit.platformAdaptationStrategyMatrix({
        ...strategyOptions,
        platforms: strategyOptions.platforms ?? strategyOptions.platform_keys ?? platforms,
      });
    },
    resolvePlatform(input = {}, resolveOptions = {}) {
      return integrationRuntime.resolvePlatform(input, {
        ...resolveOptions,
        platforms: resolveOptions.platforms ?? resolveOptions.platform_keys ?? platforms,
      });
    },
    resolvePlatformCandidates(input = {}, resolveOptions = {}) {
      return integrationRuntime.resolvePlatformCandidates(input, {
        ...resolveOptions,
        platforms: resolveOptions.platforms ?? resolveOptions.platform_keys ?? platforms,
      });
    },
    observePlatformCandidates(input = {}, observeOptions = {}) {
      return integrationRuntime.observePlatformCandidates(input, {
        ...observeOptions,
        platforms: observeOptions.platforms ?? observeOptions.platform_keys ?? platforms,
      });
    },
    runtimeBundle(platform, bundleOptions = {}) {
      return kit.platformRuntimeBundle(platform, bundleOptions);
    },
    extensionInstallPlan(extensionOptions = {}) {
      return kit.meetingAppExtensionInstallPlan({
        ...extensionOptions,
        baseUrl: extensionOptions.baseUrl ?? extensionOptions.base_url ?? ${JSON.stringify(defaultBaseUrl)},
        platforms: extensionOptions.platforms ?? extensionOptions.platform_keys ?? extensionPlatforms(platforms),
      });
    },
    integrationRuntimeManifest(manifestOptions = {}) {
      return integrationRuntime.manifest({
        ...manifestOptions,
        platforms: manifestOptions.platforms ?? manifestOptions.platform_keys ?? platforms,
      });
    },
    runIntegrationRuntimeManifest(manifestOptions = {}) {
      return integrationRuntime.runManifest({
        ...manifestOptions,
        platforms: manifestOptions.platforms ?? manifestOptions.platform_keys ?? platforms,
      });
    },
    assertRunIntegrationRuntimeManifest(manifestOptions = {}) {
      return integrationRuntime.runAndAssertManifest({
        ...manifestOptions,
        platforms: manifestOptions.platforms ?? manifestOptions.platform_keys ?? platforms,
      });
    },
    integrationRuntimeSummary(summaryOptions = {}) {
      return integrationRuntime.summary({
        ...summaryOptions,
        platforms: summaryOptions.platforms ?? summaryOptions.platform_keys ?? platforms,
      });
    },
    handoffReadiness(readinessOptions = {}) {
      return integrationRuntime.handoffReadiness({
        ...readinessOptions,
        platforms: readinessOptions.platforms ?? readinessOptions.platform_keys ?? platforms,
      });
    },
    runHandoffReadiness(readinessOptions = {}) {
      return integrationRuntime.runHandoffReadiness({
        ...readinessOptions,
        platforms: readinessOptions.platforms ?? readinessOptions.platform_keys ?? platforms,
      });
    },
    handleRuntimeEvent(input = {}, payload, eventOptions = {}) {
      return integrationRuntime.handleEvent(input, payload, eventOptions);
    },
    adapter(platform, adapterOptions = {}) {
      return integrationRuntime.adapter(platform, adapterOptions);
    },
    observeMeetingApp(platform, snapshot = {}, observeOptions = {}) {
      return integrationRuntime.observeMeetingApp(platform, snapshot, observeOptions);
    },
    ingestProvider(platform, requestOrPayload = {}, payload, ingestOptions = {}) {
      return integrationRuntime.ingestProvider(platform, requestOrPayload, payload, ingestOptions);
    },
    insertAnnotation(platform, input = {}, markOptions = {}) {
      return integrationRuntime.insertAnnotation(platform, {
        ...input,
        platform,
        capturedAtMs: input.capturedAtMs ?? input.captured_at_ms ?? Date.now(),
      }, markOptions);
    },
    speakerTrack(platform, input = {}, trackOptions = {}) {
      return integrationRuntime.speakerTrack(platform, input, trackOptions);
    },
    participantTrack(platform, input = {}, trackOptions = {}) {
      return integrationRuntime.participantTrack(platform, input, trackOptions);
    },
    platformAdapters(adapterOptions = {}) {
      return createMeetingPlatformAdapters(host, {
        ...options,
        ...adapterOptions,
      });
    },
    platformAdapter(platform, adapterOptions = {}) {
      return createMeetingPlatformAdapter(host, platform, {
        ...options,
        ...adapterOptions,
      });
    },
    verifyReadiness(readinessOptions = {}) {
      return liveAdapters.assertReadinessMatrix({
        ...readinessOptions,
        platforms: readinessOptions.platforms ?? readinessOptions.platform_keys ?? platforms,
      });
    },
  };
  return host;
}
`;
}

function routesSource(options = {}) {
  const defaultBasePath = firstNonEmpty(options.basePath, options.base_path, '/api/platform-events');
  return `export async function handleMeetingPlatformRequest(host, request, options = {}) {
  const basePath = options.basePath ?? options.base_path ?? ${JSON.stringify(defaultBasePath)};
  const url = new URL(request.url);
  if (url.pathname.startsWith(basePath)) {
    return host.kit.handleFetchRequest(request, options);
  }
  if (url.pathname === '/api/annotations' && request.method === 'POST') {
    const payload = await request.json();
    const result = await host.insertAnnotation(payload.platform, payload);
    return Response.json(result, { status: 201 });
  }
  if (url.pathname === '/api/meeting-platform/readiness') {
    return Response.json(host.liveAdapters.readinessMatrix(options));
  }
  if (url.pathname === '/api/meeting-platform/handoff') {
    return Response.json(host.handoffBundle(options));
  }
  if (url.pathname === '/api/meeting-platform/contracts') {
    return Response.json(host.adapterContracts(options));
  }
  if (url.pathname === '/api/meeting-platform/contract-acceptance') {
    return Response.json(host.adapterContractAcceptance(options));
  }
  if (url.pathname === '/api/meeting-platform/conformance') {
    return Response.json(host.platformConformance(options));
  }
  if (url.pathname === '/api/meeting-platform/consumer-handoff') {
    return Response.json(host.consumerHandoff(options));
  }
  if (url.pathname === '/api/meeting-platform/runtime-bundles') {
    return Response.json(host.runtimeBundles(options));
  }
  if (url.pathname === '/api/meeting-platform/observer-plans') {
    return Response.json(host.observerPlans(options));
  }
  if (url.pathname === '/api/meeting-platform/runtime-event-plans') {
    return Response.json(host.runtimeEventPlans(options));
  }
  if (url.pathname === '/api/meeting-platform/adapter-routes') {
    return Response.json(host.adapterRoutes(options));
  }
  if (url.pathname === '/api/meeting-platform/adapter-blueprints') {
    return Response.json(host.adapterBlueprints(options));
  }
  if (url.pathname === '/api/meeting-platform/adapter-startup') {
    return Response.json(host.adapterStartupPlans(options));
  }
  if (url.pathname === '/api/meeting-platform/adapter-preflight') {
    const payload = request.method === 'POST' ? await request.json() : Object.fromEntries(url.searchParams.entries());
    const mode = payload.mode ?? options.mode;
    const input = payload.input ?? payload;
    if (mode === 'matrix') return Response.json(host.adapterPreflightMatrix(input, options));
    if (mode === 'candidates' || mode === 'candidate') return Response.json(host.adapterCandidatePreflight(input, options));
    return Response.json(host.adapterPreflight(input, options));
  }
  if (url.pathname === '/api/meeting-platform/strategy') {
    return Response.json(host.adaptationStrategyMatrix(options));
  }
  if (url.pathname === '/api/meeting-platform/resolve') {
    const payload = request.method === 'POST' ? await request.json() : {};
    const input = {
      ...Object.fromEntries(url.searchParams.entries()),
      ...payload,
    };
    return Response.json(host.resolvePlatform(input, options));
  }
  if (url.pathname === '/api/meeting-platform/resolve-candidates' && request.method === 'POST') {
    const payload = await request.json();
    return Response.json(host.resolvePlatformCandidates(payload, options));
  }
  if (url.pathname === '/api/meeting-platform/observe-candidates' && request.method === 'POST') {
    const payload = await request.json();
    const result = await host.observePlatformCandidates(payload, options);
    return Response.json(result);
  }
  if (url.pathname === '/api/meeting-platform/extension-plan') {
    return Response.json(host.extensionInstallPlan(options));
  }
  if (url.pathname === '/api/meeting-platform/integration-runtime') {
    return Response.json(host.integrationRuntimeSummary(options));
  }
  if (url.pathname === '/api/meeting-platform/integration-runtime/manifest') {
    return Response.json(host.integrationRuntimeManifest(options));
  }
  if (url.pathname === '/api/meeting-platform/integration-runtime/run-manifest') {
    const payload = request.method === 'POST' ? await request.json() : {};
    const result = await host.runIntegrationRuntimeManifest({ ...options, ...payload });
    return Response.json(result);
  }
  if (url.pathname === '/api/meeting-platform/handoff-readiness') {
    const payload = request.method === 'POST' ? await request.json() : {};
    const result = await host.runHandoffReadiness({ ...options, ...payload });
    return Response.json(result);
  }
  if (url.pathname === '/api/meeting-platform/runtime-events' && request.method === 'POST') {
    const payload = await request.json();
    const result = await host.handleRuntimeEvent(payload, payload.payload ?? payload.raw, options);
    return Response.json(result);
  }
  return Response.json({ ok: false, error: 'not_found' }, { status: 404 });
}
`;
}

function platformAdapterSource(platform, plan = {}) {
  const normalized = normalizeMeetingPlatform(platform);
  const pascal = pascalPlatformName(normalized);
  const route = rowsByPlatform(plan.adapter_route_matrix)[normalized] ?? {};
  const blueprint = rowsByPlatform(plan.adapter_blueprint_matrix)[normalized] ?? {};
  const contract = rowsByPlatform(plan.adapter_runtime_contract)[normalized] ?? {};
  return `const PLATFORM = ${JSON.stringify(normalized)};
const ADAPTER_ROUTE = ${json(route)};
const ADAPTER_BLUEPRINT = ${json(blueprint)};
const ADAPTER_RUNTIME_CONTRACT = ${json(contract)};

function mergedOptions(defaults = {}, overrides = {}) {
  return { ...defaults, ...overrides };
}

function withPlatform(input = {}) {
  return { ...input, platform: input.platform ?? PLATFORM };
}

function capturedAtMs(input = {}) {
  const value = input.capturedAtMs ?? input.captured_at_ms;
  if (value == null) {
    throw new Error(\`\${PLATFORM} annotation requires captured_at_ms\`);
  }
  return value;
}

export function create${pascal}TimelineAdapter(host, options = {}) {
  if (!host) throw new Error('host is required for meeting platform adapter');
  return {
    platform: PLATFORM,
    display_name: ADAPTER_ROUTE.display_name ?? ADAPTER_BLUEPRINT.display_name,
    selected_surface: ADAPTER_BLUEPRINT.primary_surface ?? ADAPTER_ROUTE.primary_surface,
    first_route: ADAPTER_ROUTE.first_route,
    route: ADAPTER_ROUTE,
    blueprint: ADAPTER_BLUEPRINT,
    runtime_contract: ADAPTER_RUNTIME_CONTRACT,
    resolve(input = {}, adapterOptions = {}) {
      return host.resolvePlatform(withPlatform(input), mergedOptions(options, adapterOptions));
    },
    preflight(input = {}, adapterOptions = {}) {
      return host.resolvePlatform(withPlatform(input), mergedOptions(options, adapterOptions));
    },
    resolveCandidates(input = {}, adapterOptions = {}) {
      return host.resolvePlatformCandidates(withPlatform(input), mergedOptions(options, adapterOptions));
    },
    observeCandidates(input = {}, adapterOptions = {}) {
      return host.observePlatformCandidates(withPlatform(input), mergedOptions(options, adapterOptions));
    },
    insertAnnotation(input = {}, adapterOptions = {}) {
      return host.insertAnnotation(PLATFORM, {
        ...input,
        platform: PLATFORM,
        capturedAtMs: capturedAtMs(input),
      }, mergedOptions(options, adapterOptions));
    },
    speakerTrack(input = {}, adapterOptions = {}) {
      return host.speakerTrack(PLATFORM, withPlatform(input), mergedOptions(options, adapterOptions));
    },
    participantTrack(input = {}, adapterOptions = {}) {
      return host.participantTrack(PLATFORM, withPlatform(input), mergedOptions(options, adapterOptions));
    },
    ingestProvider(requestOrPayload = {}, payload, adapterOptions = {}) {
      return host.ingestProvider(PLATFORM, requestOrPayload, payload, mergedOptions(options, adapterOptions));
    },
  };
}

export default create${pascal}TimelineAdapter;
`;
}

function platformAdaptersIndexSource(platforms = []) {
  const imports = platforms.map((platform) => {
    const normalized = normalizeMeetingPlatform(platform);
    const pascal = pascalPlatformName(normalized);
    return `import { create${pascal}TimelineAdapter } from './${normalized}.mjs';`;
  });
  const factoryEntries = platforms.map((platform) => {
    const normalized = normalizeMeetingPlatform(platform);
    const pascal = pascalPlatformName(normalized);
    return `  ${JSON.stringify(normalized)}: create${pascal}TimelineAdapter,`;
  });
  return `${imports.join('\n')}

const FACTORIES = {
${factoryEntries.join('\n')}
};

function normalizePlatform(value) {
  return String(value ?? '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}

export function createMeetingPlatformAdapter(host, platform, options = {}) {
  const key = normalizePlatform(platform);
  const factory = FACTORIES[key];
  if (!factory) throw new Error(\`Unsupported meeting platform adapter: \${platform}\`);
  return factory(host, options);
}

export function createMeetingPlatformAdapters(host, options = {}) {
  return Object.fromEntries(
    Object.entries(FACTORIES).map(([platform, factory]) => [platform, factory(host, options)]),
  );
}

export const meetingPlatformAdapterFactories = FACTORIES;
`;
}

function readmeSource(plan = {}) {
  const platforms = plan.platforms ?? [];
  const candidateRows = plan.candidate_observation_contract?.rows ?? [];
  const trackRows = plan.meeting_track_contract?.rows ?? [];
  const adapterRows = plan.adapter_runtime_contract?.rows ?? [];
  const startupRows = plan.adapter_startup_plan_matrix?.rows ?? [];
  const preflightRows = plan.adapter_preflight_matrix?.rows ?? [];
  return `# Meeting Platform Timeline Host

This scaffold wires a host project to @ai-annotation/meeting-timeline-sdk.

Runtime rule:

- Realtime annotation marks use captured_at_ms.
- Local meeting app observation may create the meeting axis before provider events arrive.
- Provider webhooks reconcile or backfill the axis; they do not block realtime annotation.
- Transcript and artifact import are post-meeting work.
- Each meeting keeps an isolated annotation set.
- Candidate observation is a required host-axis binding contract: browser extensions or native hosts send active meeting windows through meeting_timeline.observe_candidates or /api/meeting-platform/observe-candidates before realtime marks are inserted.
- Observer plans standardize the local DOM/AX observation loop, throttling, speaker follow-up, and meeting-end grace windows for each meeting app surface.
- Adapter routes describe each platform's implementation path: local observer realtime axis first, provider reconciliation second, transcript/artifact import last.
- Adapter blueprints expose the concrete browser/native/provider/artifact surface contract each host must wire for Google Meet, Teams, Zoom, Webex, Lark, and local detector.
- Adapter startup plans tell the host which browser/native/webview surface to launch first and which observe/insert actions to call before realtime marks are accepted.
- Adapter preflight validates current live DOM/native evidence before the host lets realtime annotations bind to a meeting axis.
- Per-platform adapter runtime entries live under src/platform-adapters/*.mjs. They provide resolve, observeCandidates, insertAnnotation, speakerTrack, participantTrack, and ingestProvider methods while preserving captured_at_ms.
- Speaker and participant position tracks are required realtime contracts. They use local samples/snapshots first and must not wait for provider events or transcript export.
- Platform conformance is the static SDK handoff gate before real evidence replay: normalizer, setup, contract, route, runtime bundle, registry, candidate observation, captured_at_ms, and provider/transcript non-blocking rules must all pass.

Supported platforms in this scaffold:

${platforms.map((platform) => `- ${platform}`).join('\n')}

Candidate observation contract:

- endpoint: ${plan.candidate_observation_contract?.endpoint ?? '/api/meeting-platform/observe-candidates'}
- runtime event action: ${plan.candidate_observation_contract?.runtime_event_action ?? 'observe_platform_candidates'}
- required permissions: ${(plan.candidate_observation_contract?.required_permissions ?? ['tabs']).join(', ')}
- all ready: ${plan.candidate_observation_contract?.all_ready === true ? 'true' : 'false'}

${candidateRows.map((row) => `- ${row.platform}: ${row.ready ? 'ready' : 'missing'} via ${row.message_type ?? 'meeting_timeline.observe_candidates'} (${row.required_permission ?? 'tabs'})`).join('\n')}

Observer plan contract:

- endpoint: ${plan.endpoints?.observer_plans ?? '/api/meeting-platform/observer-plans'}
- platform count: ${plan.observer_plan_matrix?.platform_count ?? 0}
- sdk ready count: ${plan.observer_plan_matrix?.sdk_ready_count ?? 0}
- preflight accepted count: ${plan.observer_plan_matrix?.preflight_accepted_count ?? 0}
- timestamp field: captured_at_ms

Meeting track contract:

- all ready: ${plan.meeting_track_contract?.all_ready === true ? 'true' : 'false'}
- speaker ready count: ${plan.meeting_track_contract?.speaker_ready_count ?? 0}
- participant ready count: ${plan.meeting_track_contract?.participant_ready_count ?? 0}
- provider blocking count: ${plan.meeting_track_contract?.provider_blocking_count ?? 0}
- transcript blocking count: ${plan.meeting_track_contract?.transcript_blocking_count ?? 0}

${trackRows.map((row) => `- ${row.platform}: speaker=${row.speaker_track_ready ? 'ready' : 'missing'}, participant=${row.participant_track_ready ? 'ready' : 'missing'}`).join('\n')}

Adapter runtime entries:

- index file: ${plan.adapter_runtime_contract?.index_file ?? 'src/platform-adapters/index.mjs'}
- all ready: ${plan.adapter_runtime_contract?.all_ready === true ? 'true' : 'false'}
- ready count: ${plan.adapter_runtime_contract?.ready_count ?? 0}
- timestamp field: captured_at_ms

${adapterRows.map((row) => `- ${row.platform}: ${row.ready ? 'ready' : 'missing'} via ${row.adapter_file} (${row.selected_surface ?? 'unknown_surface'})`).join('\n')}

Adapter startup plans:

- endpoint: ${plan.endpoints?.adapter_startup ?? '/api/meeting-platform/adapter-startup'}
- script: npm run meeting-platform:adapter-startup
- realtime ready count: ${plan.adapter_startup_plan_matrix?.realtime_startup_ready_count ?? 0}
- provider reconcile only count: ${plan.adapter_startup_plan_matrix?.provider_reconcile_surface_count ?? 0}

${startupRows.map((row) => `- ${row.platform}: surface=${row.selected_surface ?? 'unknown_surface'}, install=${row.install_target ?? 'unknown_target'}, observe=${row.observe_action ?? 'observePlatformCandidates'}, insert=${row.insert_action ?? 'insertAnnotation'}`).join('\n')}

Adapter preflight:

- endpoint: ${plan.endpoints?.adapter_preflight ?? '/api/meeting-platform/adapter-preflight'}
- script: npm run meeting-platform:adapter-preflight
- platform count: ${plan.adapter_preflight_matrix?.platform_count ?? 0}
- accepted count without live evidence: ${plan.adapter_preflight_matrix?.accepted_count ?? 0}
- realtime ready count without live evidence: ${plan.adapter_preflight_matrix?.realtime_ready_count ?? 0}

${preflightRows.map((row) => `- ${row.platform}: status=${row.status ?? 'unknown_status'}, surface=${row.selected_surface ?? 'unknown_surface'}, startup=${row.startup_ready ? 'ready' : 'missing'}, live=${row.live_evidence_ready ? 'ready' : 'missing'}`).join('\n')}

Platform conformance:

- accepted: ${plan.platform_conformance_report?.accepted === true ? 'true' : 'false'}
- accepted count: ${plan.platform_conformance_report?.accepted_count ?? 0}
- blocking count: ${plan.platform_conformance_report?.blocking_count ?? 0}

Consumer handoff:

- endpoint: ${plan.endpoints?.consumer_handoff ?? '/api/meeting-platform/consumer-handoff'}
- script: npm run meeting-platform:consumer-handoff
- purpose: one machine-readable SDK entrypoint index for downstream projects that need to plug marks into a meeting timeline across Google Meet, Teams, Zoom, Webex, and Lark.

Minimal usage:

\`\`\`js
import { createMeetingPlatformHost } from './src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: 'https://timeline.example.com',
});

await host.insertAnnotation('google_meet', {
  label: 'why?',
  capturedAtMs: Date.now(),
});

const handoff = host.handoffBundle();
const consumerHandoff = host.consumerHandoff();
const conformance = host.platformConformance();
const strategy = host.adaptationStrategyMatrix();
const platformResolution = host.resolvePlatform({
  url: 'https://meet.google.com/abc-defg-hij',
});
const candidateResolution = host.resolvePlatformCandidates({
  windows: [{
    tabs: [{ url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet', active: true }],
  }],
});
const candidateObservation = await host.observePlatformCandidates({
  windows: [{
    tabs: [{ url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet', active: true }],
  }],
});
const runtimeBundles = host.runtimeBundles();
const observerPlans = host.observerPlans();
const runtimeEventPlans = host.runtimeEventPlans();
const adapterRoutes = host.adapterRoutes();
const adapterBlueprints = host.adapterBlueprints();
const adapterStartupPlans = host.adapterStartupPlans();
const adapterPreflight = host.adapterPreflight({
  platform: 'google-meet',
  url: 'https://meet.google.com/abc-defg-hij',
});
const googleMeetAdapter = host.platformAdapter('google_meet');
await googleMeetAdapter.observeCandidates({
  windows: [{
    tabs: [{ url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet', active: true }],
  }],
});
await googleMeetAdapter.insertAnnotation({
  label: 'why?',
  captured_at_ms: Date.now(),
});
const extensionPlan = host.extensionInstallPlan();
const integrationRuntime = host.integrationRuntimeSummary();
const runtimeHandoffGate = await host.runIntegrationRuntimeManifest({
  requireHandoffReady: process.env.MEETING_PLATFORM_REQUIRE_HANDOFF_READY === '1',
});
const candidateContract = host.integrationRuntimeManifest().rows.map((row) => ({
  platform: row.platform,
  candidate_observation_ready: row.candidate_observation_ready,
  speaker_track_ready: row.speaker_track_ready,
  participant_track_ready: row.participant_track_ready,
}));
\`\`\`

Verification:

\`\`\`sh
npm run meeting-platform:strategy
npm run meeting-app:extension:build
npm run meeting-app:evidence-matrix
npm run meeting-platform:rollout-matrix
npm run meeting-platform:live-readiness
npm run meeting-platform:contracts
npm run meeting-platform:contract-acceptance
npm run meeting-platform:conformance
npm run meeting-platform:consumer-handoff
npm run meeting-platform:runtime-bundles
npm run meeting-platform:observer-plans
npm run meeting-platform:runtime-event-plans
npm run meeting-platform:resolve
npm run meeting-platform:adapter-blueprints
npm run meeting-platform:adapter-startup
npm run meeting-platform:adapter-preflight
npm run meeting-platform:resolve-candidates
npm run meeting-platform:observe-candidates
npm run meeting-platform:extension-plan
npm run meeting-platform:integration-runtime
npm run meeting-platform:integration-runtime-manifest
npm run meeting-platform:integration-runtime-run-manifest
npm run meeting-platform:handoff-readiness
npm run sdk:package-smoke
\`\`\`
`;
}

function buildCandidateObservationContract(platforms = [], runtimeBundleMatrix = {}) {
  const rowsByPlatform = Object.fromEntries(asArray(runtimeBundleMatrix.rows).map((row) => [row.platform, row]));
  const rows = platforms.map((platform) => {
    const row = rowsByPlatform[platform] ?? {};
    const ready = row.candidate_observation_ready === true;
    return compactObject({
      platform,
      ready,
      message_type: row.candidate_observer_message_type ?? 'meeting_timeline.observe_candidates',
      required_permission: row.candidate_observer_permission ?? 'tabs',
      endpoint: '/api/meeting-platform/observe-candidates',
      runtime_event_action: 'observe_platform_candidates',
      producer: 'browser_extension_background_or_native_host',
      missing_items: ready ? [] : ['candidate_observation_runtime_bundle'],
    });
  });
  const readyCount = rows.filter((row) => row.ready === true).length;
  return {
    type: 'meeting_platform_candidate_observation_contract',
    platform_count: platforms.length,
    ready_count: readyCount,
    missing_count: platforms.length - readyCount,
    all_ready: readyCount === platforms.length,
    endpoint: '/api/meeting-platform/observe-candidates',
    runtime_event_action: 'observe_platform_candidates',
    message_types: unique(rows.map((row) => row.message_type)),
    required_permissions: unique(rows.map((row) => row.required_permission)),
    acceptance_gate: 'all_selected_platforms_must_support_candidate_observation',
    rationale: 'Host projects need low-latency meeting-window candidate observation to bind realtime annotations to the current meeting axis before provider events arrive.',
    rows,
  };
}

function buildMeetingTrackContract(platforms = [], speakerTrackMatrix = {}, participantTrackMatrix = {}) {
  const speakerRows = Object.fromEntries(asArray(speakerTrackMatrix.rows).map((row) => [row.platform, row]));
  const participantRows = Object.fromEntries(asArray(participantTrackMatrix.rows).map((row) => [row.platform, row]));
  const rows = platforms.map((platform) => {
    const speaker = speakerRows[platform] ?? {};
    const participant = participantRows[platform] ?? {};
    const speakerReady = Boolean(speakerRows[platform])
      && speaker.provider_events_block_realtime !== true
      && speaker.transcript_blocks_realtime !== true;
    const participantReady = Boolean(participantRows[platform])
      && participant.provider_events_block_realtime !== true
      && participant.transcript_blocks_realtime !== true;
    const missingItems = [
      !speakerReady ? 'speaker_track_realtime_contract' : null,
      !participantReady ? 'participant_track_realtime_contract' : null,
    ].filter(Boolean);
    return compactObject({
      platform,
      ready: speakerReady && participantReady,
      speaker_track_ready: speakerReady,
      participant_track_ready: participantReady,
      speaker_min_stable_ms: speaker.min_stable_ms,
      speaker_switch_stable_ms: speaker.switch_stable_ms,
      speaker_end_idle_ms: speaker.end_idle_ms,
      participant_duplicate_window_ms: participant.duplicate_window_ms,
      participant_leave_stable_ms: participant.leave_stable_ms,
      speaker_provider_blocks_realtime: speaker.provider_events_block_realtime === true,
      speaker_transcript_blocks_realtime: speaker.transcript_blocks_realtime === true,
      participant_provider_blocks_realtime: participant.provider_events_block_realtime === true,
      participant_transcript_blocks_realtime: participant.transcript_blocks_realtime === true,
      output_intents: ['speaker_track', 'participant_track'],
      runtime_event_actions: ['speaker_track', 'participant_track'],
      missing_items: missingItems,
    });
  });
  const speakerReadyCount = rows.filter((row) => row.speaker_track_ready === true).length;
  const participantReadyCount = rows.filter((row) => row.participant_track_ready === true).length;
  const providerBlockingCount = (speakerTrackMatrix.provider_blocking_count ?? 0)
    + (participantTrackMatrix.provider_blocking_count ?? 0);
  const transcriptBlockingCount = (speakerTrackMatrix.transcript_blocking_count ?? 0)
    + (participantTrackMatrix.transcript_blocking_count ?? 0);
  return {
    type: 'meeting_platform_meeting_track_contract',
    platform_count: platforms.length,
    speaker_ready_count: speakerReadyCount,
    participant_ready_count: participantReadyCount,
    ready_count: rows.filter((row) => row.ready === true).length,
    missing_count: rows.filter((row) => row.ready !== true).length,
    provider_blocking_count: providerBlockingCount,
    transcript_blocking_count: transcriptBlockingCount,
    all_ready: speakerReadyCount === platforms.length
      && participantReadyCount === platforms.length
      && providerBlockingCount === 0
      && transcriptBlockingCount === 0,
    acceptance_gate: 'all_selected_platforms_must_support_realtime_speaker_and_participant_tracks',
    rationale: 'Host projects need speaker and participant position markers on the realtime timeline without waiting for provider events or transcript export.',
    rows,
  };
}

export function buildMeetingPlatformHostIntegrationPlan(options = {}) {
  const platforms = selectedPlatforms(options);
  const baseUrl = firstNonEmpty(options.baseUrl, options.base_url, 'http://localhost:8787');
  const basePath = firstNonEmpty(options.basePath, options.base_path, '/api/platform-events');
  const handoff = buildMeetingPlatformLiveAdapterHandoffBundle({
    ...options,
    platforms,
    baseUrl,
    basePath,
  });
  const adapterContractMatrix = buildMeetingPlatformAdapterContractMatrix({
    ...options,
    platforms,
    baseUrl,
    basePath,
  });
  const adapterContractAcceptanceMatrix = buildMeetingPlatformAdapterContractAcceptanceMatrix({
    ...options,
    platforms,
    baseUrl,
    basePath,
  });
  const platformConformanceReport = buildMeetingPlatformConformanceReport({
    ...options,
    platforms,
    baseUrl,
    basePath,
  });
  const runtimeBundleMatrix = buildMeetingPlatformRuntimeBundleMatrix({
    ...options,
    platforms,
    baseUrl,
  });
  const observerPlanMatrix = buildMeetingAppRuntimeObserverPlanMatrix({
    ...options,
    platforms: platforms.filter((platform) => platform !== 'local_detector'),
  });
  const candidateObservationContract = buildCandidateObservationContract(platforms, runtimeBundleMatrix);
  const speakerTrackMatrix = buildMeetingPlatformSpeakerTrackMatrix({
    ...options,
    platforms,
    baseUrl,
  });
  const participantTrackMatrix = buildMeetingPlatformParticipantTrackMatrix({
    ...options,
    platforms,
    baseUrl,
  });
  const meetingTrackContract = buildMeetingTrackContract(platforms, speakerTrackMatrix, participantTrackMatrix);
  const runtimeEventPlanMatrix = buildMeetingPlatformRuntimeEventPlanMatrix({
    ...options,
    platforms,
    baseUrl,
  });
  const adapterRouteMatrix = buildMeetingPlatformAdapterRouteMatrix({
    ...options,
    platforms,
    baseUrl,
    basePath,
  });
  const adapterBlueprintMatrix = buildMeetingPlatformAdapterBlueprintMatrix({
    ...options,
    platforms,
    baseUrl,
    basePath,
  });
  const adapterStartupPlanMatrix = buildMeetingPlatformAdapterStartupPlanMatrix({}, {
    ...options,
    platforms,
    baseUrl,
    basePath,
  });
  const adapterPreflightMatrix = buildMeetingPlatformAdapterPreflightMatrix({}, {
    ...options,
    platforms,
    baseUrl,
    basePath,
  });
  const adapterRuntimeContract = buildAdapterRuntimeContract(platforms, adapterRouteMatrix, adapterBlueprintMatrix);
  const adaptationStrategyMatrix = buildMeetingPlatformAdaptationStrategyMatrix({
    ...options,
    platforms,
    baseUrl,
    basePath,
  });
  const extensionInstallPlan = buildMeetingAppExtensionInstallPlan({
    ...options,
    platforms: platforms.filter((platform) => platform !== 'local_detector'),
    baseUrl,
  });
  return compactObject({
    type: 'meeting_platform_host_integration',
    schema: MEETING_PLATFORM_HOST_INTEGRATION_SCHEMA,
    schema_version: MEETING_PLATFORM_HOST_INTEGRATION_SCHEMA_VERSION,
    base_url: baseUrl,
    base_path: basePath,
    platforms,
    platform_imports: platformImportNames(platforms),
    sdk: {
      package: '@ai-annotation/meeting-timeline-sdk',
      kit_module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit',
      integration_runtime_module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime',
      runtime_event_module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-event',
      live_adapter_module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-live-adapter',
      host_integration_module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-host-integration',
      platform_conformance_module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-conformance',
      consumer_handoff_module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-consumer-handoff',
      adapter_startup_module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-startup',
      adapter_preflight_module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-preflight',
    },
    runtime_contract: {
      annotation_timestamp_field: 'captured_at_ms',
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      per_meeting_annotation_isolation_required: true,
      local_observer_may_start_axis: true,
      candidate_observation_required_for_host_axis_binding: true,
      speaker_track_required_for_realtime_timeline: true,
      participant_track_required_for_realtime_timeline: true,
    },
    endpoints: {
      platform_events: basePath,
      annotations: '/api/annotations',
      readiness: '/api/meeting-platform/readiness',
      handoff: '/api/meeting-platform/handoff',
      adapter_contracts: '/api/meeting-platform/contracts',
      adapter_contract_acceptance: '/api/meeting-platform/contract-acceptance',
      platform_conformance: '/api/meeting-platform/conformance',
      consumer_handoff: '/api/meeting-platform/consumer-handoff',
      runtime_bundles: '/api/meeting-platform/runtime-bundles',
      observer_plans: '/api/meeting-platform/observer-plans',
      runtime_event_plans: '/api/meeting-platform/runtime-event-plans',
      adapter_routes: '/api/meeting-platform/adapter-routes',
      adapter_blueprints: '/api/meeting-platform/adapter-blueprints',
      adapter_startup: '/api/meeting-platform/adapter-startup',
      adapter_preflight: '/api/meeting-platform/adapter-preflight',
      strategy: '/api/meeting-platform/strategy',
      platform_resolution: '/api/meeting-platform/resolve',
      platform_candidate_resolution: '/api/meeting-platform/resolve-candidates',
      platform_candidate_observation: '/api/meeting-platform/observe-candidates',
      extension_plan: '/api/meeting-platform/extension-plan',
      integration_runtime: '/api/meeting-platform/integration-runtime',
      integration_runtime_manifest: '/api/meeting-platform/integration-runtime/manifest',
      integration_runtime_run_manifest: '/api/meeting-platform/integration-runtime/run-manifest',
      handoff_readiness: '/api/meeting-platform/handoff-readiness',
      runtime_events: '/api/meeting-platform/runtime-events',
    },
    commands: {
      ...(handoff.commands ?? {}),
      validate_platform_conformance: 'npm run meeting-platform:conformance',
      export_consumer_handoff: 'npm run meeting-platform:consumer-handoff',
      export_adapter_startup: 'npm run meeting-platform:adapter-startup',
      run_adapter_preflight: 'npm run meeting-platform:adapter-preflight',
      validate_integration_runtime_manifest: 'npm run meeting-platform:integration-runtime-run-manifest',
      validate_handoff_readiness: 'npm run meeting-platform:handoff-readiness',
    },
    evidence_paths: handoff.evidence_paths,
    integration_plans: Object.fromEntries(platforms.map((platform) => [
      platform,
      buildPlatformIntegrationPlan(platform, { ...options, baseUrl, basePath }),
    ])),
    extension_install_plan: extensionInstallPlan,
    adapter_route_matrix: adapterRouteMatrix,
    adapter_blueprint_matrix: adapterBlueprintMatrix,
    adapter_startup_plan_matrix: adapterStartupPlanMatrix,
    adapter_preflight_matrix: adapterPreflightMatrix,
    adapter_runtime_contract: adapterRuntimeContract,
    adaptation_strategy_matrix: adaptationStrategyMatrix,
    runtime_event_plan_matrix: runtimeEventPlanMatrix,
    runtime_bundle_matrix: runtimeBundleMatrix,
    observer_plan_matrix: observerPlanMatrix,
    candidate_observation_contract: candidateObservationContract,
    speaker_track_matrix: speakerTrackMatrix,
    participant_track_matrix: participantTrackMatrix,
    meeting_track_contract: meetingTrackContract,
    handoff_bundle: handoff,
    adapter_contract_matrix: adapterContractMatrix,
    adapter_contract_acceptance_matrix: adapterContractAcceptanceMatrix,
    platform_conformance_report: platformConformanceReport,
    next_actions: unique([
      ...(handoff.next_actions ?? []),
      'export_meeting_platform_consumer_handoff_for_downstream_project',
      'run_meeting_platform_conformance_before_host_handoff',
      'run_meeting_platform_contract_acceptance_before_enabling_new_platform',
      'wire_host_routes_to_handleMeetingPlatformRequest',
      'capture_real_meeting_app_snapshots_for_each_target_platform',
      'wire_observer_plans_to_host_scheduler',
      'publish_adapter_blueprints_endpoint_for_downstream_hosts',
      'publish_adapter_startup_plan_endpoint_for_downstream_hosts',
      'generate_adapter_startup_plan_before_runtime_surface_install',
      'run_adapter_preflight_with_live_window_evidence_before_realtime_marks',
      'wire_platform_adapter_runtime_entries_into_host_surface',
      'verify_candidate_observation_contract_for_each_target_platform',
      'verify_meeting_track_contract_for_each_target_platform',
      'run_meeting_platform_integration_runtime_manifest_before_host_handoff',
      'run_meeting_platform_handoff_readiness_before_host_handoff',
      'capture_real_provider_events_where_available',
    ]),
  });
}

export function buildMeetingPlatformHostIntegrationScaffold(options = {}) {
  const plan = buildMeetingPlatformHostIntegrationPlan(options);
  const packageName = firstNonEmpty(options.packageName, options.package_name, 'meeting-platform-timeline-host');
  const packageJson = {
    type: 'module',
    private: true,
    name: packageName,
    scripts: {
      'meeting-platform:handoff': 'node ./scripts/print-handoff.mjs',
      'meeting-platform:strategy': 'node ./scripts/print-strategy.mjs',
      'meeting-platform:readiness': 'node ./scripts/print-readiness.mjs',
      'meeting-platform:contracts': 'node ./scripts/print-contracts.mjs',
      'meeting-platform:contract-acceptance': 'node ./scripts/verify-contracts.mjs',
      'meeting-platform:conformance': 'node ./scripts/verify-conformance.mjs',
      'meeting-platform:consumer-handoff': 'node ./scripts/print-consumer-handoff.mjs',
      'meeting-platform:runtime-bundles': 'node ./scripts/print-runtime-bundles.mjs',
      'meeting-platform:observer-plans': 'node ./scripts/print-observer-plans.mjs',
      'meeting-platform:runtime-event-plans': 'node ./scripts/print-runtime-event-plans.mjs',
      'meeting-platform:adapter-routes': 'node ./scripts/print-adapter-routes.mjs',
      'meeting-platform:adapter-blueprints': 'node ./scripts/print-adapter-blueprints.mjs',
      'meeting-platform:adapter-startup': 'node ./scripts/print-adapter-startup.mjs',
      'meeting-platform:adapter-preflight': 'node ./scripts/print-adapter-preflight.mjs',
      'meeting-platform:adapters': 'node ./scripts/print-platform-adapters.mjs',
      'meeting-platform:extension-plan': 'node ./scripts/print-extension-plan.mjs',
      'meeting-platform:resolve': 'node ./scripts/resolve-platform.mjs',
      'meeting-platform:resolve-candidates': 'node ./scripts/resolve-platform-candidates.mjs',
      'meeting-platform:observe-candidates': 'node ./scripts/observe-platform-candidates.mjs',
      'meeting-platform:integration-runtime': 'node ./scripts/print-integration-runtime.mjs',
      'meeting-platform:integration-runtime-manifest': 'node ./scripts/print-integration-runtime-manifest.mjs',
      'meeting-platform:integration-runtime-run-manifest': 'node ./scripts/run-integration-runtime-manifest.mjs',
      'meeting-platform:handoff-readiness': 'node ./scripts/run-handoff-readiness.mjs',
    },
    dependencies: {
      '@ai-annotation/meeting-timeline-sdk': '^0.1.0',
    },
  };
  const handoffScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

console.log(JSON.stringify(host.handoffBundle(), null, 2));
`;
  const strategyScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

console.log(JSON.stringify(host.adaptationStrategyMatrix(), null, 2));
`;
  const platformResolutionScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

console.log(JSON.stringify(host.resolvePlatform({
  url: process.env.MEETING_PLATFORM_URL,
  title: process.env.MEETING_PLATFORM_TITLE,
  platform: process.env.MEETING_PLATFORM,
}), null, 2));
`;
  const platformCandidateResolutionScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

const payload = process.env.MEETING_PLATFORM_CANDIDATES_JSON
  ? JSON.parse(process.env.MEETING_PLATFORM_CANDIDATES_JSON)
  : {
    windows: [{
      tabs: [{
        url: process.env.MEETING_PLATFORM_URL,
        title: process.env.MEETING_PLATFORM_TITLE,
        active: true,
      }],
    }],
  };

console.log(JSON.stringify(host.resolvePlatformCandidates(payload), null, 2));
`;
  const platformCandidateObservationScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

const payload = process.env.MEETING_PLATFORM_CANDIDATES_JSON
  ? JSON.parse(process.env.MEETING_PLATFORM_CANDIDATES_JSON)
  : {
    windows: [{
      tabs: [{
        url: process.env.MEETING_PLATFORM_URL,
        title: process.env.MEETING_PLATFORM_TITLE,
        active: true,
      }],
    }],
  };

console.log(JSON.stringify(await host.observePlatformCandidates(payload), null, 2));
`;
  const readinessScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

console.log(JSON.stringify(host.liveAdapters.readinessMatrix(), null, 2));
`;
  const contractsScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

console.log(JSON.stringify(host.adapterContracts(), null, 2));
`;
  const contractAcceptanceScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

const report = host.adapterContractAcceptance({
  acceptanceTarget: process.env.MEETING_PLATFORM_ACCEPTANCE_TARGET ?? 'contract',
});

console.log(JSON.stringify(report, null, 2));

if (report.rejected_count > 0) {
  process.exitCode = 1;
}
`;
  const conformanceScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

const report = host.platformConformance();

console.log(JSON.stringify(report, null, 2));

if (report.accepted !== true) {
  process.exitCode = 1;
}
`;
  const consumerHandoffScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

const report = host.consumerHandoff({
  requireHandoffReady: process.env.MEETING_PLATFORM_REQUIRE_HANDOFF_READY === '1',
  requireProductionReady: process.env.MEETING_PLATFORM_REQUIRE_PRODUCTION_READY === '1',
});

console.log(JSON.stringify(report, null, 2));

if (report.accepted !== true) {
  process.exitCode = 1;
}
`;
  const runtimeBundlesScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

console.log(JSON.stringify(host.runtimeBundles(), null, 2));
`;
  const observerPlansScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

console.log(JSON.stringify(host.observerPlans(), null, 2));
`;
  const runtimeEventPlansScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

console.log(JSON.stringify(host.runtimeEventPlans(), null, 2));
`;
  const adapterRoutesScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

console.log(JSON.stringify(host.adapterRoutes(), null, 2));
`;
  const adapterBlueprintsScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

console.log(JSON.stringify(host.adapterBlueprints(), null, 2));
`;
  const adapterStartupScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

const input = process.env.MEETING_PLATFORM_STARTUP_INPUT_JSON
  ? JSON.parse(process.env.MEETING_PLATFORM_STARTUP_INPUT_JSON)
  : {};

console.log(JSON.stringify(host.adapterStartupPlans({ input }), null, 2));
`;
  const adapterPreflightScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

const input = process.env.MEETING_PLATFORM_PREFLIGHT_INPUT_JSON
  ? JSON.parse(process.env.MEETING_PLATFORM_PREFLIGHT_INPUT_JSON)
  : {};
const mode = process.env.MEETING_PLATFORM_PREFLIGHT_MODE ?? 'matrix';

const result = mode === 'candidates' || mode === 'candidate'
  ? host.adapterCandidatePreflight(input)
  : mode === 'single'
    ? host.adapterPreflight(input)
    : host.adapterPreflightMatrix(input);

console.log(JSON.stringify(result, null, 2));
`;
  const platformAdaptersScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

const adapters = host.platformAdapters();
const summary = Object.fromEntries(Object.entries(adapters).map(([platform, adapter]) => [platform, {
  platform: adapter.platform,
  selected_surface: adapter.selected_surface,
  first_route: adapter.first_route,
  timestamp_field: adapter.runtime_contract?.timestamp_field,
  provider_blocks_realtime: adapter.runtime_contract?.provider_blocks_realtime,
  transcript_blocks_realtime: adapter.runtime_contract?.transcript_blocks_realtime,
}]));

console.log(JSON.stringify(summary, null, 2));
`;
  const extensionPlanScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

console.log(JSON.stringify(host.extensionInstallPlan(), null, 2));
`;
  const integrationRuntimeScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

console.log(JSON.stringify(host.integrationRuntimeSummary(), null, 2));
`;
  const integrationRuntimeManifestScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

console.log(JSON.stringify(host.integrationRuntimeManifest(), null, 2));
`;
  const integrationRuntimeRunManifestScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

const options = process.env.MEETING_PLATFORM_MANIFEST_OPTIONS_JSON
  ? JSON.parse(process.env.MEETING_PLATFORM_MANIFEST_OPTIONS_JSON)
  : {
    requireHandoffReady: process.env.MEETING_PLATFORM_REQUIRE_HANDOFF_READY === '1',
    target: process.env.MEETING_PLATFORM_HANDOFF_TARGET ?? 'production',
  };

const report = await host.runIntegrationRuntimeManifest(options);

console.log(JSON.stringify(report, null, 2));

if (report.host_integration_ready !== true) {
  process.exitCode = 1;
}
`;
  const handoffReadinessScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

const options = process.env.MEETING_PLATFORM_HANDOFF_OPTIONS_JSON
  ? JSON.parse(process.env.MEETING_PLATFORM_HANDOFF_OPTIONS_JSON)
  : {
    target: process.env.MEETING_PLATFORM_HANDOFF_TARGET ?? 'production',
  };

const report = await host.runHandoffReadiness(options);

console.log(JSON.stringify(report, null, 2));

if (
  process.env.MEETING_PLATFORM_REQUIRE_HANDOFF_READY === '1'
  && report.handoff_ready_count !== report.platform_count
) {
  process.exitCode = 1;
}
`;
  return {
    type: 'meeting_platform_host_integration_scaffold',
    schema: MEETING_PLATFORM_HOST_INTEGRATION_SCAFFOLD_SCHEMA,
    schema_version: MEETING_PLATFORM_HOST_INTEGRATION_SCHEMA_VERSION,
    platforms: plan.platforms,
    plan,
    files: [
      sourceFile('package.json', json(packageJson), 'package_manifest', 'application/json'),
      sourceFile('src/timeline-client.mjs', timelineClientSource(plan), 'timeline_client_source', 'text/javascript'),
      sourceFile('src/meeting-platform-host.mjs', hostSource(plan.platforms, plan), 'host_source', 'text/javascript'),
      sourceFile('src/http-routes.mjs', routesSource(plan), 'http_routes_source', 'text/javascript'),
      sourceFile('src/platform-adapters/index.mjs', platformAdaptersIndexSource(plan.platforms), 'platform_adapter_registry_source', 'text/javascript'),
      ...plan.platforms.map((platform) => sourceFile(
        adapterFilePath(platform),
        platformAdapterSource(platform, plan),
        'platform_adapter_source',
        'text/javascript',
      )),
      sourceFile('scripts/print-handoff.mjs', handoffScript, 'handoff_script', 'text/javascript'),
      sourceFile('scripts/print-strategy.mjs', strategyScript, 'adaptation_strategy_script', 'text/javascript'),
      sourceFile('scripts/print-readiness.mjs', readinessScript, 'readiness_script', 'text/javascript'),
      sourceFile('scripts/print-contracts.mjs', contractsScript, 'adapter_contract_script', 'text/javascript'),
      sourceFile('scripts/verify-contracts.mjs', contractAcceptanceScript, 'adapter_contract_acceptance_script', 'text/javascript'),
      sourceFile('scripts/verify-conformance.mjs', conformanceScript, 'platform_conformance_script', 'text/javascript'),
      sourceFile('scripts/print-consumer-handoff.mjs', consumerHandoffScript, 'consumer_handoff_script', 'text/javascript'),
      sourceFile('scripts/print-runtime-bundles.mjs', runtimeBundlesScript, 'runtime_bundle_script', 'text/javascript'),
      sourceFile('scripts/print-observer-plans.mjs', observerPlansScript, 'observer_plan_script', 'text/javascript'),
      sourceFile('scripts/print-runtime-event-plans.mjs', runtimeEventPlansScript, 'runtime_event_plan_script', 'text/javascript'),
      sourceFile('scripts/print-adapter-routes.mjs', adapterRoutesScript, 'adapter_route_script', 'text/javascript'),
      sourceFile('scripts/print-adapter-blueprints.mjs', adapterBlueprintsScript, 'adapter_blueprint_script', 'text/javascript'),
      sourceFile('scripts/print-adapter-startup.mjs', adapterStartupScript, 'adapter_startup_script', 'text/javascript'),
      sourceFile('scripts/print-adapter-preflight.mjs', adapterPreflightScript, 'adapter_preflight_script', 'text/javascript'),
      sourceFile('scripts/print-platform-adapters.mjs', platformAdaptersScript, 'platform_adapter_script', 'text/javascript'),
      sourceFile('scripts/print-extension-plan.mjs', extensionPlanScript, 'extension_plan_script', 'text/javascript'),
      sourceFile('scripts/resolve-platform.mjs', platformResolutionScript, 'platform_resolution_script', 'text/javascript'),
      sourceFile('scripts/resolve-platform-candidates.mjs', platformCandidateResolutionScript, 'platform_candidate_resolution_script', 'text/javascript'),
      sourceFile('scripts/observe-platform-candidates.mjs', platformCandidateObservationScript, 'platform_candidate_observation_script', 'text/javascript'),
      sourceFile('scripts/print-integration-runtime.mjs', integrationRuntimeScript, 'integration_runtime_script', 'text/javascript'),
      sourceFile('scripts/print-integration-runtime-manifest.mjs', integrationRuntimeManifestScript, 'integration_runtime_manifest_script', 'text/javascript'),
      sourceFile('scripts/run-integration-runtime-manifest.mjs', integrationRuntimeRunManifestScript, 'integration_runtime_run_manifest_script', 'text/javascript'),
      sourceFile('scripts/run-handoff-readiness.mjs', handoffReadinessScript, 'handoff_readiness_script', 'text/javascript'),
      sourceFile('README.md', readmeSource(plan), 'readme', 'text/markdown'),
    ],
  };
}

export function buildMeetingPlatformHostIntegrationScaffoldAcceptanceReport(scaffoldOrOptions = {}, options = {}) {
  const scaffold = scaffoldOrOptions?.schema === MEETING_PLATFORM_HOST_INTEGRATION_SCAFFOLD_SCHEMA
    ? scaffoldOrOptions
    : buildMeetingPlatformHostIntegrationScaffold({ ...scaffoldOrOptions, ...options });
  const issues = [];
  const requiredFiles = [
    'package.json',
    'src/timeline-client.mjs',
    'src/meeting-platform-host.mjs',
    'src/http-routes.mjs',
    'src/platform-adapters/index.mjs',
    ...asArray(scaffold.platforms).map((platform) => adapterFilePath(platform)),
    'scripts/print-handoff.mjs',
    'scripts/print-strategy.mjs',
    'scripts/print-readiness.mjs',
    'scripts/print-contracts.mjs',
    'scripts/verify-contracts.mjs',
    'scripts/verify-conformance.mjs',
    'scripts/print-consumer-handoff.mjs',
    'scripts/print-runtime-bundles.mjs',
    'scripts/print-observer-plans.mjs',
    'scripts/print-runtime-event-plans.mjs',
    'scripts/print-adapter-routes.mjs',
    'scripts/print-adapter-blueprints.mjs',
    'scripts/print-adapter-startup.mjs',
    'scripts/print-adapter-preflight.mjs',
    'scripts/print-platform-adapters.mjs',
    'scripts/resolve-platform.mjs',
    'scripts/resolve-platform-candidates.mjs',
    'scripts/observe-platform-candidates.mjs',
    'scripts/print-extension-plan.mjs',
    'scripts/print-integration-runtime.mjs',
    'scripts/print-integration-runtime-manifest.mjs',
    'scripts/run-integration-runtime-manifest.mjs',
    'scripts/run-handoff-readiness.mjs',
    'README.md',
  ];
  for (const path of requiredFiles) {
    if (!fileByPath(scaffold, path)) {
      issues.push(issue('error', 'missing_required_file', `Missing required scaffold file: ${path}`, { path }));
    }
  }
  const host = fileByPath(scaffold, 'src/meeting-platform-host.mjs')?.content ?? '';
  const adapterRegistry = fileByPath(scaffold, 'src/platform-adapters/index.mjs')?.content ?? '';
  const routes = fileByPath(scaffold, 'src/http-routes.mjs')?.content ?? '';
  const readme = fileByPath(scaffold, 'README.md')?.content ?? '';
  if (!host.includes('createMeetingPlatformIntegrationRuntime')) {
    issues.push(issue('error', 'missing_integration_runtime_import', 'Host source must create the platform integration runtime.'));
  }
  if (!host.includes('integrationRuntimeManifest')) {
    issues.push(issue('error', 'missing_integration_runtime_manifest', 'Host source must expose the integration runtime manifest.'));
  }
  if (!host.includes('runIntegrationRuntimeManifest')) {
    issues.push(issue('error', 'missing_integration_runtime_run_manifest', 'Host source must expose the executable integration runtime manifest gate.'));
  }
  if (!host.includes('runHandoffReadiness')) {
    issues.push(issue('error', 'missing_handoff_readiness_runner', 'Host source must expose the runtime handoff readiness runner.'));
  }
  if (!host.includes('integrationRuntimeSummary')) {
    issues.push(issue('error', 'missing_integration_runtime_summary', 'Host source must expose the integration runtime summary.'));
  }
  if (!host.includes('liveAdapters')) {
    issues.push(issue('error', 'missing_live_adapter_suite', 'Host source must expose the live adapter suite.'));
  }
  if (!host.includes('platformAdapterContractMatrix')) {
    issues.push(issue('error', 'missing_adapter_contract_matrix', 'Host source must expose the adapter contract matrix.'));
  }
  if (!host.includes('platformAdapterContractAcceptanceMatrix')) {
    issues.push(issue('error', 'missing_adapter_contract_acceptance_matrix', 'Host source must expose the adapter contract acceptance matrix.'));
  }
  if (!host.includes('platformConformance')) {
    issues.push(issue('error', 'missing_platform_conformance', 'Host source must expose the platform conformance report.'));
  }
  if (!host.includes('consumerHandoff')) {
    issues.push(issue('error', 'missing_consumer_handoff', 'Host source must expose the consumer handoff index.'));
  }
  if (!host.includes('platformRuntimeBundleMatrix')) {
    issues.push(issue('error', 'missing_runtime_bundle_matrix', 'Host source must expose the runtime bundle matrix.'));
  }
  if (!host.includes('meetingAppRuntimeObserverPlanMatrix')) {
    issues.push(issue('error', 'missing_observer_plan_matrix', 'Host source must expose meeting app runtime observer plans.'));
  }
  if (!host.includes('platformRuntimeEventPlanMatrix')) {
    issues.push(issue('error', 'missing_runtime_event_plan_matrix', 'Host source must expose the runtime event plan matrix.'));
  }
  if (!host.includes('platformAdapterRouteMatrix')) {
    issues.push(issue('error', 'missing_adapter_route_matrix', 'Host source must expose the platform adapter route matrix.'));
  }
  if (!host.includes('platformAdapterBlueprintMatrix')) {
    issues.push(issue('error', 'missing_adapter_blueprint_matrix', 'Host source must expose the platform adapter blueprint matrix.'));
  }
  if (!host.includes('platformAdapterStartupPlanMatrix')) {
    issues.push(issue('error', 'missing_adapter_startup_plan_matrix', 'Host source must expose the platform adapter startup plan matrix.'));
  }
  if (!host.includes('platformAdapterPreflightMatrix')) {
    issues.push(issue('error', 'missing_adapter_preflight_matrix', 'Host source must expose the platform adapter preflight matrix.'));
  }
  if (!host.includes('platformAdapterCandidatePreflight')) {
    issues.push(issue('error', 'missing_adapter_candidate_preflight', 'Host source must expose candidate preflight for active meeting windows.'));
  }
  if (!host.includes('platformAdaptationStrategyMatrix')) {
    issues.push(issue('error', 'missing_adaptation_strategy_matrix', 'Host source must expose the adaptation strategy matrix.'));
  }
  if (!host.includes('resolvePlatform')) {
    issues.push(issue('error', 'missing_platform_resolution', 'Host source must expose platform resolution for URL/window based adapter selection.'));
  }
  if (!host.includes('resolvePlatformCandidates')) {
    issues.push(issue('error', 'missing_platform_candidate_resolution', 'Host source must expose platform candidate resolution for desktop/window snapshots.'));
  }
  if (!host.includes('observePlatformCandidates')) {
    issues.push(issue('error', 'missing_platform_candidate_observation', 'Host source must observe platform candidates and apply current-axis signals.'));
  }
  if (!host.includes('meetingAppExtensionInstallPlan')) {
    issues.push(issue('error', 'missing_extension_install_plan', 'Host source must expose the extension install plan.'));
  }
  if (!host.includes('capturedAtMs')) {
    issues.push(issue('error', 'missing_captured_at_ms_write', 'Host source must write annotations with capturedAtMs.'));
  }
  if (!host.includes('speakerTrack')) {
    issues.push(issue('error', 'missing_speaker_track_method', 'Host source must expose speakerTrack for realtime speaker-position markers.'));
  }
  if (!host.includes('participantTrack')) {
    issues.push(issue('error', 'missing_participant_track_method', 'Host source must expose participantTrack for realtime participant-position markers.'));
  }
  if (!host.includes('platformAdapter')) {
    issues.push(issue('error', 'missing_platform_adapter_method', 'Host source must expose per-platform adapter factories.'));
  }
  if (!adapterRegistry.includes('createMeetingPlatformAdapter')) {
    issues.push(issue('error', 'missing_platform_adapter_registry', 'Scaffold must expose a platform adapter registry.'));
  }
  for (const platform of asArray(scaffold.platforms)) {
    const normalized = normalizeMeetingPlatform(platform);
    const adapterSource = fileByPath(scaffold, adapterFilePath(normalized))?.content ?? '';
    if (!adapterSource.includes('insertAnnotation')) {
      issues.push(issue('error', 'missing_platform_adapter_insert_annotation', 'Platform adapter must expose insertAnnotation.', { platform: normalized }));
    }
    if (!adapterSource.includes('captured_at_ms')) {
      issues.push(issue('error', 'missing_platform_adapter_captured_at_ms', 'Platform adapter must enforce captured_at_ms.', { platform: normalized }));
    }
    if (!adapterSource.includes('observeCandidates')) {
      issues.push(issue('error', 'missing_platform_adapter_observe_candidates', 'Platform adapter must expose observeCandidates.', { platform: normalized }));
    }
  }
  if (!routes.includes('handleFetchRequest')) {
    issues.push(issue('error', 'missing_provider_route_handler', 'Route source must forward provider webhooks to the SDK handler.'));
  }
  if (!routes.includes('/api/meeting-platform/contracts')) {
    issues.push(issue('error', 'missing_contract_route', 'Route source must expose the adapter contract matrix endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/conformance')) {
    issues.push(issue('error', 'missing_platform_conformance_route', 'Route source must expose the platform conformance endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/consumer-handoff')) {
    issues.push(issue('error', 'missing_consumer_handoff_route', 'Route source must expose the consumer handoff endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/runtime-bundles')) {
    issues.push(issue('error', 'missing_runtime_bundle_route', 'Route source must expose the runtime bundle matrix endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/observer-plans')) {
    issues.push(issue('error', 'missing_observer_plan_route', 'Route source must expose the observer plan matrix endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/runtime-event-plans')) {
    issues.push(issue('error', 'missing_runtime_event_plan_route', 'Route source must expose the runtime event plan matrix endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/adapter-routes')) {
    issues.push(issue('error', 'missing_adapter_route_route', 'Route source must expose the platform adapter route matrix endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/adapter-blueprints')) {
    issues.push(issue('error', 'missing_adapter_blueprint_route', 'Route source must expose the platform adapter blueprint matrix endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/adapter-startup')) {
    issues.push(issue('error', 'missing_adapter_startup_route', 'Route source must expose the platform adapter startup plan endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/adapter-preflight')) {
    issues.push(issue('error', 'missing_adapter_preflight_route', 'Route source must expose the platform adapter preflight endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/strategy')) {
    issues.push(issue('error', 'missing_adaptation_strategy_route', 'Route source must expose the adaptation strategy matrix endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/resolve')) {
    issues.push(issue('error', 'missing_platform_resolution_route', 'Route source must expose the platform resolution endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/resolve-candidates')) {
    issues.push(issue('error', 'missing_platform_candidate_resolution_route', 'Route source must expose the platform candidate resolution endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/observe-candidates')) {
    issues.push(issue('error', 'missing_platform_candidate_observation_route', 'Route source must expose the platform candidate observation endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/extension-plan')) {
    issues.push(issue('error', 'missing_extension_plan_route', 'Route source must expose the extension install plan endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/integration-runtime')) {
    issues.push(issue('error', 'missing_integration_runtime_route', 'Route source must expose the integration runtime summary endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/integration-runtime/run-manifest')) {
    issues.push(issue('error', 'missing_integration_runtime_run_manifest_route', 'Route source must expose the executable integration runtime manifest endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/handoff-readiness')) {
    issues.push(issue('error', 'missing_handoff_readiness_route', 'Route source must expose the handoff readiness endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/runtime-events')) {
    issues.push(issue('error', 'missing_runtime_events_route', 'Route source must expose the runtime event endpoint.'));
  }
  if (!readme.includes('captured_at_ms')) {
    issues.push(issue('warning', 'readme_missing_timestamp_contract', 'README should state the captured_at_ms contract.'));
  }
  if (!readme.includes('meeting_timeline.observe_candidates')) {
    issues.push(issue('warning', 'readme_missing_candidate_observation_contract', 'README should state the candidate observation contract.'));
  }
  if (!readme.includes('Meeting track contract')) {
    issues.push(issue('warning', 'readme_missing_meeting_track_contract', 'README should state the speaker/participant track contract.'));
  }
  if (!readme.includes('Observer plan contract')) {
    issues.push(issue('warning', 'readme_missing_observer_plan_contract', 'README should state the observer plan contract.'));
  }
  if (!readme.includes('Adapter blueprints')) {
    issues.push(issue('warning', 'readme_missing_adapter_blueprint_contract', 'README should state the adapter blueprint contract.'));
  }
  if (!readme.includes('Adapter startup plans')) {
    issues.push(issue('warning', 'readme_missing_adapter_startup_plan', 'README should state the adapter startup plan contract.'));
  }
  if (!readme.includes('Adapter preflight')) {
    issues.push(issue('warning', 'readme_missing_adapter_preflight', 'README should state the adapter preflight contract.'));
  }
  if (!readme.includes('Platform conformance')) {
    issues.push(issue('warning', 'readme_missing_platform_conformance', 'README should state the platform conformance gate.'));
  }
  if (!readme.includes('Consumer handoff')) {
    issues.push(issue('warning', 'readme_missing_consumer_handoff', 'README should state the consumer handoff index.'));
  }
  const platformConformanceReport = scaffold.plan?.platform_conformance_report;
  if (!platformConformanceReport) {
    issues.push(issue('error', 'missing_platform_conformance_report', 'Scaffold plan must include the platform conformance report.'));
  } else if (platformConformanceReport.accepted !== true) {
    issues.push(issue('error', 'platform_conformance_not_accepted', 'Every selected platform must pass SDK conformance before host handoff.', {
      accepted_count: platformConformanceReport.accepted_count,
      platform_count: platformConformanceReport.platform_count,
      blocking_count: platformConformanceReport.blocking_count,
    }));
  }
  const candidateObservationContract = scaffold.plan?.candidate_observation_contract;
  if (!candidateObservationContract) {
    issues.push(issue('error', 'missing_candidate_observation_contract', 'Scaffold plan must include the candidate observation contract.'));
  } else if (candidateObservationContract.all_ready !== true) {
    issues.push(issue('error', 'candidate_observation_contract_not_ready', 'Every selected platform must support candidate observation before host handoff.', {
      missing_count: candidateObservationContract.missing_count,
    }));
  }
  const runtimeCandidateObserverCount = scaffold.plan?.runtime_bundle_matrix?.candidate_observer_count;
  if (runtimeCandidateObserverCount != null && runtimeCandidateObserverCount !== scaffold.platforms?.length) {
    issues.push(issue('error', 'runtime_bundle_candidate_observer_not_ready', 'Runtime bundle matrix must expose candidate observation for every selected platform.', {
      candidate_observer_count: runtimeCandidateObserverCount,
      platform_count: scaffold.platforms?.length ?? 0,
    }));
  }
  const observerPlanMatrix = scaffold.plan?.observer_plan_matrix;
  if (!observerPlanMatrix) {
    issues.push(issue('error', 'missing_observer_plan_matrix', 'Scaffold plan must include the observer plan matrix.'));
  } else if (observerPlanMatrix.sdk_ready_count !== observerPlanMatrix.platform_count) {
    issues.push(issue('error', 'observer_plan_matrix_not_ready', 'Every selected meeting app platform must expose a runtime observer plan.', {
      platform_count: observerPlanMatrix.platform_count,
      sdk_ready_count: observerPlanMatrix.sdk_ready_count,
    }));
  }
  const meetingTrackContract = scaffold.plan?.meeting_track_contract;
  if (!meetingTrackContract) {
    issues.push(issue('error', 'missing_meeting_track_contract', 'Scaffold plan must include the speaker/participant meeting track contract.'));
  } else if (meetingTrackContract.all_ready !== true) {
    issues.push(issue('error', 'meeting_track_contract_not_ready', 'Every selected platform must support realtime speaker and participant tracks before host handoff.', {
      missing_count: meetingTrackContract.missing_count,
      speaker_ready_count: meetingTrackContract.speaker_ready_count,
      participant_ready_count: meetingTrackContract.participant_ready_count,
      provider_blocking_count: meetingTrackContract.provider_blocking_count,
      transcript_blocking_count: meetingTrackContract.transcript_blocking_count,
    }));
  }
  const adapterBlueprintMatrix = scaffold.plan?.adapter_blueprint_matrix;
  if (!adapterBlueprintMatrix) {
    issues.push(issue('error', 'missing_adapter_blueprint_matrix', 'Scaffold plan must include the adapter blueprint matrix.'));
  } else if (adapterBlueprintMatrix.ready_count !== adapterBlueprintMatrix.platform_count) {
    issues.push(issue('error', 'adapter_blueprint_matrix_not_ready', 'Every selected platform must expose a ready adapter blueprint before host handoff.', {
      platform_count: adapterBlueprintMatrix.platform_count,
      ready_count: adapterBlueprintMatrix.ready_count,
    }));
  }
  const adapterStartupPlanMatrix = scaffold.plan?.adapter_startup_plan_matrix;
  if (!adapterStartupPlanMatrix) {
    issues.push(issue('error', 'missing_adapter_startup_plan_matrix', 'Scaffold plan must include the adapter startup plan matrix.'));
  } else if (adapterStartupPlanMatrix.realtime_startup_ready_count !== adapterStartupPlanMatrix.platform_count) {
    issues.push(issue('error', 'adapter_startup_plan_matrix_not_ready', 'Every selected platform must expose a ready adapter startup plan before host handoff.', {
      platform_count: adapterStartupPlanMatrix.platform_count,
      realtime_startup_ready_count: adapterStartupPlanMatrix.realtime_startup_ready_count,
      accepted_count: adapterStartupPlanMatrix.accepted_count,
    }));
  }
  const adapterPreflightMatrix = scaffold.plan?.adapter_preflight_matrix;
  if (!adapterPreflightMatrix) {
    issues.push(issue('error', 'missing_adapter_preflight_matrix', 'Scaffold plan must include the adapter preflight matrix.'));
  } else if (adapterPreflightMatrix.platform_count !== (scaffold.platforms?.length ?? 0)) {
    issues.push(issue('error', 'adapter_preflight_matrix_platform_mismatch', 'Adapter preflight matrix must cover every selected platform.', {
      platform_count: adapterPreflightMatrix.platform_count,
      scaffold_platform_count: scaffold.platforms?.length ?? 0,
    }));
  }
  const adapterRuntimeContract = scaffold.plan?.adapter_runtime_contract;
  if (!adapterRuntimeContract) {
    issues.push(issue('error', 'missing_adapter_runtime_contract', 'Scaffold plan must include the per-platform adapter runtime contract.'));
  } else if (adapterRuntimeContract.all_ready !== true) {
    issues.push(issue('error', 'adapter_runtime_contract_not_ready', 'Every selected platform must expose a ready per-platform adapter runtime entry.', {
      platform_count: adapterRuntimeContract.platform_count,
      ready_count: adapterRuntimeContract.ready_count,
      missing_count: adapterRuntimeContract.missing_count,
    }));
  }
  const packageJsonFile = fileByPath(scaffold, 'package.json');
  if (packageJsonFile) {
    try {
      const manifest = JSON.parse(packageJsonFile.content);
      if (!manifest.dependencies?.['@ai-annotation/meeting-timeline-sdk']) {
        issues.push(issue('error', 'missing_sdk_dependency', 'package.json must depend on @ai-annotation/meeting-timeline-sdk.'));
      }
    } catch (error) {
      issues.push(issue('error', 'invalid_package_json', 'package.json must be valid JSON.', {
        error: String(error?.message ?? error),
      }));
    }
  }
  const blocking = issues.filter((item) => item.severity === 'error');
  return {
    type: 'meeting_platform_host_integration_acceptance',
    schema: MEETING_PLATFORM_HOST_INTEGRATION_ACCEPTANCE_SCHEMA,
    schema_version: MEETING_PLATFORM_HOST_INTEGRATION_SCHEMA_VERSION,
    accepted: blocking.length === 0,
    platform_count: scaffold.platforms?.length ?? 0,
    file_count: scaffold.files?.length ?? 0,
    candidate_observation_ready: candidateObservationContract?.all_ready === true,
    candidate_observer_count: candidateObservationContract?.ready_count ?? runtimeCandidateObserverCount ?? 0,
    candidate_observer_missing_count: candidateObservationContract?.missing_count,
    candidate_observation_contract: candidateObservationContract,
    platform_conformance_ready: platformConformanceReport?.accepted === true,
    platform_conformance_accepted_count: platformConformanceReport?.accepted_count ?? 0,
    platform_conformance_blocking_count: platformConformanceReport?.blocking_count ?? 0,
    platform_conformance_report: platformConformanceReport,
    observer_plan_ready: observerPlanMatrix
      ? observerPlanMatrix.sdk_ready_count === observerPlanMatrix.platform_count
      : false,
    observer_plan_ready_count: observerPlanMatrix?.sdk_ready_count ?? 0,
    observer_plan_preflight_accepted_count: observerPlanMatrix?.preflight_accepted_count ?? 0,
    observer_plan_matrix: observerPlanMatrix,
    adapter_blueprint_ready: adapterBlueprintMatrix
      ? adapterBlueprintMatrix.ready_count === adapterBlueprintMatrix.platform_count
      : false,
    adapter_blueprint_ready_count: adapterBlueprintMatrix?.ready_count ?? 0,
    adapter_blueprint_matrix: adapterBlueprintMatrix,
    adapter_startup_ready: adapterStartupPlanMatrix
      ? adapterStartupPlanMatrix.realtime_startup_ready_count === adapterStartupPlanMatrix.platform_count
      : false,
    adapter_startup_ready_count: adapterStartupPlanMatrix?.realtime_startup_ready_count ?? 0,
    adapter_startup_plan_matrix: adapterStartupPlanMatrix,
    adapter_preflight_available: adapterPreflightMatrix?.platform_count === (scaffold.platforms?.length ?? 0),
    adapter_preflight_platform_count: adapterPreflightMatrix?.platform_count ?? 0,
    adapter_preflight_accepted_count: adapterPreflightMatrix?.accepted_count ?? 0,
    adapter_preflight_realtime_ready_count: adapterPreflightMatrix?.realtime_ready_count ?? 0,
    adapter_preflight_matrix: adapterPreflightMatrix,
    adapter_runtime_ready: adapterRuntimeContract?.all_ready === true,
    adapter_runtime_ready_count: adapterRuntimeContract?.ready_count ?? 0,
    adapter_runtime_contract: adapterRuntimeContract,
    meeting_track_ready: meetingTrackContract?.all_ready === true,
    speaker_track_ready_count: meetingTrackContract?.speaker_ready_count ?? 0,
    participant_track_ready_count: meetingTrackContract?.participant_ready_count ?? 0,
    meeting_track_contract: meetingTrackContract,
    required_files: requiredFiles,
    blocking_count: blocking.length,
    warning_count: issues.length - blocking.length,
    issues,
    scaffold,
  };
}

export function assertMeetingPlatformHostIntegrationScaffold(scaffoldOrOptions = {}, options = {}) {
  const report = buildMeetingPlatformHostIntegrationScaffoldAcceptanceReport(scaffoldOrOptions, options);
  if (!report.accepted) {
    throw new MeetingTimelineSdkError('Meeting platform host integration scaffold failed acceptance', {
      issues: report.issues,
      report,
    });
  }
  return report;
}
