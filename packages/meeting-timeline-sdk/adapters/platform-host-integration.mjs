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
  buildMeetingPlatformRuntimeBundleMatrix,
} from './platform-runtime-bundle.mjs';
import {
  buildMeetingPlatformRuntimeEventPlanMatrix,
} from './platform-runtime-event.mjs';
import {
  buildMeetingPlatformAdaptationStrategyMatrix,
} from './platform-strategy.mjs';
import {
  buildMeetingPlatformAdapterContractAcceptanceMatrix,
  buildMeetingPlatformAdapterContractMatrix,
} from './platform-adapter-contract.mjs';

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

  return {
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
    runtimeBundles(bundleOptions = {}) {
      return kit.platformRuntimeBundleMatrix({
        ...bundleOptions,
        platforms: bundleOptions.platforms ?? bundleOptions.platform_keys ?? platforms,
      });
    },
    runtimeEventPlans(planOptions = {}) {
      return kit.platformRuntimeEventPlanMatrix({
        ...planOptions,
        platforms: planOptions.platforms ?? planOptions.platform_keys ?? platforms,
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
    integrationRuntimeSummary(summaryOptions = {}) {
      return integrationRuntime.summary({
        ...summaryOptions,
        platforms: summaryOptions.platforms ?? summaryOptions.platform_keys ?? platforms,
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
    verifyReadiness(readinessOptions = {}) {
      return liveAdapters.assertReadinessMatrix({
        ...readinessOptions,
        platforms: readinessOptions.platforms ?? readinessOptions.platform_keys ?? platforms,
      });
    },
  };
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
  if (url.pathname === '/api/meeting-platform/runtime-bundles') {
    return Response.json(host.runtimeBundles(options));
  }
  if (url.pathname === '/api/meeting-platform/runtime-event-plans') {
    return Response.json(host.runtimeEventPlans(options));
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
  if (url.pathname === '/api/meeting-platform/extension-plan') {
    return Response.json(host.extensionInstallPlan(options));
  }
  if (url.pathname === '/api/meeting-platform/integration-runtime') {
    return Response.json(host.integrationRuntimeSummary(options));
  }
  if (url.pathname === '/api/meeting-platform/integration-runtime/manifest') {
    return Response.json(host.integrationRuntimeManifest(options));
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

function readmeSource(plan = {}) {
  const platforms = plan.platforms ?? [];
  return `# Meeting Platform Timeline Host

This scaffold wires a host project to @ai-annotation/meeting-timeline-sdk.

Runtime rule:

- Realtime annotation marks use captured_at_ms.
- Local meeting app observation may create the meeting axis before provider events arrive.
- Provider webhooks reconcile or backfill the axis; they do not block realtime annotation.
- Transcript and artifact import are post-meeting work.
- Each meeting keeps an isolated annotation set.

Supported platforms in this scaffold:

${platforms.map((platform) => `- ${platform}`).join('\n')}

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
const strategy = host.adaptationStrategyMatrix();
const platformResolution = host.resolvePlatform({
  url: 'https://meet.google.com/abc-defg-hij',
});
const runtimeBundles = host.runtimeBundles();
const runtimeEventPlans = host.runtimeEventPlans();
const extensionPlan = host.extensionInstallPlan();
const integrationRuntime = host.integrationRuntimeSummary();
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
npm run meeting-platform:runtime-bundles
npm run meeting-platform:runtime-event-plans
npm run meeting-platform:resolve
npm run meeting-platform:extension-plan
npm run meeting-platform:integration-runtime
npm run meeting-platform:integration-runtime-manifest
npm run sdk:package-smoke
\`\`\`
`;
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
  const runtimeBundleMatrix = buildMeetingPlatformRuntimeBundleMatrix({
    ...options,
    platforms,
    baseUrl,
  });
  const runtimeEventPlanMatrix = buildMeetingPlatformRuntimeEventPlanMatrix({
    ...options,
    platforms,
    baseUrl,
  });
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
    },
    runtime_contract: {
      annotation_timestamp_field: 'captured_at_ms',
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      per_meeting_annotation_isolation_required: true,
      local_observer_may_start_axis: true,
    },
    endpoints: {
      platform_events: basePath,
      annotations: '/api/annotations',
      readiness: '/api/meeting-platform/readiness',
      handoff: '/api/meeting-platform/handoff',
      adapter_contracts: '/api/meeting-platform/contracts',
      adapter_contract_acceptance: '/api/meeting-platform/contract-acceptance',
      runtime_bundles: '/api/meeting-platform/runtime-bundles',
      runtime_event_plans: '/api/meeting-platform/runtime-event-plans',
      strategy: '/api/meeting-platform/strategy',
      platform_resolution: '/api/meeting-platform/resolve',
      extension_plan: '/api/meeting-platform/extension-plan',
      integration_runtime: '/api/meeting-platform/integration-runtime',
      integration_runtime_manifest: '/api/meeting-platform/integration-runtime/manifest',
      runtime_events: '/api/meeting-platform/runtime-events',
    },
    commands: handoff.commands,
    evidence_paths: handoff.evidence_paths,
    integration_plans: Object.fromEntries(platforms.map((platform) => [
      platform,
      buildPlatformIntegrationPlan(platform, { ...options, baseUrl, basePath }),
    ])),
    extension_install_plan: extensionInstallPlan,
    adaptation_strategy_matrix: adaptationStrategyMatrix,
    runtime_event_plan_matrix: runtimeEventPlanMatrix,
    runtime_bundle_matrix: runtimeBundleMatrix,
    handoff_bundle: handoff,
    adapter_contract_matrix: adapterContractMatrix,
    adapter_contract_acceptance_matrix: adapterContractAcceptanceMatrix,
    next_actions: unique([
      ...(handoff.next_actions ?? []),
      'run_meeting_platform_contract_acceptance_before_enabling_new_platform',
      'wire_host_routes_to_handleMeetingPlatformRequest',
      'capture_real_meeting_app_snapshots_for_each_target_platform',
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
      'meeting-platform:runtime-bundles': 'node ./scripts/print-runtime-bundles.mjs',
      'meeting-platform:runtime-event-plans': 'node ./scripts/print-runtime-event-plans.mjs',
      'meeting-platform:extension-plan': 'node ./scripts/print-extension-plan.mjs',
      'meeting-platform:resolve': 'node ./scripts/resolve-platform.mjs',
      'meeting-platform:integration-runtime': 'node ./scripts/print-integration-runtime.mjs',
      'meeting-platform:integration-runtime-manifest': 'node ./scripts/print-integration-runtime-manifest.mjs',
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
  const runtimeBundlesScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

console.log(JSON.stringify(host.runtimeBundles(), null, 2));
`;
  const runtimeEventPlansScript = `import { createMeetingPlatformHost } from '../src/meeting-platform-host.mjs';

const host = createMeetingPlatformHost({
  baseUrl: process.env.MEETING_TIMELINE_BASE_URL ?? ${JSON.stringify(plan.base_url)},
});

console.log(JSON.stringify(host.runtimeEventPlans(), null, 2));
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
      sourceFile('scripts/print-handoff.mjs', handoffScript, 'handoff_script', 'text/javascript'),
      sourceFile('scripts/print-strategy.mjs', strategyScript, 'adaptation_strategy_script', 'text/javascript'),
      sourceFile('scripts/print-readiness.mjs', readinessScript, 'readiness_script', 'text/javascript'),
      sourceFile('scripts/print-contracts.mjs', contractsScript, 'adapter_contract_script', 'text/javascript'),
      sourceFile('scripts/verify-contracts.mjs', contractAcceptanceScript, 'adapter_contract_acceptance_script', 'text/javascript'),
      sourceFile('scripts/print-runtime-bundles.mjs', runtimeBundlesScript, 'runtime_bundle_script', 'text/javascript'),
      sourceFile('scripts/print-runtime-event-plans.mjs', runtimeEventPlansScript, 'runtime_event_plan_script', 'text/javascript'),
      sourceFile('scripts/print-extension-plan.mjs', extensionPlanScript, 'extension_plan_script', 'text/javascript'),
      sourceFile('scripts/resolve-platform.mjs', platformResolutionScript, 'platform_resolution_script', 'text/javascript'),
      sourceFile('scripts/print-integration-runtime.mjs', integrationRuntimeScript, 'integration_runtime_script', 'text/javascript'),
      sourceFile('scripts/print-integration-runtime-manifest.mjs', integrationRuntimeManifestScript, 'integration_runtime_manifest_script', 'text/javascript'),
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
    'scripts/print-handoff.mjs',
    'scripts/print-strategy.mjs',
    'scripts/print-readiness.mjs',
    'scripts/print-contracts.mjs',
    'scripts/verify-contracts.mjs',
    'scripts/print-runtime-bundles.mjs',
    'scripts/print-runtime-event-plans.mjs',
    'scripts/resolve-platform.mjs',
    'scripts/print-extension-plan.mjs',
    'scripts/print-integration-runtime.mjs',
    'scripts/print-integration-runtime-manifest.mjs',
    'README.md',
  ];
  for (const path of requiredFiles) {
    if (!fileByPath(scaffold, path)) {
      issues.push(issue('error', 'missing_required_file', `Missing required scaffold file: ${path}`, { path }));
    }
  }
  const host = fileByPath(scaffold, 'src/meeting-platform-host.mjs')?.content ?? '';
  const routes = fileByPath(scaffold, 'src/http-routes.mjs')?.content ?? '';
  const readme = fileByPath(scaffold, 'README.md')?.content ?? '';
  if (!host.includes('createMeetingPlatformIntegrationRuntime')) {
    issues.push(issue('error', 'missing_integration_runtime_import', 'Host source must create the platform integration runtime.'));
  }
  if (!host.includes('integrationRuntimeManifest')) {
    issues.push(issue('error', 'missing_integration_runtime_manifest', 'Host source must expose the integration runtime manifest.'));
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
  if (!host.includes('platformRuntimeBundleMatrix')) {
    issues.push(issue('error', 'missing_runtime_bundle_matrix', 'Host source must expose the runtime bundle matrix.'));
  }
  if (!host.includes('platformRuntimeEventPlanMatrix')) {
    issues.push(issue('error', 'missing_runtime_event_plan_matrix', 'Host source must expose the runtime event plan matrix.'));
  }
  if (!host.includes('platformAdaptationStrategyMatrix')) {
    issues.push(issue('error', 'missing_adaptation_strategy_matrix', 'Host source must expose the adaptation strategy matrix.'));
  }
  if (!host.includes('resolvePlatform')) {
    issues.push(issue('error', 'missing_platform_resolution', 'Host source must expose platform resolution for URL/window based adapter selection.'));
  }
  if (!host.includes('meetingAppExtensionInstallPlan')) {
    issues.push(issue('error', 'missing_extension_install_plan', 'Host source must expose the extension install plan.'));
  }
  if (!host.includes('capturedAtMs')) {
    issues.push(issue('error', 'missing_captured_at_ms_write', 'Host source must write annotations with capturedAtMs.'));
  }
  if (!routes.includes('handleFetchRequest')) {
    issues.push(issue('error', 'missing_provider_route_handler', 'Route source must forward provider webhooks to the SDK handler.'));
  }
  if (!routes.includes('/api/meeting-platform/contracts')) {
    issues.push(issue('error', 'missing_contract_route', 'Route source must expose the adapter contract matrix endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/runtime-bundles')) {
    issues.push(issue('error', 'missing_runtime_bundle_route', 'Route source must expose the runtime bundle matrix endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/runtime-event-plans')) {
    issues.push(issue('error', 'missing_runtime_event_plan_route', 'Route source must expose the runtime event plan matrix endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/strategy')) {
    issues.push(issue('error', 'missing_adaptation_strategy_route', 'Route source must expose the adaptation strategy matrix endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/resolve')) {
    issues.push(issue('error', 'missing_platform_resolution_route', 'Route source must expose the platform resolution endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/extension-plan')) {
    issues.push(issue('error', 'missing_extension_plan_route', 'Route source must expose the extension install plan endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/integration-runtime')) {
    issues.push(issue('error', 'missing_integration_runtime_route', 'Route source must expose the integration runtime summary endpoint.'));
  }
  if (!routes.includes('/api/meeting-platform/runtime-events')) {
    issues.push(issue('error', 'missing_runtime_events_route', 'Route source must expose the runtime event endpoint.'));
  }
  if (!readme.includes('captured_at_ms')) {
    issues.push(issue('warning', 'readme_missing_timestamp_contract', 'README should state the captured_at_ms contract.'));
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
