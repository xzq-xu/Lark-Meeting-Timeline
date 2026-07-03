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
  return `import { createMeetingPlatformTimelineKit } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit';
import { createHostTimelineClient } from './timeline-client.mjs';

const DEFAULT_PLATFORMS = ${defaultPlatforms};

export function createMeetingPlatformHost(options = {}) {
  const client = options.client ?? createHostTimelineClient({
    baseUrl: options.baseUrl ?? options.base_url ?? ${JSON.stringify(defaultBaseUrl)},
    fetch: options.fetch,
    headers: options.headers,
  });
  const platforms = options.platforms ?? options.platform_keys ?? DEFAULT_PLATFORMS;
  const kit = createMeetingPlatformTimelineKit(client, {
    ...options,
    platforms,
  });
  const liveAdapters = kit.platformLiveAdapterSuite({ platforms });

  return {
    client,
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
    adapter(platform, adapterOptions = {}) {
      return liveAdapters.adapter(platform, adapterOptions);
    },
    observeMeetingApp(platform, snapshot = {}, observeOptions = {}) {
      return liveAdapters.adapter(platform, observeOptions.adapterOptions ?? {}).observeMeetingApp(snapshot, observeOptions);
    },
    ingestProvider(platform, requestOrPayload = {}, payload, ingestOptions = {}) {
      return liveAdapters.adapter(platform, ingestOptions.adapterOptions ?? {}).ingestProvider(requestOrPayload, payload, ingestOptions);
    },
    insertAnnotation(platform, input = {}, markOptions = {}) {
      return liveAdapters.adapter(platform, markOptions.adapterOptions ?? {}).insertAnnotation({
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
\`\`\`

Verification:

\`\`\`sh
npm run meeting-app:extension:build
npm run meeting-app:evidence-matrix
npm run meeting-platform:rollout-matrix
npm run meeting-platform:live-readiness
npm run meeting-platform:contracts
npm run meeting-platform:contract-acceptance
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
    },
    commands: handoff.commands,
    evidence_paths: handoff.evidence_paths,
    integration_plans: Object.fromEntries(platforms.map((platform) => [
      platform,
      buildPlatformIntegrationPlan(platform, { ...options, baseUrl, basePath }),
    ])),
    extension_install_plan: extensionInstallPlan,
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
      'meeting-platform:readiness': 'node ./scripts/print-readiness.mjs',
      'meeting-platform:contracts': 'node ./scripts/print-contracts.mjs',
      'meeting-platform:contract-acceptance': 'node ./scripts/verify-contracts.mjs',
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
      sourceFile('scripts/print-readiness.mjs', readinessScript, 'readiness_script', 'text/javascript'),
      sourceFile('scripts/print-contracts.mjs', contractsScript, 'adapter_contract_script', 'text/javascript'),
      sourceFile('scripts/verify-contracts.mjs', contractAcceptanceScript, 'adapter_contract_acceptance_script', 'text/javascript'),
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
    'scripts/print-readiness.mjs',
    'scripts/print-contracts.mjs',
    'scripts/verify-contracts.mjs',
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
  if (!host.includes('createMeetingPlatformTimelineKit')) {
    issues.push(issue('error', 'missing_platform_kit_import', 'Host source must create the platform kit.'));
  }
  if (!host.includes('platformLiveAdapterSuite')) {
    issues.push(issue('error', 'missing_live_adapter_suite', 'Host source must create the live adapter suite.'));
  }
  if (!host.includes('platformAdapterContractMatrix')) {
    issues.push(issue('error', 'missing_adapter_contract_matrix', 'Host source must expose the adapter contract matrix.'));
  }
  if (!host.includes('platformAdapterContractAcceptanceMatrix')) {
    issues.push(issue('error', 'missing_adapter_contract_acceptance_matrix', 'Host source must expose the adapter contract acceptance matrix.'));
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
