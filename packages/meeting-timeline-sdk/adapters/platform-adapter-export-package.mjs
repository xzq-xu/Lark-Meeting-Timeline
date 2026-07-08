import { compactObject } from '../index.mjs';
import {
  buildMeetingPlatformAdapterAuthoringPlan,
} from './platform-adapter-authoring.mjs';
import {
  buildMeetingPlatformAdapterPortfolio,
  buildMeetingPlatformAdapterPortfolioItem,
} from './platform-adapter-portfolio.mjs';
import {
  buildMeetingPlatformAdapterBlueprint,
} from './platform-adapter-blueprint.mjs';
import {
  buildMeetingPlatformAdapterAcceptanceChecklist,
  buildMeetingPlatformAdapterAcceptanceChecklistMatrix,
} from './platform-adapter-acceptance-checklist.mjs';
import {
  buildMeetingPlatformImplementationHandoff,
} from './platform-implementation-handoff.mjs';
import {
  buildMeetingPlatformRuntimeBundle,
} from './platform-runtime-bundle.mjs';
import {
  buildMeetingPlatformProviderConnectionPack,
} from './platform-provider-connection.mjs';
import {
  buildMeetingPlatformAdapterContract,
} from './platform-adapter-contract.mjs';
import {
  buildMeetingPlatformAdaptationPackage,
} from './platform-adaptation-package.mjs';

export const MEETING_PLATFORM_ADAPTER_EXPORT_PACKAGE_SCHEMA = 'meeting_platform_adapter_export_package';
export const MEETING_PLATFORM_ADAPTER_EXPORT_PACKAGE_MATRIX_SCHEMA = 'meeting_platform_adapter_export_package_matrix';
export const MEETING_PLATFORM_ADAPTER_EXPORT_PACKAGE_SCHEMA_VERSION = 1;

const DEFAULT_EXPORT_PLATFORMS = Object.freeze([
  'google-meet',
  'teams',
  'zoom',
  'webex',
  'lark',
]);

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
  return unique(asArray(firstNonEmpty(
    options.platforms,
    options.platform_keys,
    DEFAULT_EXPORT_PLATFORMS,
  )));
}

function boolOption(options = {}, camel, snake, defaultValue = false) {
  return firstNonEmpty(options[camel], options[snake], defaultValue) === true;
}

function targetFrom(input = {}, options = {}) {
  const target = String(firstNonEmpty(
    options.target,
    options.acceptanceTarget,
    options.acceptance_target,
    input.target,
    input.acceptanceTarget,
    input.acceptance_target,
    'static',
  ));
  return ['static', 'pilot', 'production'].includes(target) ? target : 'static';
}

function commandWithBase(name, options = {}, extra = '') {
  const baseUrl = firstNonEmpty(options.baseUrl, options.base_url);
  const args = [
    baseUrl ? `--base-url=${baseUrl}` : '',
    String(extra).trim(),
  ].filter(Boolean).join(' ');
  return args ? `npm run ${name} -- ${args}` : `npm run ${name}`;
}

function safeArtifact(builtIn, build, platform, options = {}) {
  if (!builtIn) return undefined;
  try {
    return build(platform, options);
  } catch {
    return undefined;
  }
}

function artifactRef(name, value = {}, path, requiredFrom = 'static') {
  return compactObject({
    name,
    schema: value.schema,
    type: value.type,
    path,
    required_from: requiredFrom,
    platform: value.platform,
  });
}

function packagePath(platform, file) {
  return `${platform}/${file}`;
}

function buildHostFiles(platform, artifacts = {}, builtIn = false) {
  const files = [
    {
      path: packagePath(platform, 'adapter-export-package.json'),
      source: 'adapter_export_package',
      required_from: 'static',
      purpose: 'single entry index for importing this platform adapter into a host project',
    },
    {
      path: packagePath(platform, 'authoring-plan.json'),
      source: 'authoring_plan',
      required_from: 'static',
      purpose: 'adapter scope, required contracts, first local surface, and missing authoring work',
    },
    {
      path: packagePath(platform, 'acceptance-checklist.json'),
      source: 'acceptance_checklist',
      required_from: 'static',
      purpose: 'static/pilot/production acceptance gate for this adapter',
    },
  ];
  if (builtIn || artifacts.implementation_handoff) {
    files.push({
      path: packagePath(platform, 'implementation-handoff.json'),
      source: 'implementation_handoff',
      required_from: 'static',
      purpose: 'step-by-step host integration flow and SDK entrypoints',
    });
  }
  if (builtIn || artifacts.adapter_blueprint) {
    files.push({
      path: packagePath(platform, 'adapter-blueprint.json'),
      source: 'adapter_blueprint',
      required_from: 'static',
      purpose: 'surface-level adapter contract for browser/native/provider/artifact wiring',
    });
  }
  if (builtIn || artifacts.runtime_bundle) {
    files.push({
      path: packagePath(platform, 'runtime-bundle.json'),
      source: 'runtime_bundle',
      required_from: 'static',
      purpose: 'browser/WebView/native observer runtime configuration and message contract',
    });
  }
  if (builtIn || artifacts.adapter_contract) {
    files.push({
      path: packagePath(platform, 'adapter-contract.json'),
      source: 'adapter_contract',
      required_from: 'static',
      purpose: 'hard realtime timeline contract, timestamp contract, and provider nonblocking rules',
    });
  }
  if (builtIn || artifacts.provider_connection) {
    files.push({
      path: packagePath(platform, 'provider-connection.json'),
      source: 'provider_connection',
      required_from: 'production',
      purpose: 'official provider event mapping, security verifier, docs, and subscription setup',
    });
  }
  if (builtIn || artifacts.adaptation_package) {
    files.push({
      path: packagePath(platform, 'adaptation-package.json'),
      source: 'adaptation_package',
      required_from: 'static',
      purpose: 'portable strategy package for realtime axis plus provider reconcile/backfill',
    });
  }
  return files;
}

function importPaths() {
  return {
    sdk_root: '@ai-annotation/meeting-timeline-sdk',
    platform_kit: '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit',
    adapter_export_package: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-export-package',
    adapter_authoring: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-authoring',
    adapter_portfolio: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-portfolio',
    adapter_blueprint: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-blueprint',
    adapter_preflight: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-preflight',
    adapter_acceptance_checklist: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-acceptance-checklist',
    implementation_handoff: '@ai-annotation/meeting-timeline-sdk/adapters/platform-implementation-handoff',
    runtime_bundle: '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-bundle',
    adapter_contract: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-contract',
    provider_connection: '@ai-annotation/meeting-timeline-sdk/adapters/platform-provider-connection',
    adaptation_package: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adaptation-package',
    integration_runtime: '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime',
    runtime_host: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-runtime-host',
    connector: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-connector',
  };
}

function commandSet(platform, target, options = {}) {
  return {
    export_package: commandWithBase('meeting-platform:adapter-export-package', options, `--platforms=${platform} --target=${target}`),
    adapter_authoring: commandWithBase('meeting-platform:adapter-authoring', options, `--platforms=${platform}`),
    adapter_portfolio: commandWithBase('meeting-platform:adapter-portfolio', options, `--platforms=${platform}`),
    adapter_blueprint: commandWithBase('meeting-platform:adapter-blueprint', options, `--platforms=${platform}`),
    adapter_preflight: commandWithBase('meeting-platform:adapter-preflight', options, `--platforms=${platform}`),
    acceptance_checklist: commandWithBase('meeting-platform:adapter-acceptance-checklist', options, `--platforms=${platform} --target=${target}`),
    implementation_handoff: commandWithBase('meeting-platform:implementation-handoff', options, `--platforms=${platform}`),
    runtime_bundle: commandWithBase('meeting-platform:runtime-bundle', options, `--platforms=${platform}`),
    provider_connection: commandWithBase('meeting-platform:provider-connection', options, `--platforms=${platform}`),
    handoff_readiness: commandWithBase('meeting-platform:handoff-readiness', options, `--platforms=${platform}`),
    runtime_host_verify: commandWithBase('meeting-platform:runtime-host-verify', options, `--platforms=${platform}`),
    sdk_test: 'npm run sdk:test',
  };
}

function browserSurface(runtimeBundle = {}, portfolioItem = {}) {
  const p0 = portfolioItem.p0_realtime_axis ?? {};
  const browser = runtimeBundle.browser ?? {};
  return compactObject({
    ready: (browser.matches?.length ?? p0.browser_match_count ?? 0) > 0,
    role: 'lowest_latency_local_axis_and_realtime_annotation_surface',
    matches: browser.matches ?? p0.browser_matches ?? [],
    permissions: browser.permissions ?? [],
    host_permissions: browser.host_permissions ?? [],
    content_script_manifest_source: 'runtime-bundle.json#/browser/manifest',
    first_message_type: p0.candidate_observation_message_type ?? 'meeting_timeline.observe_candidates',
    preflight_required_before_insert: true,
    preflight_message_types: [
      'meeting_timeline.preflight_current_window',
      'meeting_timeline.preflight_candidates',
    ],
    first_sdk_method: 'observePlatformCandidates',
    mark_insert_method: 'insertAnnotation',
    timestamp_field: p0.timestamp_field ?? 'captured_at_ms',
  });
}

function webviewSurface(runtimeBundle = {}, portfolioItem = {}) {
  const bridge = runtimeBundle.runtime?.lightweight_connector_bridge ?? {};
  const browserReady = (runtimeBundle.browser?.matches?.length ?? portfolioItem.p0_realtime_axis?.browser_match_count ?? 0) > 0;
  return compactObject({
    ready: browserReady,
    role: 'embedded_web_or_electron_preload_bridge_for_meeting_ui',
    install_function: bridge.install_function,
    options_source: 'runtime-bundle.json#/runtime/lightweight_connector_bridge/options',
    window_messaging_supported: true,
    preflight_required_before_insert: true,
    preflight_message_types: [
      'meeting_timeline.preflight_current_window',
      'meeting_timeline.preflight_candidates',
    ],
    first_sdk_method: 'observeMeetingApp_or_observePlatformCandidates',
    timestamp_field: 'captured_at_ms',
  });
}

function nativeSurface(portfolioItem = {}, implementation = {}) {
  return compactObject({
    ready: Boolean(portfolioItem.recommended_first_surface),
    role: 'host_controlled_detector_or_desktop_observer_when_browser_extension_is_not_available',
    recommended_first_surface: portfolioItem.recommended_first_surface,
    first_sdk_method: 'observePlatformCandidates',
    preflight_required_before_insert: true,
    preflight_sdk_method: 'platformAdapterPreflight',
    mark_insert_method: 'insertAnnotation',
    speaker_track_method: 'speakerTrack',
    participant_track_method: 'participantTrack',
    flow_source: implementation.flow?.length ? 'implementation-handoff.json#/flow' : undefined,
    timestamp_field: 'captured_at_ms',
  });
}

function providerSurface(providerConnection = {}, portfolioItem = {}) {
  const p1 = portfolioItem.p1_provider_reconcile ?? {};
  return compactObject({
    ready: Boolean(providerConnection.transport ?? p1.transport),
    role: 'reconcile_and_post_meeting_backfill_only',
    path: p1.path,
    transport: providerConnection.transport ?? p1.transport,
    required_for_realtime: false,
    blocks_realtime_annotation: false,
    official_docs: p1.official_docs,
    event_mapping_count: providerConnection.event_mapping?.length ?? p1.event_mapping_count,
    security: providerConnection.security ?? p1.security,
    first_sdk_method: 'ingestProvider',
  });
}

function lightweightPortfolioItem(plan = {}) {
  const contracts = plan.required_contracts ?? {};
  const browser = plan.browser_surface ?? {};
  const provider = plan.provider_reconcile ?? {};
  return compactObject({
    type: 'meeting_platform_adapter_portfolio_item',
    schema: 'meeting_platform_adapter_portfolio_item',
    schema_version: 1,
    platform: plan.platform,
    display_name: plan.display_name,
    built_in: plan.built_in === true,
    adapter_status: plan.built_in ? 'built_in_adapter_available' : 'adapter_authoring_required',
    recommended_first_surface: plan.recommended_first_surface,
    p0_realtime_axis: {
      source: plan.recommended_first_surface,
      local_axis_first: contracts.local_axis_first === true,
      browser_matches: browser.matches ?? [],
      browser_match_count: browser.matches?.length ?? 0,
      candidate_observation_required: browser.candidate_observation_required === true,
      candidate_observation_message_type: contracts.candidate_observation_message_type,
      timestamp_field: contracts.timestamp_field,
      provider_events_block_realtime: contracts.provider_events_block_realtime === true,
      transcript_blocks_realtime: contracts.transcript_blocks_realtime === true,
    },
    p1_provider_reconcile: compactObject({
      path: provider.path,
      transport: provider.transport,
      role: provider.role,
      required_for_realtime: provider.required_for_realtime === true,
    }),
    p2_post_meeting_backfill: {
      realtime_dependency: false,
      transcript_blocks_realtime: contracts.transcript_blocks_realtime === true,
      supported: plan.built_in === true,
    },
    implementation: {
      implementation_ready: plan.built_in === true,
      pilot_ready: false,
      production_ready: false,
      recommended_first_surface: plan.recommended_first_surface,
    },
    commands: compactObject({
      adapter_authoring: plan.commands?.authoring_plan,
      implementation_handoff: plan.commands?.implementation_handoff,
      sdk_test: plan.commands?.sdk_test,
    }),
    next_actions: plan.next_actions ?? [],
  });
}

function adapterPreflightExport(implementation = {}, portfolioItem = {}) {
  const preflight = implementation.adapter_preflight ?? {};
  return compactObject({
    status: preflight.status ?? 'needs_live_page_evidence',
    accepted: preflight.accepted === true,
    selected_surface: preflight.selected_surface ?? portfolioItem.recommended_first_surface,
    startup_ready: preflight.startup_ready === true || portfolioItem.implementation?.implementation_ready === true,
    live_evidence_ready: preflight.live_evidence_ready === true,
    realtime_annotation_ready: preflight.realtime_annotation_ready === true,
    required_before: 'bind_current_axis_or_insert_realtime_annotation',
    live_evidence_required: true,
    url_only_status: 'needs_live_page_evidence',
    source: implementation.schema === 'meeting_platform_implementation_handoff'
      ? 'implementation_handoff'
      : 'export_package_default_gate',
    sdk_methods: preflight.sdk_methods ?? [
      'platformAdapterPreflight',
      'platformAdapterCurrentWindowPreflight',
      'platformAdapterCandidatePreflight',
    ],
    bridge_messages: preflight.bridge_messages ?? [
      'meeting_timeline.preflight_current_window',
      'meeting_timeline.preflight_candidates',
    ],
    command: implementation.acceptance?.commands?.adapter_preflight,
    first_next_action: preflight.first_next_action ?? 'collect_live_dom_or_native_window_evidence_for_adapter_preflight',
  });
}

function lightweightChecklistItem(target, input = {}) {
  const passed = input.passed === true;
  return compactObject({
    id: input.id,
    group: input.group,
    label: input.label,
    required_from: 'static',
    required: true,
    passed,
    status: passed ? 'pass' : 'fail',
    blocking: !passed,
    expected: input.expected,
    actual: input.actual,
    next_action: input.next_action,
    target,
  });
}

function lightweightAcceptanceChecklist(plan = {}, portfolioItem = {}, target = 'static') {
  const p0 = portfolioItem.p0_realtime_axis ?? {};
  const p1 = portfolioItem.p1_provider_reconcile ?? {};
  const checks = [
    lightweightChecklistItem(target, {
      id: 'adapter_registered_or_authorable',
      group: 'static_contract',
      label: 'Platform adapter is registered or has an authoring plan',
      passed: plan.built_in === true,
      expected: 'built_in_adapter_available',
      actual: portfolioItem.adapter_status,
      next_action: plan.built_in ? undefined : 'add_platform_setup_entry',
    }),
    lightweightChecklistItem(target, {
      id: 'p0_local_axis_first',
      group: 'p0_realtime_axis',
      label: 'Realtime annotations use a local axis before provider reconcile',
      passed: p0.local_axis_first === true,
      expected: true,
      actual: p0.local_axis_first,
      next_action: 'fix_adapter_route_local_axis_first',
    }),
    lightweightChecklistItem(target, {
      id: 'captured_at_ms_contract',
      group: 'p0_realtime_axis',
      label: 'Realtime annotation timestamp field is captured_at_ms',
      passed: p0.timestamp_field === 'captured_at_ms',
      expected: 'captured_at_ms',
      actual: p0.timestamp_field,
      next_action: 'fix_annotation_timestamp_contract',
    }),
    lightweightChecklistItem(target, {
      id: 'provider_nonblocking_contract',
      group: 'p0_realtime_axis',
      label: 'Provider events never block realtime annotation insertion',
      passed: p0.provider_events_block_realtime === false,
      expected: false,
      actual: p0.provider_events_block_realtime,
      next_action: 'set_provider_events_block_realtime_false',
    }),
    lightweightChecklistItem(target, {
      id: 'transcript_nonblocking_contract',
      group: 'p2_post_meeting_backfill',
      label: 'Transcript and recording artifacts are post-meeting backfill only',
      passed: p0.transcript_blocks_realtime === false,
      expected: false,
      actual: p0.transcript_blocks_realtime,
      next_action: 'set_transcript_blocks_realtime_false',
    }),
    lightweightChecklistItem(target, {
      id: 'candidate_observation_contract',
      group: 'p0_realtime_axis',
      label: 'Host can observe current meeting candidates',
      passed: p0.candidate_observation_message_type === 'meeting_timeline.observe_candidates',
      expected: 'meeting_timeline.observe_candidates',
      actual: p0.candidate_observation_message_type,
      next_action: 'wire_observe_platform_candidates',
    }),
    lightweightChecklistItem(target, {
      id: 'adapter_preflight_gate_declared',
      group: 'p0_realtime_axis',
      label: 'Adapter preflight gate is declared before realtime mark insertion',
      passed: true,
      expected: 'needs_live_page_evidence_until_live_dom_or_native_window_sample',
      actual: 'needs_live_page_evidence_until_live_dom_or_native_window_sample',
      next_action: 'run_adapter_preflight_with_live_window_evidence_before_first_insert',
    }),
    lightweightChecklistItem(target, {
      id: 'provider_reconcile_path_declared',
      group: 'p1_provider_reconcile',
      label: 'Provider reconcile path is declared',
      passed: Boolean(p1.path),
      actual: p1.path,
      next_action: 'declare_provider_reconcile_path',
    }),
    lightweightChecklistItem(target, {
      id: 'implementation_handoff_available',
      group: 'static_contract',
      label: 'Implementation handoff is available for the platform',
      passed: plan.built_in === true,
      expected: true,
      actual: plan.built_in === true,
      next_action: 'run_meeting_platform_implementation_handoff',
    }),
  ];
  const failedRequiredIds = checks.filter((item) => item.blocking).map((item) => item.id);
  return {
    type: 'meeting_platform_adapter_acceptance_checklist',
    schema: 'meeting_platform_adapter_acceptance_checklist',
    schema_version: 1,
    platform: plan.platform,
    display_name: plan.display_name,
    target,
    accepted: failedRequiredIds.length === 0,
    adapter_status: portfolioItem.adapter_status,
    recommended_first_surface: plan.recommended_first_surface,
    checklist: checks,
    summary: {
      blocking_count: failedRequiredIds.length,
      failed_required_ids: failedRequiredIds,
      static_ready: failedRequiredIds.length === 0,
      pilot_ready: false,
      production_ready: false,
    },
    commands: compactObject({
      adapter_authoring: plan.commands?.authoring_plan,
      implementation_handoff: plan.commands?.implementation_handoff,
      sdk_test: plan.commands?.sdk_test,
    }),
    next_actions: unique([
      ...(plan.next_actions ?? []),
      ...checks.filter((item) => item.blocking && item.next_action).map((item) => item.next_action),
    ]),
  };
}

function artifactRefs(artifacts = {}, files = []) {
  const bySource = Object.fromEntries(files.map((file) => [file.source, file.path]));
  return compactObject({
    authoring_plan: artifactRef('authoring_plan', artifacts.authoring_plan, bySource.authoring_plan, 'static'),
    portfolio_item: artifactRef('portfolio_item', artifacts.portfolio_item, undefined, 'static'),
    acceptance_checklist: artifactRef('acceptance_checklist', artifacts.acceptance_checklist, bySource.acceptance_checklist, 'static'),
    implementation_handoff: artifactRef('implementation_handoff', artifacts.implementation_handoff, bySource.implementation_handoff, 'static'),
    adapter_blueprint: artifactRef('adapter_blueprint', artifacts.adapter_blueprint, bySource.adapter_blueprint, 'static'),
    runtime_bundle: artifactRef('runtime_bundle', artifacts.runtime_bundle, bySource.runtime_bundle, 'static'),
    adapter_contract: artifactRef('adapter_contract', artifacts.adapter_contract, bySource.adapter_contract, 'static'),
    provider_connection: artifactRef('provider_connection', artifacts.provider_connection, bySource.provider_connection, 'production'),
    adaptation_package: artifactRef('adaptation_package', artifacts.adaptation_package, bySource.adaptation_package, 'static'),
  });
}

function builtInArtifactRefs(platform, files = [], refs = {}) {
  const bySource = Object.fromEntries(files.map((file) => [file.source, file.path]));
  return compactObject({
    ...refs,
    implementation_handoff: refs.implementation_handoff ?? {
      name: 'implementation_handoff',
      schema: 'meeting_platform_implementation_handoff',
      path: bySource.implementation_handoff,
      required_from: 'static',
      platform,
    },
    adapter_blueprint: refs.adapter_blueprint ?? {
      name: 'adapter_blueprint',
      schema: 'meeting_platform_adapter_blueprint',
      path: bySource.adapter_blueprint,
      required_from: 'static',
      platform,
    },
    runtime_bundle: refs.runtime_bundle ?? {
      name: 'runtime_bundle',
      schema: 'meeting_platform_runtime_bundle',
      path: bySource.runtime_bundle,
      required_from: 'static',
      platform,
    },
    adapter_contract: refs.adapter_contract ?? {
      name: 'adapter_contract',
      schema: 'meeting_platform_adapter_contract',
      path: bySource.adapter_contract,
      required_from: 'static',
      platform,
    },
    provider_connection: refs.provider_connection ?? {
      name: 'provider_connection',
      schema: 'meeting_platform_provider_connection_pack',
      path: bySource.provider_connection,
      required_from: 'production',
      platform,
    },
    adaptation_package: refs.adaptation_package ?? {
      name: 'adaptation_package',
      schema: 'meeting_platform_adaptation_package',
      path: bySource.adaptation_package,
      required_from: 'static',
      platform,
    },
  });
}

function setupOrder(platform, builtIn, target) {
  return [
    {
      step: 1,
      id: 'load_export_package',
      action: 'read_adapter_export_package_and_host_files',
      required: true,
      output: 'platform adapter import plan',
    },
    {
      step: 2,
      id: 'wire_sdk_facade',
      action: 'createMeetingAppTimelineSdk_or_createMeetingPlatformTimelineKit',
      required: true,
      output: 'SDK object with observePlatformCandidates/insertAnnotation/speakerTrack',
    },
    {
      step: 3,
      id: 'install_local_surface',
      action: builtIn ? 'install_browser_extension_webview_or_native_detector_from_runtime_bundle' : 'author_first_local_surface_from_authoring_plan',
      required: true,
      output: 'meeting candidate observation before marks',
    },
    {
      step: 4,
      id: 'run_adapter_preflight',
      action: 'run adapter preflight with live DOM/native-window evidence before opening a realtime annotation session',
      required: true,
      output: 'adapter_preflight.realtime_annotation_ready',
      url_only_status: 'needs_live_page_evidence',
    },
    {
      step: 5,
      id: 'insert_realtime_marks',
      action: `call insertAnnotation('${platform}', { captured_at_ms, ...mark })`,
      required: true,
      output: 'timeline marks aligned to device capture time',
    },
    {
      step: 6,
      id: 'emit_speaker_positions',
      action: 'call speakerTrack/participantTrack when local evidence is available',
      required: false,
      output: 'speaker and participant position markers without transcript text',
    },
    {
      step: 7,
      id: 'connect_provider_reconcile',
      action: 'ingest official provider events only after local realtime axis is available',
      required: target === 'production',
      output: 'start/end/participant/artifact reconcile and backfill',
    },
    {
      step: 8,
      id: 'run_acceptance',
      action: `run adapter acceptance checklist target=${target}`,
      required: true,
      output: 'static/pilot/production readiness gate',
    },
  ];
}

function nextActionsFor(packageObject = {}) {
  const checklist = packageObject.acceptance_checklist ?? {};
  const portfolio = packageObject.portfolio_item ?? {};
  const actions = unique([
    ...(checklist.next_actions ?? []),
    ...(portfolio.next_actions ?? []),
    packageObject.built_in ? 'wire_export_package_into_host_project' : 'author_missing_platform_adapter_before_runtime_install',
    'run_adapter_preflight_with_live_window_evidence_before_first_insert',
    'bind_current_meeting_axis_before_first_annotation',
    'keep_provider_events_nonblocking_for_realtime_marks',
  ]);
  if (packageObject.accepted === true) {
    return actions.filter((action) => !String(action).startsWith('fix_'));
  }
  return actions;
}

export function buildMeetingPlatformAdapterExportPackage(platform, input = {}, options = {}) {
  const target = targetFrom(input, options);
  const merged = {
    ...input,
    ...options,
    target,
  };
  const includeArtifacts = boolOption(merged, 'includeArtifacts', 'include_artifacts', false);
  const authoringPlan = buildMeetingPlatformAdapterAuthoringPlan(platform, merged);
  const key = authoringPlan.platform;
  const buildArtifacts = includeArtifacts || boolOption(merged, 'buildArtifacts', 'build_artifacts', false);
  const buildFullCompanions = buildArtifacts || boolOption(merged, 'fullCompanions', 'full_companions', false);
  const portfolioItem = buildMeetingPlatformAdapterPortfolioItem(platform, merged);
  const checklist = buildFullCompanions
    ? buildMeetingPlatformAdapterAcceptanceChecklist(platform, input, merged)
    : lightweightAcceptanceChecklist(authoringPlan, portfolioItem, target);
  const builtIn = portfolioItem.built_in === true;
  const artifacts = compactObject({
    authoring_plan: authoringPlan,
    portfolio_item: portfolioItem,
    acceptance_checklist: checklist,
    implementation_handoff: buildArtifacts ? safeArtifact(builtIn, buildMeetingPlatformImplementationHandoff, key, merged) : undefined,
    adapter_blueprint: buildArtifacts ? safeArtifact(builtIn, buildMeetingPlatformAdapterBlueprint, key, merged) : undefined,
    runtime_bundle: buildArtifacts ? safeArtifact(builtIn, buildMeetingPlatformRuntimeBundle, key, merged) : undefined,
    provider_connection: buildArtifacts ? safeArtifact(builtIn, buildMeetingPlatformProviderConnectionPack, key, merged) : undefined,
    adapter_contract: buildArtifacts ? safeArtifact(builtIn, buildMeetingPlatformAdapterContract, key, merged) : undefined,
    adaptation_package: buildArtifacts ? safeArtifact(builtIn, buildMeetingPlatformAdaptationPackage, key, merged) : undefined,
  });
  const implementationForGate = artifacts.implementation_handoff
    ?? safeArtifact(builtIn, buildMeetingPlatformImplementationHandoff, key, merged);
  const adapterPreflight = adapterPreflightExport(implementationForGate, portfolioItem);
  const hostFiles = buildHostFiles(authoringPlan.platform, artifacts, builtIn);
  const refs = builtIn
    ? builtInArtifactRefs(authoringPlan.platform, hostFiles, artifactRefs(artifacts, hostFiles))
    : artifactRefs(artifacts, hostFiles);
  const base = compactObject({
    type: 'meeting_platform_adapter_export_package',
    schema: MEETING_PLATFORM_ADAPTER_EXPORT_PACKAGE_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_EXPORT_PACKAGE_SCHEMA_VERSION,
    platform: authoringPlan.platform,
    display_name: authoringPlan.display_name,
    built_in: builtIn,
    target,
    accepted: checklist.accepted === true,
    export_ready: builtIn && checklist.summary?.blocking_count === 0,
    package_role: 'portable_adapter_handoff_for_external_host_project',
    recommended_first_surface: portfolioItem.recommended_first_surface,
    local_axis_first: true,
    timestamp_field: 'captured_at_ms',
    provider_events_block_realtime: false,
    transcript_blocks_realtime: false,
    adapter_preflight: adapterPreflight,
    import_paths: importPaths(),
    artifact_refs: refs,
    host_files: hostFiles,
    surface_entrypoints: {
      browser_extension: browserSurface(artifacts.runtime_bundle, portfolioItem),
      webview_preload: webviewSurface(artifacts.runtime_bundle, portfolioItem),
      native_detector: nativeSurface(portfolioItem, artifacts.implementation_handoff),
      native_host: nativeSurface(portfolioItem, artifacts.implementation_handoff),
      provider_reconcile: providerSurface(artifacts.provider_connection, portfolioItem),
    },
    setup_order: setupOrder(authoringPlan.platform, builtIn, target),
    commands: commandSet(authoringPlan.platform, target, merged),
    readiness: compactObject({
      adapter_status: portfolioItem.adapter_status,
      static_accepted: target === 'static'
        ? checklist.accepted === true
        : buildMeetingPlatformAdapterAcceptanceChecklist(platform, input, { ...merged, target: 'static' }).accepted === true,
      target_accepted: checklist.accepted === true,
      target_blocking_count: checklist.summary?.blocking_count,
      failed_required_ids: checklist.summary?.failed_required_ids ?? [],
      implementation_ready: portfolioItem.implementation?.implementation_ready === true,
      adapter_preflight_startup_ready: adapterPreflight.startup_ready === true,
      adapter_preflight_realtime_ready: adapterPreflight.realtime_annotation_ready === true,
      pilot_ready: portfolioItem.implementation?.pilot_ready === true,
      production_ready: portfolioItem.implementation?.production_ready === true,
    }),
    authoring_plan: includeArtifacts ? authoringPlan : undefined,
    portfolio_item: includeArtifacts ? portfolioItem : compactObject({
      platform: portfolioItem.platform,
      display_name: portfolioItem.display_name,
      adapter_status: portfolioItem.adapter_status,
      recommended_first_surface: portfolioItem.recommended_first_surface,
      adapter_surfaces: portfolioItem.adapter_surfaces,
      p0_realtime_axis: portfolioItem.p0_realtime_axis,
      p1_provider_reconcile: portfolioItem.p1_provider_reconcile,
      p2_post_meeting_backfill: portfolioItem.p2_post_meeting_backfill,
      implementation: portfolioItem.implementation,
    }),
    acceptance_checklist: includeArtifacts ? checklist : compactObject({
      platform: checklist.platform,
      target: checklist.target,
      accepted: checklist.accepted,
      summary: checklist.summary,
      commands: checklist.commands,
    }),
    artifacts: includeArtifacts ? artifacts : undefined,
  });
  return {
    ...base,
    next_actions: nextActionsFor(base),
  };
}

export function buildMeetingPlatformAdapterExportPackageMatrix(input = {}, options = {}) {
  const target = targetFrom(input, options);
  const merged = {
    ...options,
    ...input,
    target,
  };
  const platforms = selectedPlatforms(merged);
  const packages = platforms.map((platform) => buildMeetingPlatformAdapterExportPackage(platform, input, {
    ...merged,
    platforms: undefined,
    platform_keys: undefined,
  }));
  const includeCompanionMatrices = boolOption(merged, 'includeCompanionMatrices', 'include_companion_matrices', false);
  const portfolio = includeCompanionMatrices
    ? buildMeetingPlatformAdapterPortfolio({
      ...merged,
      platforms,
    })
    : {
      type: 'meeting_platform_adapter_portfolio',
      schema: 'meeting_platform_adapter_portfolio',
      platform_count: packages.length,
      built_in_count: packages.filter((item) => item.built_in === true).length,
      external_authoring_count: packages.filter((item) => item.built_in !== true).length,
      rows: packages.map((item) => ({
        platform: item.platform,
        display_name: item.display_name,
        built_in: item.built_in,
        recommended_first_surface: item.recommended_first_surface,
      })),
    };
  const checklistMatrix = includeCompanionMatrices
    ? buildMeetingPlatformAdapterAcceptanceChecklistMatrix({
      ...input,
      platforms,
    }, merged)
    : {
      type: 'meeting_platform_adapter_acceptance_checklist_matrix',
      schema: 'meeting_platform_adapter_acceptance_checklist_matrix',
      target,
      platform_count: packages.length,
      accepted_count: packages.filter((item) => item.accepted === true).length,
      blocked_count: packages.filter((item) => item.accepted !== true).length,
    };
  return {
    type: 'meeting_platform_adapter_export_package_matrix',
    schema: MEETING_PLATFORM_ADAPTER_EXPORT_PACKAGE_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_EXPORT_PACKAGE_SCHEMA_VERSION,
    target,
    platform_count: packages.length,
    accepted_count: packages.filter((item) => item.accepted === true).length,
    export_ready_count: packages.filter((item) => item.export_ready === true).length,
    built_in_count: packages.filter((item) => item.built_in === true).length,
    custom_authoring_count: packages.filter((item) => item.built_in !== true).length,
    browser_extension_ready_count: packages.filter((item) => item.surface_entrypoints?.browser_extension?.ready === true).length,
    provider_reconcile_count: packages.filter((item) => item.surface_entrypoints?.provider_reconcile?.ready === true).length,
    adapter_preflight_startup_ready_count: packages.filter((item) => item.adapter_preflight?.startup_ready === true).length,
    adapter_preflight_realtime_ready_count: packages.filter((item) => item.adapter_preflight?.realtime_annotation_ready === true).length,
    platforms: packages.map((item) => item.platform),
    rows: packages.map((item) => ({
      platform: item.platform,
      display_name: item.display_name,
      built_in: item.built_in,
      accepted: item.accepted,
      export_ready: item.export_ready,
      target: item.target,
      adapter_status: item.readiness.adapter_status,
      recommended_first_surface: item.recommended_first_surface,
      browser_extension_ready: item.surface_entrypoints?.browser_extension?.ready === true,
      provider_reconcile_ready: item.surface_entrypoints?.provider_reconcile?.ready === true,
      adapter_preflight_status: item.adapter_preflight?.status,
      adapter_preflight_selected_surface: item.adapter_preflight?.selected_surface,
      adapter_preflight_startup_ready: item.adapter_preflight?.startup_ready === true,
      adapter_preflight_realtime_ready: item.adapter_preflight?.realtime_annotation_ready === true,
      host_file_count: item.host_files.length,
      failed_required_ids: item.readiness.failed_required_ids,
      first_next_action: item.next_actions?.[0],
    })),
    packages,
    portfolio,
    acceptance_checklist_matrix: checklistMatrix,
    next_actions: unique(packages.flatMap((item) => item.next_actions ?? [])),
  };
}
