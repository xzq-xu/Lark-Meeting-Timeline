import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  buildMeetingAppAdapterCapabilityMatrix,
  buildMeetingAppAdapterCapabilityReport,
  buildMeetingAppAdapterExecutionPlan,
} from './meeting-app-adapter-capability.mjs';
import {
  buildMeetingAppAdapterHandoffPackage,
} from './meeting-app-adapter-handoff-package.mjs';
import {
  normalizeMeetingPlatform,
} from './platform-setup.mjs';
import {
  buildMeetingPlatformRuntimeBundle,
} from './platform-runtime-bundle.mjs';

export const MEETING_APP_ADAPTER_INTEGRATION_PACKAGE_SCHEMA = 'meeting_app_adapter_integration_package';
export const MEETING_APP_ADAPTER_INTEGRATION_PACKAGE_MATRIX_SCHEMA = 'meeting_app_adapter_integration_package_matrix';
export const MEETING_APP_ADAPTER_INTEGRATION_PACKAGE_SCHEMA_VERSION = 1;

const DEFAULT_MEETING_APP_ADAPTER_PLATFORMS = Object.freeze([
  'lark',
  'google_meet',
  'microsoft_teams',
  'zoom',
  'webex',
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

function keyAliases(platform) {
  const dashed = platform.replaceAll('_', '-');
  const aliases = platform === 'microsoft_teams'
    ? ['teams', 'microsoft-teams']
    : platform === 'google_meet'
      ? ['meet', 'google-meet']
      : platform === 'lark'
        ? ['feishu', 'larksuite']
        : [];
  return unique([platform, dashed, ...aliases]);
}

function selectedPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(
    options.platforms,
    options.platform_keys,
    options.platformKeys,
    DEFAULT_MEETING_APP_ADAPTER_PLATFORMS,
  )).map((platform) => normalizeMeetingPlatform(platform)));
}

function inputForPlatform(platform, options = {}) {
  const source = firstNonEmpty(
    options.inputs,
    options.inputByPlatform,
    options.input_by_platform,
    options.snapshots,
    options.snapshotByPlatform,
    options.snapshot_by_platform,
  );
  if (source && typeof source === 'object' && !Array.isArray(source)) {
    for (const key of keyAliases(platform)) {
      if (source[key] != null) return source[key];
    }
  }
  return firstNonEmpty(options.input, options.snapshot, options.sample);
}

function evidenceForPlatform(platform, options = {}) {
  const source = firstNonEmpty(
    options.evidenceByAdapter,
    options.evidence_by_adapter,
    options.evidenceByPlatform,
    options.evidence_by_platform,
  );
  if (source && typeof source === 'object' && !Array.isArray(source)) {
    for (const key of keyAliases(platform)) {
      if (source[key] != null) return source[key];
    }
  }
  return firstNonEmpty(options.evidence, options.evidence_package, options.evidencePackage);
}

function capabilityForInput(platformOrCapability = {}, options = {}) {
  if (platformOrCapability?.schema === 'meeting_app_adapter_capability_report') return platformOrCapability;
  if (platformOrCapability?.schema === 'meeting_app_adapter_execution_plan') {
    return platformOrCapability.capability_report;
  }
  if (typeof platformOrCapability === 'string') {
    const platform = normalizeMeetingPlatform(platformOrCapability);
    return buildMeetingAppAdapterCapabilityReport(platform, {
      ...options,
      input: inputForPlatform(platform, options),
      evidence: evidenceForPlatform(platform, options),
    });
  }
  return buildMeetingAppAdapterCapabilityReport(platformOrCapability, options);
}

function executionPlanForCapability(capability = {}, options = {}) {
  return buildMeetingAppAdapterExecutionPlan(capability, options);
}

function buildFiles({ handoffPackage, capabilityReport, executionPlan, runtimeDelivery }) {
  return [
    ...(handoffPackage.files ?? []),
    {
      path: 'capability-report.json',
      media_type: 'application/json',
      schema: capabilityReport.schema,
      content: capabilityReport,
    },
    {
      path: 'execution-plan.json',
      media_type: 'application/json',
      schema: executionPlan.schema,
      content: executionPlan,
    },
    {
      path: 'runtime-delivery.json',
      media_type: 'application/json',
      schema: 'meeting_app_adapter_runtime_delivery',
      content: runtimeDelivery,
    },
  ];
}

function integrationEntryPoints(handoffPackage = {}) {
  const runtime = handoffPackage.consumer_entrypoints ?? {};
  return compactObject({
    sdk_core: '@ai-annotation/meeting-timeline-sdk',
    platform_kit: '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit',
    adapter_integration_package: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-integration-package',
    adapter_capability: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-capability',
    adapter_handoff_package: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-handoff-package',
    platform_runtime_bundle: '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-bundle',
    browser_runtime_factory: runtime.browser_runtime_factory,
    content_script_bridge_factory: runtime.content_script_bridge_factory,
    capture_function: runtime.capture_function,
    track_runtime_factory: runtime.track_runtime_factory,
    candidate_observer: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-apps#observeMeetingAppSample',
    runtime_event_client: '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-event#createMeetingPlatformRuntimeEventClient',
    timeline_client: '@ai-annotation/meeting-timeline-sdk#createMeetingTimelineClient',
  });
}

function commandSet(platform, options = {}) {
  const baseUrl = firstNonEmpty(options.baseUrl, options.base_url);
  const baseUrlArg = baseUrl ? ` --base-url=${baseUrl}` : '';
  return {
    generate_capability_report: `npm run meeting-app:adapter-capability -- --platforms=${platform}${baseUrlArg}`,
    export_adapter_handoff_package: `npm run meeting-app:adapter-handoff-package -- --platforms=${platform}${baseUrlArg}`,
    export_runtime_bundle: `npm run meeting-platform:runtime-bundle -- --platforms=${platform}${baseUrlArg}`,
    verify_adapter_evidence: `npm run meeting-app:adapter-verify -- --platforms=${platform}${baseUrlArg}`,
    verify_platform_handoff: `npm run meeting-platform:handoff-readiness -- --platforms=${platform}${baseUrlArg}`,
    run_runtime_host_replay: `npm run meeting-platform:runtime-host-replay -- --platforms=${platform}${baseUrlArg}`,
    package_smoke: 'npm run sdk:package-smoke',
  };
}

function evidenceContract(handoffPackage = {}, executionPlan = {}) {
  const liveEvidence = handoffPackage.validation?.required_live_evidence ?? [];
  const planEvidence = executionPlan.steps?.flatMap((step) => step.evidence ?? []) ?? [];
  return {
    timestamp_field: handoffPackage.contracts?.timestamp_field ?? 'captured_at_ms',
    required_for_realtime: unique([
      'candidate_observation',
      'annotation_insert_current_axis',
      ...executionPlan.steps
        ?.filter((step) => step.required_for_realtime)
        .flatMap((step) => step.evidence ?? []) ?? [],
    ]),
    required_for_production: unique([
      ...liveEvidence,
      'provider_meeting_started_event',
      'provider_meeting_ended_event',
    ]),
    accepted_evidence: unique([...liveEvidence, ...planEvidence]),
    provider_events_block_realtime: handoffPackage.contracts?.provider_events_block_realtime,
    transcript_blocks_realtime: handoffPackage.contracts?.transcript_blocks_realtime,
  };
}

function runtimeDelivery(platform, options = {}) {
  const bundle = buildMeetingPlatformRuntimeBundle(platform, options);
  return compactObject({
    schema: 'meeting_app_adapter_runtime_delivery',
    platform: bundle.platform,
    display_name: bundle.display_name,
    runtime_bundle_schema: bundle.schema,
    runtime_bundle_id: bundle.id,
    objective: bundle.objective,
    modules: bundle.modules,
    adapter_route: compactObject({
      schema: bundle.adapter_route?.schema,
      ready: bundle.readiness?.runtime_ready === true && (bundle.adapter_route?.route_count ?? 0) > 0,
      recommended_mode: bundle.adapter_route?.recommended_mode,
      selected_routes: bundle.adapter_route?.route_order
        ?? bundle.adapter_route?.routes?.map((route) => firstNonEmpty(route.route, route.id, route.name, route.type)),
      first_route: bundle.adapter_route?.route_order?.[0]
        ?? firstNonEmpty(
          bundle.adapter_route?.routes?.[0]?.route,
          bundle.adapter_route?.routes?.[0]?.id,
          bundle.adapter_route?.routes?.[0]?.name,
          bundle.adapter_route?.routes?.[0]?.type,
        ),
      local_observer_first: String(bundle.adapter_route?.recommended_mode ?? '').includes('local_observer_first'),
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
    }),
    browser: {
      matches: bundle.browser?.matches ?? [],
      host_permissions: bundle.browser?.host_permissions ?? [],
      content_script_count: bundle.browser?.content_scripts?.length ?? 0,
      candidate_observation: bundle.browser?.candidate_observation,
    },
    runtime: compactObject({
      preset: bundle.runtime?.preset,
      capture_profile: bundle.runtime?.capture_profile,
      mutation_observer: bundle.runtime?.mutation_observer,
      content_script_bridge: bundle.runtime?.content_script_bridge,
      observation_loop: bundle.runtime?.observation_loop,
      observer_scheduler: bundle.runtime?.observer_scheduler,
      runtime_host: bundle.runtime?.runtime_host,
    }),
    messaging: compactObject({
      runtime_event_endpoint: bundle.messaging?.runtime_event?.endpoint,
      runtime_event_schema: bundle.messaging?.runtime_event?.schema,
      accepted_methods: bundle.messaging?.accepted_methods,
      bridge_message_types: bundle.messaging?.bridge_message_types,
      background_message_types: bundle.messaging?.background_message_types,
    }),
    host: bundle.host,
    readiness: bundle.readiness,
    next_actions: bundle.next_actions,
  });
}

function integrationSteps(executionPlan = {}) {
  return (executionPlan.steps ?? []).map((step) => compactObject({
    id: step.id,
    source: step.source,
    status: step.status,
    role: step.role,
    required_for_realtime: step.required_for_realtime,
    required_for_pilot: step.required_for_pilot,
    required_for_production: step.required_for_production,
    sdk_modules: step.sdk_modules,
    evidence: step.evidence,
  }));
}

function readiness({ capabilityReport, handoffPackage, executionPlan }) {
  const staticReady = capabilityReport.static_ready === true && handoffPackage.accepted === true;
  const pilotReady = capabilityReport.pilot_ready === true && executionPlan.accepted === true;
  const productionReady = capabilityReport.production_ready === true;
  const realtimeReady = executionPlan.realtime_ready === true;
  return {
    accepted: staticReady && pilotReady,
    static_ready: staticReady,
    pilot_ready: pilotReady,
    production_ready: productionReady,
    realtime_ready: realtimeReady,
    first_blocked_step: executionPlan.first_blocked_step,
    risk_level: capabilityReport.risk_level,
    recommended_mode: capabilityReport.recommended_mode,
    missing_evidence_count: capabilityReport.evidence_state?.missing_evidence_count,
  };
}

function acceptanceTarget(options = {}) {
  return String(firstNonEmpty(
    options.target,
    options.acceptance_target,
    options.acceptanceTarget,
    options.requireProductionReady === true || options.require_production_ready === true ? 'production' : undefined,
    options.requireRealtimeReady === true || options.require_realtime_ready === true ? 'realtime' : undefined,
    'pilot',
  ));
}

function targetAccepted(pkg = {}, target = 'pilot') {
  switch (target) {
    case 'static':
      return pkg.static_ready === true;
    case 'realtime':
      return pkg.accepted === true && pkg.realtime_ready === true;
    case 'production':
      return pkg.production_ready === true;
    case 'pilot':
    default:
      return pkg.accepted === true && pkg.pilot_ready === true;
  }
}

function targetMissingGates(pkg = {}, target = 'pilot') {
  const missing = [];
  if (!pkg.static_ready) missing.push('static_ready');
  if ((target === 'pilot' || target === 'realtime' || target === 'production') && !pkg.pilot_ready) missing.push('pilot_ready');
  if ((target === 'realtime' || target === 'production') && !pkg.realtime_ready) missing.push('realtime_ready');
  if (target === 'production' && !pkg.production_ready) missing.push('production_ready');
  if (pkg.first_blocked_step) missing.push(pkg.first_blocked_step);
  return unique(missing);
}

export function buildMeetingAppAdapterIntegrationPackage(platformOrCapability = {}, options = {}) {
  const capabilityReport = capabilityForInput(platformOrCapability, options);
  const platform = capabilityReport.platform;
  const executionPlan = executionPlanForCapability(capabilityReport, options);
  const handoffPackage = buildMeetingAppAdapterHandoffPackage(platform, options);
  const runtime = runtimeDelivery(platform, options);
  const packageReadiness = readiness({ capabilityReport, handoffPackage, executionPlan });
  const files = buildFiles({ handoffPackage, capabilityReport, executionPlan, runtimeDelivery: runtime });
  return {
    type: 'meeting_app_adapter_integration_package',
    schema: MEETING_APP_ADAPTER_INTEGRATION_PACKAGE_SCHEMA,
    schema_version: MEETING_APP_ADAPTER_INTEGRATION_PACKAGE_SCHEMA_VERSION,
    platform,
    display_name: capabilityReport.display_name,
    accepted: packageReadiness.accepted,
    static_ready: packageReadiness.static_ready,
    pilot_ready: packageReadiness.pilot_ready,
    production_ready: packageReadiness.production_ready,
    realtime_ready: packageReadiness.realtime_ready,
    recommended_mode: packageReadiness.recommended_mode,
    risk_level: packageReadiness.risk_level,
    first_blocked_step: packageReadiness.first_blocked_step,
    artifact_count: files.length,
    file_count: files.length,
    files,
    file_paths: files.map((file) => file.path),
    entrypoints: integrationEntryPoints(handoffPackage),
    commands: commandSet(platform, options),
    runtime_delivery: runtime,
    evidence_contract: evidenceContract(handoffPackage, executionPlan),
    integration_steps: integrationSteps(executionPlan),
    gates: executionPlan.gates,
    readiness: packageReadiness,
    acceptance: {
      default_target: 'pilot',
      static_ready: targetAccepted(packageReadiness, 'static'),
      pilot_ready: targetAccepted(packageReadiness, 'pilot'),
      realtime_ready: targetAccepted(packageReadiness, 'realtime'),
      production_ready: targetAccepted(packageReadiness, 'production'),
      missing_for_production: targetMissingGates(packageReadiness, 'production'),
    },
    handoff_package: handoffPackage,
    capability_report: capabilityReport,
    execution_plan: executionPlan,
    next_actions: unique([
      ...capabilityReport.next_actions,
      ...executionPlan.next_actions,
      ...(handoffPackage.next_actions ?? []),
    ]),
  };
}

export function buildMeetingAppAdapterIntegrationPackageMatrix(options = {}) {
  const capabilityMatrix = options.capabilityMatrix?.schema === 'meeting_app_adapter_capability_matrix'
    ? options.capabilityMatrix
    : buildMeetingAppAdapterCapabilityMatrix(options);
  const platforms = selectedPlatforms({
    ...options,
    platforms: capabilityMatrix.platforms,
  });
  const reportByPlatform = Object.fromEntries((capabilityMatrix.reports ?? []).map((report) => [report.platform, report]));
  const packages = platforms.map((platform) => buildMeetingAppAdapterIntegrationPackage(
    reportByPlatform[platform] ?? platform,
    {
      ...options,
      input: inputForPlatform(platform, options),
      evidence: evidenceForPlatform(platform, options),
    },
  ));
  return {
    type: 'meeting_app_adapter_integration_package_matrix',
    schema: MEETING_APP_ADAPTER_INTEGRATION_PACKAGE_MATRIX_SCHEMA,
    schema_version: MEETING_APP_ADAPTER_INTEGRATION_PACKAGE_SCHEMA_VERSION,
    accepted: packages.every((pkg) => pkg.accepted === true),
    platform_count: packages.length,
    accepted_count: packages.filter((pkg) => pkg.accepted === true).length,
    static_ready_count: packages.filter((pkg) => pkg.static_ready === true).length,
    pilot_ready_count: packages.filter((pkg) => pkg.pilot_ready === true).length,
    production_ready_count: packages.filter((pkg) => pkg.production_ready === true).length,
    realtime_ready_count: packages.filter((pkg) => pkg.realtime_ready === true).length,
    platforms: packages.map((pkg) => pkg.platform),
    rows: packages.map((pkg) => ({
      platform: pkg.platform,
      display_name: pkg.display_name,
      accepted: pkg.accepted,
      static_ready: pkg.static_ready,
      pilot_ready: pkg.pilot_ready,
      production_ready: pkg.production_ready,
      realtime_ready: pkg.realtime_ready,
      recommended_mode: pkg.recommended_mode,
      risk_level: pkg.risk_level,
      first_blocked_step: pkg.first_blocked_step,
      first_next_action: pkg.next_actions[0],
    })),
    capability_matrix: capabilityMatrix,
    packages,
    next_actions: unique(packages.flatMap((pkg) => pkg.next_actions ?? [])),
  };
}

export function assertMeetingAppAdapterIntegrationPackage(packageOrPlatform = {}, options = {}) {
  const pkg = packageOrPlatform?.schema === MEETING_APP_ADAPTER_INTEGRATION_PACKAGE_SCHEMA
    ? packageOrPlatform
    : buildMeetingAppAdapterIntegrationPackage(packageOrPlatform, options);
  const target = acceptanceTarget(options);
  if (!targetAccepted(pkg, target)) {
    throw new MeetingTimelineSdkError('Meeting app adapter integration package acceptance failed', {
      code: 'meeting_app_adapter_integration_package_rejected',
      platform: pkg.platform,
      target,
      accepted: pkg.accepted,
      static_ready: pkg.static_ready,
      pilot_ready: pkg.pilot_ready,
      realtime_ready: pkg.realtime_ready,
      production_ready: pkg.production_ready,
      first_blocked_step: pkg.first_blocked_step,
      missing_gates: targetMissingGates(pkg, target),
      next_actions: pkg.next_actions,
    });
  }
  return pkg;
}

export function assertMeetingAppAdapterIntegrationPackageMatrix(matrixOrOptions = {}, options = {}) {
  const matrix = matrixOrOptions?.schema === MEETING_APP_ADAPTER_INTEGRATION_PACKAGE_MATRIX_SCHEMA
    ? matrixOrOptions
    : buildMeetingAppAdapterIntegrationPackageMatrix({ ...matrixOrOptions, ...options });
  const target = acceptanceTarget(options);
  const rejected = matrix.packages.filter((pkg) => !targetAccepted(pkg, target));
  if (rejected.length > 0) {
    throw new MeetingTimelineSdkError('Meeting app adapter integration package matrix acceptance failed', {
      code: 'meeting_app_adapter_integration_package_matrix_rejected',
      target,
      rejected_platforms: rejected.map((pkg) => pkg.platform),
      rows: matrix.rows,
      next_actions: matrix.next_actions,
    });
  }
  return matrix;
}
