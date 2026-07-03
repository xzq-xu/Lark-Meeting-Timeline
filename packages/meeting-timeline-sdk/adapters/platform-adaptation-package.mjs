import { compactObject } from '../index.mjs';
import {
  MEETING_PLATFORM_KEYS,
  buildPlatformIntegrationPlan,
  normalizeMeetingPlatform,
  platformCapabilityContract,
} from './platform-setup.mjs';
import {
  buildMeetingPlatformAdapterContract,
  buildMeetingPlatformAdapterContractAcceptanceReport,
} from './platform-adapter-contract.mjs';
import {
  buildMeetingPlatformProviderConnectionPack,
} from './platform-provider-connection.mjs';
import {
  buildMeetingPlatformRuntimeProfile,
} from './platform-runtime-profile.mjs';
import {
  buildMeetingPlatformFieldCollectorConfig,
} from './platform-field-capture.mjs';
import {
  buildMeetingAppRuntimeAdapterConfig,
} from './meeting-app-profile.mjs';
import {
  buildMeetingAppExtensionInstallPlan,
  buildMeetingAppExtensionMatchPatterns,
} from './meeting-app-extension.mjs';
import {
  buildMeetingPlatformLiveAdapterReadiness,
} from './platform-live-adapter.mjs';
import {
  buildMeetingPlatformHandoffReadiness,
} from './platform-handoff-readiness.mjs';

export const MEETING_PLATFORM_ADAPTATION_PACKAGE_SCHEMA = 'meeting_platform_adaptation_package';
export const MEETING_PLATFORM_ADAPTATION_PACKAGE_MATRIX_SCHEMA = 'meeting_platform_adaptation_package_matrix';
export const MEETING_PLATFORM_ADAPTATION_PACKAGE_SCHEMA_VERSION = 1;

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

function packageId(platform, options = {}) {
  return String(firstNonEmpty(
    options.packageId,
    options.package_id,
    `${platform}-meeting-timeline-adaptation`,
  ));
}

function commandSet(platform, options = {}) {
  const baseUrl = firstNonEmpty(options.baseUrl, options.base_url);
  const baseUrlArg = baseUrl ? ` --base-url=${baseUrl}` : '';
  return {
    print_package: `npm run meeting-platform:adaptation-package -- --platforms=${platform}${baseUrlArg}`,
    verify_contract: `npm run meeting-platform:adapter-contract -- --platforms=${platform}${baseUrlArg} --fail-on-rejected=true`,
    collect_field_evidence: `npm run meeting-platform:field-evidence -- --platforms=${platform}${baseUrlArg}`,
    verify_handoff: `npm run meeting-platform:handoff-readiness -- --platforms=${platform}${baseUrlArg}`,
  };
}

function providerEventSummary(provider = {}, runtime = {}) {
  return compactObject({
    role: provider.provider_role,
    transport: provider.transport,
    endpoint: provider.endpoint,
    status_endpoint: provider.status_endpoint,
    required_for_realtime: false,
    required_for_production_evidence: runtime.axis?.provider_reconcile?.required_for_production_evidence,
    start_events: runtime.provider_events?.start_events ?? [],
    end_events: runtime.provider_events?.end_events ?? [],
    participant_events: runtime.provider_events?.participant_events ?? [],
    artifact_events: runtime.provider_events?.artifact_events ?? [],
    lifecycle_events: runtime.provider_events?.lifecycle_events ?? [],
    security: provider.security,
    readiness: provider.readiness,
  });
}

function localObserverSummary(collector = {}, runtimeAdapter = {}) {
  const observer = collector.browser_observer;
  const snapshots = collector.local_snapshot_collector;
  return compactObject({
    surface: observer?.type ?? 'trusted_host_or_native_observer',
    required_for_realtime_axis: true,
    browser_matches: observer?.matches ?? [],
    host_permissions: observer?.host_permissions ?? [],
    content_scripts: observer?.content_scripts ?? [],
    runtime_preset: runtimeAdapter.runtime_options?.runtimePreset ?? runtimeAdapter.bridge_options?.browser_runtime_preset,
    capture_profile: runtimeAdapter.capture_options?.captureProfile,
    required_snapshots: snapshots?.required_snapshots ?? [],
    minimum_record_count: snapshots?.minimum_record_count ?? 0,
  });
}

function annotationSummary(contract = {}, collector = {}) {
  return compactObject({
    insert_endpoint: contract.annotations?.endpoints?.insertMark ?? collector.timeline_ingest?.endpoints?.insertMark,
    timestamp_field: 'captured_at_ms',
    realtime_policy: collector.timeline_ingest?.realtime_annotation_policy ?? contract.annotations?.realtime_policy,
    sdk_module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-realtime-annotation',
    live_adapter_method: 'insertAnnotation(platform, { capturedAtMs, ...mark })',
    required_fields: contract.annotations?.policy?.required_fields ?? ['captured_at_ms'],
    recommended_fields: contract.annotations?.policy?.recommended_fields ?? [],
  });
}

function transcriptSummary(capabilities = {}, runtime = {}) {
  return compactObject({
    availability: capabilities.post_meeting_transcript?.availability ?? runtime.transcript?.availability,
    source: capabilities.post_meeting_transcript?.source ?? runtime.transcript?.source,
    sdk_normalizer: capabilities.post_meeting_transcript?.sdk_normalizer,
    import_endpoint: capabilities.post_meeting_transcript?.import_endpoint,
    realtime_dependency: false,
    blocks_realtime_annotation: false,
  });
}

function implementationSummary(capabilities = {}, contract = {}, runtimeAdapter = {}, extensionPlan = {}) {
  return {
    package: '@ai-annotation/meeting-timeline-sdk',
    imports: compactObject({
      kit: '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit',
      live_adapter: '@ai-annotation/meeting-timeline-sdk/adapters/platform-live-adapter',
      adaptation_package: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adaptation-package',
      platform_events: capabilities.sdk_modules?.events,
      runtime_adapter: runtimeAdapter.sdk_module,
      extension: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-extension',
    }),
    kit_methods: unique([
      'platformAdaptationPackage',
      'platformLiveAdapter',
      'platformRuntimeProfile',
      'platformAdapterContract',
      'meetingAppRuntimeAdapterConfig',
      'platformFieldCollectorConfig',
      'insertAnnotation',
      ...(contract.implementation?.kit_methods ?? []),
    ]),
    extension: {
      install_plan_schema: extensionPlan.schema,
      match_count: extensionPlan.matches?.length ?? extensionPlan.content_scripts?.[0]?.matches?.length ?? 0,
    },
  };
}

function safeRuntimeAdapterConfig(platform, options = {}) {
  try {
    return buildMeetingAppRuntimeAdapterConfig(platform, options);
  } catch {
    return {};
  }
}

function safeExtensionMatches(platform, options = {}) {
  try {
    return buildMeetingAppExtensionMatchPatterns([platform], options);
  } catch {
    return {
      matches: [],
      host_permissions: [],
      content_scripts: [],
    };
  }
}

function safeExtensionInstallPlan(platform, options = {}) {
  try {
    return buildMeetingAppExtensionInstallPlan({
      ...options,
      platforms: [platform],
    });
  } catch {
    return {
      schema: 'meeting_app_extension_install_plan',
      platforms: [],
      matches: [],
      content_scripts: [],
    };
  }
}

function safeLiveReadiness(platform, options = {}) {
  try {
    return buildMeetingPlatformLiveAdapterReadiness(platform, options);
  } catch {
    return {
      status: 'not_available_for_platform',
      accepted: false,
    };
  }
}

function safeHandoffReadiness(platform, options = {}) {
  try {
    return buildMeetingPlatformHandoffReadiness(platform, {}, options);
  } catch {
    return {
      status: 'not_available_for_platform',
      accepted: false,
      missing_items: [],
      next_actions: [],
    };
  }
}

function readinessSummary(contractAcceptance = {}, liveReadiness = {}, handoffReadiness = {}) {
  const contractReady = contractAcceptance.accepted === true;
  return {
    sdk_wiring_ready: contractReady,
    realtime_insertion_contract_ready: liveReadiness.accepted === true || contractReady,
    handoff_ready: handoffReadiness.accepted === true,
    production_ready: contractAcceptance.summary?.production_ready === true,
    ready_for_realtime_annotations: contractAcceptance.summary?.ready_for_realtime_annotations === true,
    acceptance_target: contractAcceptance.target,
    issue_count: contractAcceptance.issue_count ?? 0,
    first_issue: contractAcceptance.issues?.[0]?.code,
    live_readiness_status: liveReadiness.status,
    handoff_status: handoffReadiness.status,
    missing_items: unique([
      ...(contractAcceptance.contract?.readiness?.missing_items ?? []),
      ...(handoffReadiness.missing_items ?? []),
    ]),
  };
}

function nextActionsFor(packageObject = {}) {
  const actions = unique([
    ...(packageObject.contract?.next_actions ?? []),
    ...(packageObject.collector?.next_actions ?? []),
    ...(packageObject.handoff_readiness?.next_actions ?? []),
    'wire_host_to_platform_live_adapter',
    'capture_real_meeting_app_snapshots_before_production',
    'keep_provider_events_as_reconcile_backfill_not_realtime_gate',
  ]);
  if (packageObject.readiness?.sdk_wiring_ready === true) {
    return actions.filter((action) => action !== 'fix_adapter_contract_acceptance_errors');
  }
  return unique(['fix_adapter_contract_acceptance_errors', ...actions]);
}

export function buildMeetingPlatformAdaptationPackage(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const capabilities = platformCapabilityContract(key, options);
  const integration = buildPlatformIntegrationPlan(key, options);
  const runtime = buildMeetingPlatformRuntimeProfile(key, options);
  const provider = buildMeetingPlatformProviderConnectionPack(key, options);
  const collector = buildMeetingPlatformFieldCollectorConfig(key, options);
  const contract = buildMeetingPlatformAdapterContract(key, options);
  const contractAcceptance = buildMeetingPlatformAdapterContractAcceptanceReport(contract, options);
  const runtimeAdapter = safeRuntimeAdapterConfig(key, options);
  const extensionMatches = safeExtensionMatches(key, options);
  const extensionPlan = safeExtensionInstallPlan(key, options);
  const liveReadiness = safeLiveReadiness(key, options);
  const handoffReadiness = safeHandoffReadiness(key, options);

  const base = compactObject({
    type: 'meeting_platform_adaptation_package',
    schema: MEETING_PLATFORM_ADAPTATION_PACKAGE_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTATION_PACKAGE_SCHEMA_VERSION,
    id: packageId(key, options),
    platform: key,
    display_name: capabilities.display_name ?? runtime.display_name,
    objective: 'portable_package_for_adapting_a_meeting_app_to_realtime_annotation_timeline',
    mode: collector.mode,
    recommended_mode: integration.recommended_mode,
    runtime_contract: runtime.runtime_contract,
    local_observer: localObserverSummary(collector, runtimeAdapter),
    provider_observer: providerEventSummary(provider, runtime),
    annotation_pipeline: annotationSummary(contract, collector),
    speaker_markers: runtime.speaker_markers,
    transcript: transcriptSummary(capabilities, runtime),
    extension: {
      matches: extensionMatches.matches ?? [],
      host_permissions: extensionMatches.host_permissions ?? [],
      content_scripts: extensionMatches.content_scripts ?? [],
      install_plan: extensionPlan,
    },
    evidence: {
      files: collector.storage?.files,
      accepted_inputs: collector.storage?.input_contract?.accepted_inputs,
      missing_items: collector.acceptance?.current_missing_items ?? [],
      production_condition: collector.acceptance?.production_condition,
      pilot_condition: collector.acceptance?.pilot_condition,
    },
    implementation: implementationSummary(capabilities, contract, runtimeAdapter, extensionPlan),
    commands: commandSet(key, options),
    readiness: readinessSummary(contractAcceptance, liveReadiness, handoffReadiness),
    contract_acceptance: contractAcceptance,
    live_adapter_readiness: liveReadiness,
    handoff_readiness: handoffReadiness,
    contract,
    runtime_profile: runtime,
    collector,
  });
  return {
    ...base,
    next_actions: nextActionsFor(base),
  };
}

export function buildMeetingPlatformAdaptationPackageMatrix(options = {}) {
  const packages = selectedPlatforms(options).map((platform) => buildMeetingPlatformAdaptationPackage(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_adaptation_package_matrix',
    schema: MEETING_PLATFORM_ADAPTATION_PACKAGE_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTATION_PACKAGE_SCHEMA_VERSION,
    platform_count: packages.length,
    sdk_wiring_ready_count: packages.filter((item) => item.readiness.sdk_wiring_ready).length,
    browser_observer_count: packages.filter((item) => (item.local_observer?.browser_matches ?? []).length > 0).length,
    provider_observer_count: packages.filter((item) => item.provider_observer?.transport).length,
    production_ready_count: packages.filter((item) => item.readiness.production_ready).length,
    realtime_ready_count: packages.filter((item) => item.readiness.ready_for_realtime_annotations).length,
    platforms: packages.map((item) => item.platform),
    rows: packages.map((item) => ({
      platform: item.platform,
      display_name: item.display_name,
      mode: item.mode,
      recommended_mode: item.recommended_mode,
      sdk_wiring_ready: item.readiness.sdk_wiring_ready,
      production_ready: item.readiness.production_ready,
      ready_for_realtime_annotations: item.readiness.ready_for_realtime_annotations,
      browser_match_count: item.extension.matches.length,
      provider_transport: item.provider_observer?.transport,
      provider_start_event_count: item.provider_observer?.start_events?.length ?? 0,
      provider_end_event_count: item.provider_observer?.end_events?.length ?? 0,
      speaker_min_stable_ms: item.speaker_markers?.filter?.min_stable_ms,
      transcript_blocks_realtime: item.transcript?.blocks_realtime_annotation === true,
      missing_item_count: item.readiness.missing_items?.length ?? 0,
      first_next_action: item.next_actions?.[0],
    })),
    packages,
    next_actions: unique(packages.flatMap((item) => item.next_actions ?? [])),
  };
}
