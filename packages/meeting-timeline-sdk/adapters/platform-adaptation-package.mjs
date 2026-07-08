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
  buildMeetingPlatformRuntimeEventPlan,
} from './platform-runtime-event.mjs';
import {
  buildMeetingPlatformFieldCollectorConfig,
} from './platform-field-capture.mjs';
import {
  buildMeetingAppRuntimeAdapterConfig,
} from './meeting-app-profile.mjs';
import {
  MEETING_APP_EXTENSION_MESSAGE_TYPES,
  buildMeetingAppExtensionInstallPlan,
  buildMeetingAppExtensionMatchPatterns,
} from './meeting-app-extension.mjs';
import {
  buildMeetingPlatformLiveAdapterReadiness,
} from './platform-live-adapter.mjs';
import {
  buildMeetingPlatformHandoffReadiness,
} from './platform-handoff-readiness.mjs';
import {
  buildMeetingPlatformAdaptationStrategy,
} from './platform-strategy.mjs';
import {
  buildMeetingPlatformAdapterSelection,
} from './platform-adapter-selection.mjs';

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
    print_adapter_selection: `npm run meeting-platform:adapter-selection -- --platforms=${platform}${baseUrlArg}`,
    print_runtime_event_plan: `npm run meeting-platform:runtime-event-plan -- --platforms=${platform}${baseUrlArg}`,
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

function localObserverSummary(collector = {}, runtimeAdapter = {}, runtimeEventPlan = {}, extensionPlan = {}) {
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
    candidate_observation: candidateObservationSummary(runtimeAdapter, runtimeEventPlan, extensionPlan),
    required_snapshots: snapshots?.required_snapshots ?? [],
    minimum_record_count: snapshots?.minimum_record_count ?? 0,
  });
}

function candidateObservationSummary(runtimeAdapter = {}, runtimeEventPlan = {}, extensionPlan = {}) {
  const observeAction = asArray(runtimeEventPlan.actions).find((row) => row?.action === 'observe_platform_candidates');
  return compactObject({
    message_type: runtimeAdapter.extension?.message_types?.observe_candidates
      ?? MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
    runtime_event_action: 'observe_platform_candidates',
    runtime_event_client_method: observeAction?.client_method ?? 'observePlatformCandidates',
    runtime_event_endpoint: runtimeEventPlan.endpoint,
    required_permission: 'tabs',
    permissions: runtimeAdapter.extension?.permissions ?? extensionPlan.manifest?.permissions ?? [],
    producer: 'browser_extension_background_or_native_host',
    example: runtimeEventPlan.examples?.observe_platform_candidates,
  });
}

function annotationSummary(contract = {}, collector = {}, runtimeEventPlan = {}) {
  return compactObject({
    insert_endpoint: contract.annotations?.endpoints?.insertMark ?? collector.timeline_ingest?.endpoints?.insertMark,
    runtime_event_endpoint: contract.annotations?.endpoints?.runtimeEvents,
    runtime_event: contract.annotations?.runtime_event,
    runtime_event_plan: runtimeEventPlan,
    runtime_event_actions: runtimeEventPlan.actions?.map((row) => row.action) ?? [],
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
      runtime_event: '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-event',
      adapter_selection: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-selection',
      platform_events: capabilities.sdk_modules?.events,
      runtime_adapter: runtimeAdapter.sdk_module,
      extension: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-extension',
    }),
    kit_methods: unique([
      'platformAdaptationPackage',
      'platformAdapterSelection',
      'platformAdapterSelectionMatrix',
      'platformLiveAdapter',
      'platformRuntimeProfile',
      'platformRuntimeEventPlan',
      'platformAdapterContract',
      'meetingAppRuntimeAdapterConfig',
      'platformFieldCollectorConfig',
      'insertAnnotation',
      ...(contract.implementation?.kit_methods ?? []),
    ]),
    extension: {
      install_plan_schema: extensionPlan.schema,
      match_count: extensionPlan.matches?.length ?? extensionPlan.content_scripts?.[0]?.matches?.length ?? 0,
      candidate_observer_message_type: MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
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
  const runtimeEventPlan = buildMeetingPlatformRuntimeEventPlan(key, options);
  const adapterSelection = buildMeetingPlatformAdapterSelection(key, {}, options);
  const runtimeAdapter = safeRuntimeAdapterConfig(key, options);
  const extensionMatches = safeExtensionMatches(key, options);
  const extensionPlan = safeExtensionInstallPlan(key, options);
  const liveReadiness = safeLiveReadiness(key, options);
  const handoffReadiness = safeHandoffReadiness(key, options);
  const strategy = buildMeetingPlatformAdaptationStrategy(key, options);
  const candidateObservation = candidateObservationSummary(runtimeAdapter, runtimeEventPlan, extensionPlan);

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
    adapter_surfaces: runtime.adapter_surfaces,
    launch_requirements: runtime.launch_requirements,
    evidence_thresholds: runtime.evidence_thresholds,
    fallback_policy: runtime.fallback_policy,
    local_observer: localObserverSummary(collector, runtimeAdapter, runtimeEventPlan, extensionPlan),
    provider_observer: providerEventSummary(provider, runtime),
    adaptation_strategy: strategy,
    adaptation_playbook: strategy.adaptation_playbook,
    annotation_pipeline: annotationSummary(contract, collector, runtimeEventPlan),
    adapter_selection: {
      schema: adapterSelection.schema,
      ready: adapterSelection.readiness?.selection_ready === true,
      axis_source: adapterSelection.selection?.axis_source,
      axis_surface: adapterSelection.selection?.axis_surface,
      annotation_source: adapterSelection.selection?.annotation_source,
      timestamp_field: adapterSelection.selection?.timestamp_field,
      provider_reconcile_source: adapterSelection.selection?.provider_reconcile_source,
      provider_reconcile_required_for_production: adapterSelection.selection?.provider_reconcile_required_for_production,
      provider_events_block_realtime: adapterSelection.runtime_policy?.provider_events_block_realtime,
      transcript_blocks_realtime: adapterSelection.runtime_policy?.transcript_blocks_realtime,
      startup_order: adapterSelection.runtime_policy?.startup_order,
      current_evidence: adapterSelection.current_evidence,
      next_actions: adapterSelection.next_actions,
    },
    candidate_observation: candidateObservation,
    runtime_event_plan: runtimeEventPlan,
    speaker_markers: runtime.speaker_markers,
    transcript: transcriptSummary(capabilities, runtime),
    extension: {
      matches: extensionMatches.matches ?? [],
      host_permissions: extensionMatches.host_permissions ?? [],
      permissions: runtimeAdapter.extension?.permissions ?? extensionPlan.manifest?.permissions ?? [],
      content_scripts: extensionMatches.content_scripts ?? [],
      candidate_observation: candidateObservation,
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
    candidate_observer_count: packages.filter((item) => item.candidate_observation?.runtime_event_action === 'observe_platform_candidates').length,
    adapter_selection_ready_count: packages.filter((item) => item.adapter_selection?.ready).length,
    provider_observer_count: packages.filter((item) => item.provider_observer?.transport).length,
    production_ready_count: packages.filter((item) => item.readiness.production_ready).length,
    realtime_ready_count: packages.filter((item) => item.readiness.ready_for_realtime_annotations).length,
    platforms: packages.map((item) => item.platform),
    rows: packages.map((item) => ({
      platform: item.platform,
      display_name: item.display_name,
      mode: item.mode,
      recommended_mode: item.recommended_mode,
      primary_surface: item.adapter_surfaces?.primary,
      surface_order: item.adapter_surfaces?.recommended_order,
      provider_reconcile_surface: item.adapter_surfaces?.provider_reconcile_surface,
      sdk_wiring_ready: item.readiness.sdk_wiring_ready,
      production_ready: item.readiness.production_ready,
      ready_for_realtime_annotations: item.readiness.ready_for_realtime_annotations,
      next_phase: item.adaptation_playbook?.next_phase,
      next_phase_priority: item.adaptation_playbook?.next_phase_priority,
      provider_path: item.adaptation_playbook?.integration_path?.path,
      provider_permission_risk: item.adaptation_playbook?.risk_profile?.permission_risk,
      speaker_realtime_gap: item.adaptation_playbook?.risk_profile?.speaker_realtime_gap,
      post_meeting_backfill_supported: item.adaptation_playbook?.risk_profile?.transcript_availability_risk === 'post_meeting_only',
      browser_match_count: item.extension.matches.length,
      candidate_observation_ready: item.candidate_observation?.runtime_event_action === 'observe_platform_candidates',
      candidate_observer_message_type: item.candidate_observation?.message_type,
      candidate_observer_permission: item.candidate_observation?.required_permission,
      candidate_observer_client_method: item.candidate_observation?.runtime_event_client_method,
      pilot_provider_records_required: item.evidence_thresholds?.pilot?.provider_records_required,
      production_provider_records_required: item.evidence_thresholds?.production?.provider_records_required,
      provider_transport: item.provider_observer?.transport,
      provider_start_event_count: item.provider_observer?.start_events?.length ?? 0,
      provider_end_event_count: item.provider_observer?.end_events?.length ?? 0,
      runtime_event_action_count: item.annotation_pipeline?.runtime_event_actions?.length ?? 0,
      adapter_selection_ready: item.adapter_selection?.ready === true,
      adapter_selection_axis_source: item.adapter_selection?.axis_source,
      adapter_selection_axis_surface: item.adapter_selection?.axis_surface,
      adapter_selection_timestamp_field: item.adapter_selection?.timestamp_field,
      adapter_selection_provider_blocks_realtime: item.adapter_selection?.provider_events_block_realtime === true,
      adapter_selection_transcript_blocks_realtime: item.adapter_selection?.transcript_blocks_realtime === true,
      speaker_min_stable_ms: item.speaker_markers?.filter?.min_stable_ms,
      transcript_blocks_realtime: item.transcript?.blocks_realtime_annotation === true,
      missing_item_count: item.readiness.missing_items?.length ?? 0,
      first_next_action: item.next_actions?.[0],
    })),
    packages,
    next_actions: unique(packages.flatMap((item) => item.next_actions ?? [])),
  };
}
