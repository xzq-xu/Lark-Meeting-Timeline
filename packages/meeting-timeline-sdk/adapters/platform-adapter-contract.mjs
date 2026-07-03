import { compactObject } from '../index.mjs';
import {
  MEETING_PLATFORM_KEYS,
  normalizeMeetingPlatform,
  platformCapabilityContract,
} from './platform-setup.mjs';
import { buildMeetingPlatformProviderConnectionPack } from './platform-provider-connection.mjs';
import { buildMeetingPlatformRuntimeProfile } from './platform-runtime-profile.mjs';
import { buildMeetingPlatformFieldCollectorConfig } from './platform-field-capture.mjs';

export const MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA = 'meeting_platform_adapter_contract';
export const MEETING_PLATFORM_ADAPTER_CONTRACT_MATRIX_SCHEMA = 'meeting_platform_adapter_contract_matrix';
export const MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA_VERSION = 1;

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

function providerEvents(runtime = {}) {
  return {
    start_events: runtime.provider_events?.start_events ?? [],
    end_events: runtime.provider_events?.end_events ?? [],
    participant_events: runtime.provider_events?.participant_events ?? [],
    artifact_events: runtime.provider_events?.artifact_events ?? [],
    lifecycle_events: runtime.provider_events?.lifecycle_events ?? [],
  };
}

function localObserverContract(collector = {}, runtime = {}) {
  const observer = collector.browser_observer;
  const snapshots = collector.local_snapshot_collector;
  if (!observer && !snapshots && !runtime.speaker_markers) return undefined;
  return compactObject({
    required_for_realtime_axis: true,
    surface: observer?.type ?? 'trusted_host_or_native_observer',
    matches: observer?.matches ?? [],
    host_permissions: observer?.host_permissions ?? [],
    content_scripts: observer?.content_scripts ?? [],
    snapshot_collector: snapshots,
    speaker_markers: runtime.speaker_markers,
  });
}

function providerObserverContract(provider = {}, runtime = {}, collector = {}) {
  const events = providerEvents(runtime);
  return compactObject({
    role: provider.provider_role,
    transport: provider.transport,
    endpoint: provider.endpoint,
    status_endpoint: provider.status_endpoint,
    required_for_realtime: false,
    required_for_production_evidence: runtime.axis?.provider_reconcile?.required_for_production_evidence,
    event_mapping: provider.event_mapping ?? [],
    events,
    security: provider.security,
    readiness: provider.readiness,
    collector_requirements: collector.provider_observer,
  });
}

function evidenceContract(collector = {}) {
  const manifest = collector.manifest ?? {};
  const acceptance = collector.acceptance ?? {};
  return compactObject({
    status: manifest.status,
    production_ready: manifest.production_ready === true,
    ready_for_realtime_annotations: manifest.ready_for_realtime_annotations === true,
    files: collector.storage?.files,
    input_contract: collector.storage?.input_contract,
    required_provider_coverage: acceptance.required_provider_coverage ?? [],
    required_local_snapshots: acceptance.required_local_snapshots ?? [],
    minimum_provider_records: acceptance.minimum_provider_records ?? 0,
    minimum_local_records: acceptance.minimum_local_records ?? 0,
    missing_items: acceptance.current_missing_items ?? [],
    production_condition: acceptance.production_condition,
    pilot_condition: acceptance.pilot_condition,
  });
}

function implementationContract(capabilities = {}, collector = {}, runtime = {}) {
  return {
    imports: compactObject({
      client: '@ai-annotation/meeting-timeline-sdk',
      kit: '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit',
      contract: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-contract',
      events: capabilities.sdk_modules?.events,
      ingest: capabilities.sdk_modules?.ingest,
      local_observer: capabilities.sdk_modules?.local_observer,
      transcript: capabilities.sdk_modules?.transcript,
      security: capabilities.sdk_modules?.security,
    }),
    kit_methods: [
      'platformAdapterContract',
      'platformRuntimeProfile',
      'platformProviderConnectionPack',
      'platformFieldCollectorConfig',
      'platformFieldEvidenceBundle',
      'insertMark',
      'importTranscript',
    ],
    commands: collector.automation?.commands ?? {},
    order: runtime.implementation_order ?? [],
  };
}

function readinessContract(provider = {}, collector = {}, runtime = {}) {
  const evidence = evidenceContract(collector);
  return {
    status: evidence.status,
    production_ready: evidence.production_ready,
    ready_for_realtime_annotations: evidence.ready_for_realtime_annotations,
    provider_ready: provider.readiness?.ready === true,
    provider_missing_env: provider.security?.missing_env ?? [],
    missing_items: evidence.missing_items ?? [],
    risks: runtime.risks ?? [],
  };
}

export function buildMeetingPlatformAdapterContract(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const capabilities = platformCapabilityContract(key, options);
  const runtime = buildMeetingPlatformRuntimeProfile(key, options);
  const provider = buildMeetingPlatformProviderConnectionPack(key, options);
  const collector = buildMeetingPlatformFieldCollectorConfig(key, options);
  const readiness = readinessContract(provider, collector, runtime);
  const nextActions = unique([
    ...(collector.next_actions ?? []),
    ...(provider.next_actions ?? []),
    ...(runtime.next_actions ?? []),
  ]);

  return compactObject({
    type: 'meeting_platform_adapter_contract',
    schema: MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA_VERSION,
    platform: key,
    display_name: capabilities.display_name ?? runtime.display_name,
    objective: 'portable_sdk_contract_for_adapting_a_meeting_app_to_the_annotation_timeline',
    mode: collector.mode,
    supported_surfaces: {
      local_observer_or_host_detector: true,
      browser_observer: Boolean(collector.browser_observer),
      provider_webhook_or_event_subscription: key !== 'local_detector',
      post_meeting_transcript_import: capabilities.post_meeting_transcript?.status ?? 'unknown',
      realtime_transcript_required: false,
    },
    timebase: runtime.runtime_contract,
    realtime_axis: {
      primary_source: runtime.axis?.primary_source,
      start: runtime.axis?.start,
      end: runtime.axis?.end,
      provider_reconcile: runtime.axis?.provider_reconcile,
      rules: [
        'create_or_bind_axis_from_local_observer_before_provider_event_arrives',
        'insert_annotation_by_captured_at_ms_on_the_active_axis',
        'use_provider_events_only_for_reconcile_and_backfill',
        'never_wait_for_post_meeting_transcript_before_inserting_realtime_marks',
      ],
    },
    annotations: {
      endpoints: collector.timeline_ingest?.endpoints,
      timestamp_field: runtime.runtime_contract?.annotation_timestamp_field,
      policy: runtime.annotations,
      realtime_policy: collector.timeline_ingest?.realtime_annotation_policy,
    },
    local_observer: localObserverContract(collector, runtime),
    provider_observer: providerObserverContract(provider, runtime, collector),
    transcript: runtime.transcript,
    evidence: evidenceContract(collector),
    implementation: implementationContract(capabilities, collector, runtime),
    readiness,
    next_actions: nextActions,
  });
}

export function buildMeetingPlatformAdapterContractMatrix(options = {}) {
  const contracts = selectedPlatforms(options).map((platform) => buildMeetingPlatformAdapterContract(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_adapter_contract_matrix',
    schema: MEETING_PLATFORM_ADAPTER_CONTRACT_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA_VERSION,
    platform_count: contracts.length,
    browser_observer_count: contracts.filter((contract) => contract.supported_surfaces.browser_observer).length,
    provider_observer_count: contracts.filter((contract) => contract.supported_surfaces.provider_webhook_or_event_subscription).length,
    production_ready_count: contracts.filter((contract) => contract.readiness.production_ready).length,
    realtime_ready_count: contracts.filter((contract) => contract.readiness.ready_for_realtime_annotations).length,
    platforms: contracts.map((contract) => contract.platform),
    rows: contracts.map((contract) => ({
      platform: contract.platform,
      display_name: contract.display_name,
      mode: contract.mode,
      browser_observer: contract.supported_surfaces.browser_observer,
      provider_transport: contract.provider_observer?.transport,
      provider_ready: contract.readiness.provider_ready,
      production_ready: contract.readiness.production_ready,
      ready_for_realtime_annotations: contract.readiness.ready_for_realtime_annotations,
      start_create_on: contract.realtime_axis?.start?.create_on,
      end_create_on: contract.realtime_axis?.end?.create_on,
      provider_start_event_count: contract.provider_observer?.events?.start_events?.length ?? 0,
      provider_end_event_count: contract.provider_observer?.events?.end_events?.length ?? 0,
      speaker_min_stable_ms: contract.local_observer?.speaker_markers?.filter?.min_stable_ms,
      transcript_import: contract.transcript?.availability,
      missing_item_count: contract.readiness.missing_items?.length ?? 0,
    })),
    contracts,
    next_actions: unique(contracts.flatMap((contract) => contract.next_actions ?? [])),
  };
}
