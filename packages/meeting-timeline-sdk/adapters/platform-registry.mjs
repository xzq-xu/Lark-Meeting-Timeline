import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import { normalizeGoogleMeetEvent } from './google-meet.mjs';
import { normalizeLarkEvent } from './lark.mjs';
import { normalizeLocalDetectorEvent } from './local-detector.mjs';
import { normalizeMicrosoftTeamsEvent } from './microsoft-teams.mjs';
import {
  buildMeetingPlatformAdapterContract,
  buildMeetingPlatformAdapterContractAcceptanceReport,
} from './platform-adapter-contract.mjs';
import {
  buildMeetingPlatformProviderConnectionPack,
} from './platform-provider-connection.mjs';
import {
  buildMeetingPlatformRuntimeBundle,
} from './platform-runtime-bundle.mjs';
import {
  buildMeetingPlatformRuntimeEventPlan,
} from './platform-runtime-event.mjs';
import {
  MEETING_PLATFORM_ALIASES,
  MEETING_PLATFORM_KEYS,
  normalizeMeetingPlatform,
  platformCapabilityContract,
} from './platform-setup.mjs';
import { normalizeWebexEvent } from './webex.mjs';
import { normalizeZoomEvent } from './zoom.mjs';

export const MEETING_PLATFORM_REGISTRY_ENTRY_SCHEMA = 'meeting_platform_registry_entry';
export const MEETING_PLATFORM_REGISTRY_MANIFEST_SCHEMA = 'meeting_platform_registry_manifest';
export const MEETING_PLATFORM_REGISTRY_ACCEPTANCE_SCHEMA = 'meeting_platform_registry_acceptance';
export const MEETING_PLATFORM_REGISTRY_SCHEMA_VERSION = 1;

const normalizers = Object.freeze({
  local_detector: normalizeLocalDetectorEvent,
  lark: normalizeLarkEvent,
  google_meet: normalizeGoogleMeetEvent,
  microsoft_teams: normalizeMicrosoftTeamsEvent,
  zoom: normalizeZoomEvent,
  webex: normalizeWebexEvent,
});

const sourceByPlatform = Object.freeze({
  local_detector: 'local_detector',
});

function aliasesFor(platform) {
  return Object.entries(MEETING_PLATFORM_ALIASES)
    .filter(([, key]) => key === platform)
    .map(([alias]) => alias);
}

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

function issue(severity, code, message, details = {}) {
  return compactObject({ severity, code, message, ...details });
}

function selectedPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_PLATFORM_KEYS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

export const MEETING_PLATFORM_EVENT_ADAPTERS = Object.freeze(
  MEETING_PLATFORM_KEYS.map((platform) => Object.freeze({
    key: platform,
    aliases: Object.freeze(aliasesFor(platform)),
    source: sourceByPlatform[platform] ?? `${platform}_webhook`,
    normalize: normalizers[platform],
  })),
);

const adapterByAlias = new Map(
  MEETING_PLATFORM_EVENT_ADAPTERS.flatMap((adapter) => (
    adapter.aliases.map((alias) => [alias, adapter])
  )),
);

export function meetingPlatformEventAdapterFor(platform) {
  let normalized;
  try {
    normalized = normalizeMeetingPlatform(platform);
  } catch {
    return null;
  }
  return adapterByAlias.get(normalized) ?? MEETING_PLATFORM_EVENT_ADAPTERS.find((adapter) => adapter.key === normalized) ?? null;
}

export function buildMeetingPlatformRegistryEntry(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const adapter = meetingPlatformEventAdapterFor(key);
  const capabilities = platformCapabilityContract(key, options);
  const provider = buildMeetingPlatformProviderConnectionPack(key, options);
  const contract = buildMeetingPlatformAdapterContract(key, options);
  const contractAcceptance = buildMeetingPlatformAdapterContractAcceptanceReport(contract, options);
  const runtimeBundle = buildMeetingPlatformRuntimeBundle(key, options);
  const runtimeEventPlan = buildMeetingPlatformRuntimeEventPlan(key, options);
  const candidateObservation = runtimeBundle.messaging?.candidate_observation ?? runtimeBundle.browser?.candidate_observation;
  const candidateObservationReady = candidateObservation?.runtime_event_action === 'observe_platform_candidates'
    && runtimeEventPlan.supported_actions?.includes?.('observe_platform_candidates');
  const hostBasePath = firstNonEmpty(options.basePath, options.base_path, '/api/platform-events');

  return compactObject({
    type: 'meeting_platform_registry_entry',
    schema: MEETING_PLATFORM_REGISTRY_ENTRY_SCHEMA,
    schema_version: MEETING_PLATFORM_REGISTRY_SCHEMA_VERSION,
    platform: key,
    display_name: capabilities.display_name ?? contract.display_name ?? runtimeBundle.display_name,
    aliases: adapter?.aliases ?? aliasesFor(key),
    event_adapter: {
      source: adapter?.source,
      normalize_available: typeof adapter?.normalize === 'function',
      module: capabilities.sdk_modules?.events,
    },
    supported_surfaces: contract.supported_surfaces,
    runtime: {
      bundle_schema: runtimeBundle.schema,
      runtime_ready: runtimeBundle.readiness?.runtime_ready === true,
      sdk_wiring_ready: runtimeBundle.readiness?.sdk_wiring_ready === true,
      browser_matches: runtimeBundle.browser?.matches ?? [],
      host_permissions: runtimeBundle.browser?.host_permissions ?? [],
      content_script_count: runtimeBundle.browser?.content_scripts?.length ?? 0,
      preset: runtimeBundle.runtime?.preset,
      sample_interval_ms: runtimeBundle.runtime?.start_options?.sampleIntervalMs,
      mutation_debounce_ms: runtimeBundle.runtime?.mutation_observer?.debounce_ms,
      speaker_min_stable_ms: runtimeBundle.runtime?.speaker_filter?.min_stable_ms,
      candidate_observation: compactObject({
        ready: candidateObservationReady,
        message_type: candidateObservation?.message_type,
        required_permission: candidateObservation?.required_permission,
        runtime_event_action: candidateObservation?.runtime_event_action,
        runtime_event_client_method: candidateObservation?.runtime_event_client_method,
        endpoint: '/api/meeting-platform/observe-candidates',
        producer: candidateObservation?.producer,
      }),
      runtime_event_plan_schema: runtimeEventPlan.schema,
      runtime_event_action_count: runtimeEventPlan.actions?.length ?? 0,
    },
    provider: {
      transport: provider.transport,
      endpoint: provider.endpoint,
      status_endpoint: provider.status_endpoint,
      role: provider.provider_role,
      required_for_realtime: false,
      security_verifier: provider.security?.verifier,
      missing_env: provider.security?.missing_env ?? [],
      ready: provider.readiness?.ready === true,
      start_events: contract.provider_observer?.events?.start_events ?? [],
      end_events: contract.provider_observer?.events?.end_events ?? [],
      participant_events: contract.provider_observer?.events?.participant_events ?? [],
      artifact_events: contract.provider_observer?.events?.artifact_events ?? [],
      lifecycle_events: contract.provider_observer?.events?.lifecycle_events ?? [],
    },
    annotations: {
      insert_endpoint: contract.annotations?.endpoints?.insertMark ?? runtimeBundle.host?.endpoints?.insertMark,
      runtime_event_endpoint: runtimeEventPlan.endpoint,
      runtime_event_plan: {
        schema: runtimeEventPlan.schema,
        client_factory: runtimeEventPlan.client_factory,
        supported_actions: runtimeEventPlan.supported_actions,
        action_count: runtimeEventPlan.actions?.length ?? 0,
        realtime_contract: runtimeEventPlan.realtime_contract,
      },
      candidate_observation_runtime_action: 'observe_platform_candidates',
      timestamp_field: 'captured_at_ms',
      provider_events_block_realtime: contract.timebase?.provider_events_block_realtime === true,
      transcript_blocks_realtime: contract.timebase?.transcript_blocks_realtime === true,
      local_observer_may_start_axis: contract.timebase?.local_observer_may_start_axis === true,
    },
    transcript: {
      availability: contract.transcript?.availability,
      source: contract.transcript?.source,
      import_endpoint: contract.transcript?.import_endpoint,
      realtime_dependency: contract.transcript?.realtime_dependency === true,
      blocks_realtime_annotation: runtimeBundle.transcript?.blocks_realtime_annotation === true,
    },
    host: {
      base_url: firstNonEmpty(options.baseUrl, options.base_url),
      endpoints: {
        platform_events: hostBasePath,
        platform_event: `${hostBasePath}/${key.replaceAll('_', '-')}`,
        annotations: '/api/annotations',
        runtime_bundles: '/api/meeting-platform/runtime-bundles',
        runtime_event_plans: '/api/meeting-platform/runtime-event-plans',
        runtime_events: '/api/meeting-platform/runtime-events',
        platform_candidate_observation: '/api/meeting-platform/observe-candidates',
        readiness: '/api/meeting-platform/readiness',
        handoff: '/api/meeting-platform/handoff',
        adapter_contracts: '/api/meeting-platform/contracts',
      },
    },
    sdk: {
      package: '@ai-annotation/meeting-timeline-sdk',
      imports: {
        registry: '@ai-annotation/meeting-timeline-sdk/adapters/platform-registry',
        kit: '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit',
        runtime_bundle: '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-bundle',
        runtime_event: '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-event',
        adapter_contract: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-contract',
        provider_connection: '@ai-annotation/meeting-timeline-sdk/adapters/platform-provider-connection',
        events: capabilities.sdk_modules?.events,
      },
    },
    readiness: {
      contract_accepted: contractAcceptance.accepted === true,
      contract_issue_count: contractAcceptance.issue_count ?? 0,
      provider_required_for_realtime: runtimeBundle.readiness?.provider_required_for_realtime === true,
      transcript_blocks_realtime: runtimeBundle.readiness?.transcript_blocks_realtime === true,
      runtime_ready: runtimeBundle.readiness?.runtime_ready === true,
      candidate_observation_ready: candidateObservationReady,
      missing_items: unique([
        ...(contract.readiness?.missing_items ?? []),
        ...(runtimeBundle.readiness?.missing_items ?? []),
        ...(candidateObservationReady ? [] : ['candidate_observation_not_ready']),
      ]),
    },
    commands: {
      print_registry: `npm run meeting-platform:registry -- --platforms=${key}`,
      print_runtime_bundle: `npm run meeting-platform:runtime-bundle -- --platforms=${key}`,
      print_runtime_event_plan: `npm run meeting-platform:runtime-event-plan -- --platforms=${key}`,
      print_adaptation_package: `npm run meeting-platform:adaptation-package -- --platforms=${key}`,
      verify_contract: `npm run meeting-platform:adapter-contract -- --platforms=${key} --fail-on-rejected=true`,
      verify_handoff: `npm run meeting-platform:handoff-readiness -- --platforms=${key}`,
    },
    next_actions: unique([
      ...(runtimeBundle.next_actions ?? []),
      ...(contract.next_actions ?? []),
      'choose_platform_from_registry_manifest',
      'wire_host_runtime_bundles_endpoint_before_building_extension',
      'export_runtime_event_plan_before_wiring_external_host',
      'verify_candidate_observation_before_host_handoff',
    ]),
  });
}

export function buildMeetingPlatformRegistryManifest(options = {}) {
  const entries = selectedPlatforms(options).map((platform) => buildMeetingPlatformRegistryEntry(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_registry_manifest',
    schema: MEETING_PLATFORM_REGISTRY_MANIFEST_SCHEMA,
    schema_version: MEETING_PLATFORM_REGISTRY_SCHEMA_VERSION,
    platform_count: entries.length,
    normalizer_count: entries.filter((entry) => entry.event_adapter?.normalize_available).length,
    runtime_ready_count: entries.filter((entry) => entry.readiness?.runtime_ready).length,
    contract_accepted_count: entries.filter((entry) => entry.readiness?.contract_accepted).length,
    candidate_observer_count: entries.filter((entry) => entry.readiness?.candidate_observation_ready).length,
    provider_required_for_realtime_count: entries.filter((entry) => entry.readiness?.provider_required_for_realtime).length,
    transcript_blocking_count: entries.filter((entry) => entry.readiness?.transcript_blocks_realtime).length,
    platforms: entries.map((entry) => entry.platform),
    rows: entries.map((entry) => ({
      platform: entry.platform,
      display_name: entry.display_name,
      aliases: entry.aliases,
      normalize_available: entry.event_adapter?.normalize_available === true,
      runtime_ready: entry.readiness?.runtime_ready === true,
      contract_accepted: entry.readiness?.contract_accepted === true,
      browser_match_count: entry.runtime?.browser_matches?.length ?? 0,
      candidate_observation_ready: entry.readiness?.candidate_observation_ready === true,
      candidate_observer_message_type: entry.runtime?.candidate_observation?.message_type,
      candidate_observer_permission: entry.runtime?.candidate_observation?.required_permission,
      candidate_observer_endpoint: entry.runtime?.candidate_observation?.endpoint,
      runtime_event_action_count: entry.runtime?.runtime_event_action_count ?? entry.annotations?.runtime_event_plan?.action_count ?? 0,
      provider_transport: entry.provider?.transport,
      provider_ready: entry.provider?.ready === true,
      provider_required_for_realtime: entry.readiness?.provider_required_for_realtime === true,
      transcript_blocks_realtime: entry.readiness?.transcript_blocks_realtime === true,
      insert_endpoint: entry.annotations?.insert_endpoint,
      timestamp_field: entry.annotations?.timestamp_field,
    })),
    entries,
    next_actions: unique(entries.flatMap((entry) => entry.next_actions ?? [])),
  };
}

function registryManifestFrom(input = {}, options = {}) {
  if (input?.schema === MEETING_PLATFORM_REGISTRY_MANIFEST_SCHEMA) return input;
  if (input?.schema === MEETING_PLATFORM_REGISTRY_ENTRY_SCHEMA) {
    return {
      type: 'meeting_platform_registry_manifest',
      schema: MEETING_PLATFORM_REGISTRY_MANIFEST_SCHEMA,
      schema_version: MEETING_PLATFORM_REGISTRY_SCHEMA_VERSION,
      platform_count: 1,
      normalizer_count: input.event_adapter?.normalize_available === true ? 1 : 0,
      runtime_ready_count: input.readiness?.runtime_ready === true ? 1 : 0,
      contract_accepted_count: input.readiness?.contract_accepted === true ? 1 : 0,
      candidate_observer_count: input.readiness?.candidate_observation_ready === true ? 1 : 0,
      provider_required_for_realtime_count: input.readiness?.provider_required_for_realtime === true ? 1 : 0,
      transcript_blocking_count: input.readiness?.transcript_blocks_realtime === true ? 1 : 0,
      platforms: [input.platform],
      rows: [{
        platform: input.platform,
        display_name: input.display_name,
        aliases: input.aliases,
        normalize_available: input.event_adapter?.normalize_available === true,
        runtime_ready: input.readiness?.runtime_ready === true,
        contract_accepted: input.readiness?.contract_accepted === true,
        browser_match_count: input.runtime?.browser_matches?.length ?? 0,
        candidate_observation_ready: input.readiness?.candidate_observation_ready === true,
        candidate_observer_message_type: input.runtime?.candidate_observation?.message_type,
        candidate_observer_permission: input.runtime?.candidate_observation?.required_permission,
        candidate_observer_endpoint: input.runtime?.candidate_observation?.endpoint,
        runtime_event_action_count: input.runtime?.runtime_event_action_count ?? input.annotations?.runtime_event_plan?.action_count ?? 0,
        provider_transport: input.provider?.transport,
        provider_ready: input.provider?.ready === true,
        provider_required_for_realtime: input.readiness?.provider_required_for_realtime === true,
        transcript_blocks_realtime: input.readiness?.transcript_blocks_realtime === true,
        insert_endpoint: input.annotations?.insert_endpoint,
        timestamp_field: input.annotations?.timestamp_field,
      }],
      entries: [input],
      next_actions: input.next_actions ?? [],
    };
  }
  return buildMeetingPlatformRegistryManifest({ ...input, ...options });
}

function registryAcceptanceIssues(manifest = {}, options = {}) {
  const issues = [];
  const entries = asArray(manifest.entries);
  const requireProviderReady = options.requireProviderReady === true || options.require_provider_ready === true;
  if (manifest.schema !== MEETING_PLATFORM_REGISTRY_MANIFEST_SCHEMA) {
    issues.push(issue('error', 'invalid_registry_schema', 'Registry manifest schema must be meeting_platform_registry_manifest.', {
      actual_schema: manifest.schema,
    }));
  }
  if ((manifest.platform_count ?? 0) <= 0) {
    issues.push(issue('error', 'empty_registry_manifest', 'Registry manifest must include at least one platform.'));
  }
  if (manifest.normalizer_count !== manifest.platform_count) {
    issues.push(issue('error', 'missing_platform_normalizer', 'Every selected platform must have a normalizer.', {
      normalizer_count: manifest.normalizer_count,
      platform_count: manifest.platform_count,
    }));
  }
  if (manifest.runtime_ready_count !== manifest.platform_count) {
    issues.push(issue('error', 'runtime_not_ready', 'Every selected platform must have a runtime-ready bundle.', {
      runtime_ready_count: manifest.runtime_ready_count,
      platform_count: manifest.platform_count,
    }));
  }
  if (manifest.contract_accepted_count !== manifest.platform_count) {
    issues.push(issue('error', 'adapter_contract_not_accepted', 'Every selected platform must pass adapter contract acceptance.', {
      contract_accepted_count: manifest.contract_accepted_count,
      platform_count: manifest.platform_count,
    }));
  }
  if (manifest.candidate_observer_count !== manifest.platform_count) {
    issues.push(issue('error', 'candidate_observer_not_ready', 'Every selected platform must expose candidate observation for host-level axis binding.', {
      candidate_observer_count: manifest.candidate_observer_count,
      platform_count: manifest.platform_count,
    }));
  }
  if ((manifest.provider_required_for_realtime_count ?? 0) > 0) {
    issues.push(issue('error', 'provider_blocks_realtime', 'Provider events must not be required for realtime annotation insertion.', {
      provider_required_for_realtime_count: manifest.provider_required_for_realtime_count,
    }));
  }
  if ((manifest.transcript_blocking_count ?? 0) > 0) {
    issues.push(issue('error', 'transcript_blocks_realtime', 'Transcript import must not block realtime annotation insertion.', {
      transcript_blocking_count: manifest.transcript_blocking_count,
    }));
  }
  for (const entry of entries) {
    const platform = entry.platform;
    if (entry.event_adapter?.normalize_available !== true) {
      issues.push(issue('error', 'entry_missing_normalizer', 'Registry entry must expose a normalizer.', { platform }));
    }
    if (entry.readiness?.runtime_ready !== true) {
      issues.push(issue('error', 'entry_runtime_not_ready', 'Registry entry must have runtime_ready=true.', { platform }));
    }
    if (entry.readiness?.contract_accepted !== true) {
      issues.push(issue('error', 'entry_contract_not_accepted', 'Registry entry must have contract_accepted=true.', { platform }));
    }
    if (entry.annotations?.timestamp_field !== 'captured_at_ms') {
      issues.push(issue('error', 'entry_invalid_timestamp_field', 'Registry entry must use captured_at_ms for realtime annotations.', {
        platform,
        timestamp_field: entry.annotations?.timestamp_field,
      }));
    }
    if (!entry.annotations?.insert_endpoint) {
      issues.push(issue('error', 'entry_missing_insert_endpoint', 'Registry entry must include an annotation insert endpoint.', { platform }));
    }
    if (!entry.annotations?.runtime_event_endpoint) {
      issues.push(issue('error', 'entry_missing_runtime_event_endpoint', 'Registry entry must include the runtime event endpoint.', { platform }));
    }
    if (!entry.host?.endpoints?.platform_candidate_observation) {
      issues.push(issue('error', 'entry_missing_candidate_observation_endpoint', 'Registry entry must include the host candidate observation endpoint.', { platform }));
    }
    if (!entry.host?.endpoints?.runtime_event_plans) {
      issues.push(issue('error', 'entry_missing_runtime_event_plan_endpoint', 'Registry entry must include the host runtime event plan endpoint.', { platform }));
    }
    if (!entry.sdk?.imports?.runtime_event) {
      issues.push(issue('error', 'entry_missing_runtime_event_import', 'Registry entry must include the runtime event SDK import path.', { platform }));
    }
    if (!entry.annotations?.runtime_event_plan?.supported_actions?.includes?.('insert_annotation')) {
      issues.push(issue('error', 'entry_missing_runtime_insert_action', 'Registry entry runtime event plan must include insert_annotation.', { platform }));
    }
    if (!entry.annotations?.runtime_event_plan?.supported_actions?.includes?.('observe_platform_candidates')) {
      issues.push(issue('error', 'entry_missing_candidate_observation_action', 'Registry entry runtime event plan must include observe_platform_candidates.', { platform }));
    }
    if (entry.readiness?.candidate_observation_ready !== true) {
      issues.push(issue('error', 'entry_candidate_observation_not_ready', 'Registry entry must expose candidate observation for host-level axis binding.', { platform }));
    }
    if (entry.runtime?.candidate_observation?.message_type !== 'meeting_timeline.observe_candidates') {
      issues.push(issue('error', 'entry_invalid_candidate_observation_message_type', 'Registry entry must use meeting_timeline.observe_candidates for candidate observation.', {
        platform,
        message_type: entry.runtime?.candidate_observation?.message_type,
      }));
    }
    if (entry.runtime?.candidate_observation?.required_permission !== 'tabs') {
      issues.push(issue('error', 'entry_invalid_candidate_observation_permission', 'Registry entry must declare tabs permission for browser candidate observation.', {
        platform,
        required_permission: entry.runtime?.candidate_observation?.required_permission,
      }));
    }
    if (entry.annotations?.provider_events_block_realtime === true) {
      issues.push(issue('error', 'entry_provider_blocks_realtime', 'Registry entry must keep provider events non-blocking for realtime annotations.', { platform }));
    }
    if (entry.annotations?.transcript_blocks_realtime === true || entry.transcript?.blocks_realtime_annotation === true) {
      issues.push(issue('error', 'entry_transcript_blocks_realtime', 'Registry entry must keep transcript import non-blocking for realtime annotations.', { platform }));
    }
    if (entry.supported_surfaces?.browser_observer === true && (entry.runtime?.browser_matches?.length ?? 0) === 0) {
      issues.push(issue('error', 'entry_missing_browser_matches', 'Browser-observed platforms must include runtime browser match patterns.', { platform }));
    }
    if (!entry.host?.endpoints?.runtime_bundles) {
      issues.push(issue('error', 'entry_missing_runtime_bundle_endpoint', 'Registry entry must include the host runtime bundle endpoint.', { platform }));
    }
    if (!entry.sdk?.imports?.runtime_bundle) {
      issues.push(issue('error', 'entry_missing_runtime_bundle_import', 'Registry entry must include the runtime bundle SDK import path.', { platform }));
    }
    if (requireProviderReady && entry.provider?.ready !== true) {
      issues.push(issue('error', 'entry_provider_not_ready', 'Provider setup must be ready when requireProviderReady=true.', {
        platform,
        missing_env: entry.provider?.missing_env ?? [],
      }));
    }
  }
  return issues;
}

export function buildMeetingPlatformRegistryAcceptanceReport(manifestOrOptions = {}, options = {}) {
  const manifest = registryManifestFrom(manifestOrOptions, options);
  const issues = registryAcceptanceIssues(manifest, options);
  const blocking = issues.filter((item) => item.severity === 'error');
  return {
    type: 'meeting_platform_registry_acceptance',
    schema: MEETING_PLATFORM_REGISTRY_ACCEPTANCE_SCHEMA,
    schema_version: MEETING_PLATFORM_REGISTRY_SCHEMA_VERSION,
    accepted: blocking.length === 0,
    platform_count: manifest.platform_count ?? 0,
    normalizer_count: manifest.normalizer_count ?? 0,
    runtime_ready_count: manifest.runtime_ready_count ?? 0,
    contract_accepted_count: manifest.contract_accepted_count ?? 0,
    candidate_observer_count: manifest.candidate_observer_count ?? 0,
    provider_required_for_realtime_count: manifest.provider_required_for_realtime_count ?? 0,
    transcript_blocking_count: manifest.transcript_blocking_count ?? 0,
    blocking_count: blocking.length,
    warning_count: issues.length - blocking.length,
    issues,
    manifest,
    next_actions: unique([
      ...blocking.map((item) => item.code),
      ...(manifest.next_actions ?? []),
    ]),
  };
}

export function assertMeetingPlatformRegistryManifest(manifestOrOptions = {}, options = {}) {
  const report = buildMeetingPlatformRegistryAcceptanceReport(manifestOrOptions, options);
  if (!report.accepted) {
    throw new MeetingTimelineSdkError('Meeting platform registry manifest failed acceptance', {
      issues: report.issues,
      report,
    });
  }
  return report;
}
