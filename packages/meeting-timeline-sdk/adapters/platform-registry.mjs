import { compactObject } from '../index.mjs';
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
  MEETING_PLATFORM_ALIASES,
  MEETING_PLATFORM_KEYS,
  normalizeMeetingPlatform,
  platformCapabilityContract,
} from './platform-setup.mjs';
import { normalizeWebexEvent } from './webex.mjs';
import { normalizeZoomEvent } from './zoom.mjs';

export const MEETING_PLATFORM_REGISTRY_ENTRY_SCHEMA = 'meeting_platform_registry_entry';
export const MEETING_PLATFORM_REGISTRY_MANIFEST_SCHEMA = 'meeting_platform_registry_manifest';
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
      missing_items: unique([
        ...(contract.readiness?.missing_items ?? []),
        ...(runtimeBundle.readiness?.missing_items ?? []),
      ]),
    },
    commands: {
      print_registry: `npm run meeting-platform:registry -- --platforms=${key}`,
      print_runtime_bundle: `npm run meeting-platform:runtime-bundle -- --platforms=${key}`,
      print_adaptation_package: `npm run meeting-platform:adaptation-package -- --platforms=${key}`,
      verify_contract: `npm run meeting-platform:adapter-contract -- --platforms=${key} --fail-on-rejected=true`,
      verify_handoff: `npm run meeting-platform:handoff-readiness -- --platforms=${key}`,
    },
    next_actions: unique([
      ...(runtimeBundle.next_actions ?? []),
      ...(contract.next_actions ?? []),
      'choose_platform_from_registry_manifest',
      'wire_host_runtime_bundles_endpoint_before_building_extension',
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
