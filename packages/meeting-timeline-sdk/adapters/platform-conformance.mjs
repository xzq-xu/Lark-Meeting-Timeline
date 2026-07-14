import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  buildMeetingPlatformAdapterContract,
  buildMeetingPlatformAdapterContractAcceptanceReport,
} from './platform-adapter-contract.mjs';
import {
  buildMeetingPlatformAdapterRoute,
  verifyMeetingPlatformAdapterRouteReadiness,
} from './platform-adapter-route.mjs';
import {
  buildMeetingPlatformRegistryAcceptanceReport,
  buildMeetingPlatformRegistryManifest,
  meetingPlatformEventAdapterFor,
} from './platform-registry.mjs';
import {
  buildMeetingPlatformRuntimeBundle,
} from './platform-runtime-bundle.mjs';
import {
  MEETING_PLATFORM_ALIASES,
  MEETING_PLATFORM_KEYS,
  buildPlatformIntegrationPlan,
  normalizeMeetingPlatform,
  platformCapabilityContract,
  platformSetupManifest,
} from './platform-setup.mjs';
import * as transcriptAdapters from './transcript.mjs';

export const MEETING_PLATFORM_CONFORMANCE_REPORT_SCHEMA = 'meeting_platform_conformance_report';
export const MEETING_PLATFORM_CONFORMANCE_SCHEMA_VERSION = 1;

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

function aliasesFor(platform) {
  return Object.entries(MEETING_PLATFORM_ALIASES)
    .filter(([, key]) => key === platform)
    .map(([alias]) => alias);
}

function issue(severity, code, message, details = {}) {
  return compactObject({ severity, code, message, ...details });
}

function platformIssues(platform, inputs = {}) {
  const {
    adapter,
    aliases,
    manifest,
    capability,
    integration,
    contractReport,
    routeReadiness,
    runtimeBundle,
    registryEntry,
    registryRow,
  } = inputs;
  const transcriptNormalizer = capability?.post_meeting_transcript?.sdk_normalizer;
  return [
    MEETING_PLATFORM_KEYS.includes(platform)
      ? undefined
      : issue('error', 'platform_key_not_registered', 'Platform must be listed in MEETING_PLATFORM_KEYS.', { platform }),
    aliases.length > 0
      ? undefined
      : issue('error', 'platform_alias_missing', 'Platform must expose at least one alias.', { platform }),
    typeof adapter?.normalize === 'function'
      ? undefined
      : issue('error', 'platform_normalizer_missing', 'Platform must expose an event normalizer.', { platform }),
    manifest?.platform === platform && manifest?.endpoint
      ? undefined
      : issue('error', 'setup_manifest_incomplete', 'Platform setup manifest must expose platform and endpoint.', { platform }),
    capability?.platform === platform && capability?.sdk_modules?.events && capability?.sdk_modules?.ingest
      ? undefined
      : issue('error', 'capability_contract_incomplete', 'Capability contract must expose core SDK modules.', { platform }),
    integration?.platform === platform && integration?.realtime_annotations?.required_field === 'captured_at_ms'
      ? undefined
      : issue('error', 'integration_plan_invalid_timebase', 'Integration plan must require captured_at_ms for realtime annotations.', { platform }),
    contractReport?.accepted === true
      ? undefined
      : issue('error', 'adapter_contract_not_accepted', 'Adapter contract must pass static acceptance.', {
        platform,
        issue_count: contractReport?.issue_count,
      }),
    routeReadiness?.ready === true
      ? undefined
      : issue('error', 'adapter_route_not_ready', 'Adapter route must satisfy realtime non-blocking constraints.', {
        platform,
        missing: routeReadiness?.missing ?? [],
      }),
    runtimeBundle?.readiness?.runtime_ready === true
      ? undefined
      : issue('error', 'runtime_bundle_not_ready', 'Runtime bundle must be runtime-ready.', { platform }),
    registryEntry
      ? undefined
      : issue('error', 'registry_entry_missing', 'Registry manifest must expose this platform.', { platform }),
    registryRow?.candidate_observation_ready === true
      ? undefined
      : issue('error', 'candidate_observation_not_ready', 'Platform must expose host-level candidate observation.', { platform }),
    registryRow?.timestamp_field === 'captured_at_ms'
      ? undefined
      : issue('error', 'invalid_timestamp_field', 'Registry entry must use captured_at_ms for realtime annotations.', {
        platform,
        timestamp_field: registryRow?.timestamp_field,
      }),
    registryRow?.provider_required_for_realtime === false
      ? undefined
      : issue('error', 'provider_blocks_realtime', 'Provider events must not be required for realtime annotations.', { platform }),
    registryRow?.transcript_blocks_realtime === false
      ? undefined
      : issue('error', 'transcript_blocks_realtime', 'Transcript import must not block realtime annotations.', { platform }),
    transcriptNormalizer && typeof transcriptAdapters[transcriptNormalizer] !== 'function'
      ? issue('error', 'transcript_normalizer_missing', 'Declared transcript normalizer must be exported.', {
        platform,
        transcript_normalizer: transcriptNormalizer,
      })
      : undefined,
  ].filter(Boolean);
}

function platformRow(platform, options = {}, registryManifest = {}) {
  const adapter = meetingPlatformEventAdapterFor(platform);
  const aliases = aliasesFor(platform);
  const manifest = platformSetupManifest(platform, options);
  const capability = platformCapabilityContract(platform, options);
  const integration = buildPlatformIntegrationPlan(platform, options);
  const contract = buildMeetingPlatformAdapterContract(platform, options);
  const contractReport = buildMeetingPlatformAdapterContractAcceptanceReport(contract, {
    ...options,
    target: firstNonEmpty(options.contractTarget, options.contract_target, 'contract'),
  });
  const route = buildMeetingPlatformAdapterRoute(platform, options);
  const routeReadiness = verifyMeetingPlatformAdapterRouteReadiness(route);
  const runtimeBundle = buildMeetingPlatformRuntimeBundle(platform, options);
  const registryEntry = asArray(registryManifest.entries).find((entry) => entry.platform === platform);
  const registryRow = asArray(registryManifest.rows).find((row) => row.platform === platform);
  const issues = platformIssues(platform, {
    adapter,
    aliases,
    manifest,
    capability,
    integration,
    contractReport,
    routeReadiness,
    runtimeBundle,
    registryEntry,
    registryRow,
  });
  const blocking = issues.filter((item) => item.severity === 'error');
  const transcriptNormalizer = capability.post_meeting_transcript?.sdk_normalizer;
  return compactObject({
    platform,
    display_name: capability.display_name ?? contract.display_name ?? registryEntry?.display_name,
    accepted: blocking.length === 0,
    aliases,
    normalizer_available: typeof adapter?.normalize === 'function',
    setup_manifest_ready: manifest.platform === platform && Boolean(manifest.endpoint),
    capability_contract_ready: capability.platform === platform && Boolean(capability.sdk_modules?.events),
    integration_plan_ready: integration.platform === platform && integration.realtime_annotations?.required_field === 'captured_at_ms',
    contract_accepted: contractReport.accepted === true,
    adapter_route_ready: routeReadiness.ready === true,
    adapter_first_route: route.routes?.[0]?.route,
    runtime_ready: runtimeBundle.readiness?.runtime_ready === true,
    registry_entry_ready: Boolean(registryEntry),
    candidate_observation_ready: registryRow?.candidate_observation_ready === true,
    timestamp_field: registryRow?.timestamp_field,
    provider_required_for_realtime: registryRow?.provider_required_for_realtime === true,
    transcript_blocks_realtime: registryRow?.transcript_blocks_realtime === true,
    transcript_normalizer: transcriptNormalizer,
    transcript_normalizer_available: transcriptNormalizer ? typeof transcriptAdapters[transcriptNormalizer] === 'function' : undefined,
    issue_count: issues.length,
    blocking_count: blocking.length,
    issues,
    commands: {
      registry: `npm run meeting-platform:registry -- --platforms=${platform}`,
      adapter_contract: `npm run meeting-platform:adapter-contract -- --platforms=${platform} --fail-on-rejected=true`,
      adapter_route: `npm run meeting-platform:adapter-route -- --platforms=${platform}`,
      runtime_bundle: `npm run meeting-platform:runtime-bundle -- --platforms=${platform}`,
      handoff_readiness: `npm run meeting-platform:handoff-readiness -- --platforms=${platform}`,
    },
  });
}

export function buildMeetingPlatformConformanceReport(options = {}) {
  const platforms = selectedPlatforms(options);
  const registryManifest = buildMeetingPlatformRegistryManifest({
    ...options,
    platforms,
    platform_keys: undefined,
  });
  const registryAcceptance = buildMeetingPlatformRegistryAcceptanceReport(registryManifest, options);
  const rows = platforms.map((platform) => platformRow(platform, options, registryManifest));
  const blockingIssues = [
    ...rows.flatMap((row) => row.issues ?? []),
    ...(registryAcceptance.issues ?? []),
  ].filter((item) => item.severity === 'error');
  return {
    type: 'meeting_platform_conformance_report',
    schema: MEETING_PLATFORM_CONFORMANCE_REPORT_SCHEMA,
    schema_version: MEETING_PLATFORM_CONFORMANCE_SCHEMA_VERSION,
    accepted: blockingIssues.length === 0,
    platform_count: rows.length,
    accepted_count: rows.filter((row) => row.accepted).length,
    normalizer_count: rows.filter((row) => row.normalizer_available).length,
    setup_manifest_ready_count: rows.filter((row) => row.setup_manifest_ready).length,
    capability_contract_ready_count: rows.filter((row) => row.capability_contract_ready).length,
    integration_plan_ready_count: rows.filter((row) => row.integration_plan_ready).length,
    adapter_contract_accepted_count: rows.filter((row) => row.contract_accepted).length,
    adapter_route_ready_count: rows.filter((row) => row.adapter_route_ready).length,
    runtime_ready_count: rows.filter((row) => row.runtime_ready).length,
    registry_entry_ready_count: rows.filter((row) => row.registry_entry_ready).length,
    candidate_observer_count: rows.filter((row) => row.candidate_observation_ready).length,
    provider_required_for_realtime_count: rows.filter((row) => row.provider_required_for_realtime).length,
    transcript_blocking_count: rows.filter((row) => row.transcript_blocks_realtime).length,
    blocking_count: blockingIssues.length,
    platforms,
    rows,
    registry_acceptance: registryAcceptance,
    next_actions: unique([
      ...blockingIssues.map((item) => item.code),
      ...(registryAcceptance.next_actions ?? []),
    ]),
  };
}

export function assertMeetingPlatformConformanceReport(options = {}) {
  const report = buildMeetingPlatformConformanceReport(options);
  if (!report.accepted) {
    throw new MeetingTimelineSdkError('Meeting platform conformance failed', {
      report,
      issues: report.rows.flatMap((row) => row.issues ?? []),
    });
  }
  return report;
}
