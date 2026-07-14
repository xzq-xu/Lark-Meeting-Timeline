import { compactObject } from '../index.mjs';
import {
  buildMeetingPlatformAdapterAuthoringPlan,
} from './platform-adapter-authoring.mjs';
import {
  buildMeetingPlatformImplementationHandoff,
} from './platform-implementation-handoff.mjs';
import {
  buildMeetingPlatformProviderConnectionPack,
} from './platform-provider-connection.mjs';
import {
  buildMeetingPlatformRuntimeProfile,
} from './platform-runtime-profile.mjs';

export const MEETING_PLATFORM_ADAPTER_PORTFOLIO_ITEM_SCHEMA = 'meeting_platform_adapter_portfolio_item';
export const MEETING_PLATFORM_ADAPTER_PORTFOLIO_SCHEMA = 'meeting_platform_adapter_portfolio';
export const MEETING_PLATFORM_ADAPTER_PORTFOLIO_SCHEMA_VERSION = 1;

const DEFAULT_PORTFOLIO_PLATFORMS = Object.freeze([
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
    DEFAULT_PORTFOLIO_PLATFORMS,
  )));
}

function optionalProviderConnection(platform, options = {}, builtIn = false) {
  if (!builtIn) return undefined;
  try {
    return buildMeetingPlatformProviderConnectionPack(platform, options);
  } catch {
    return undefined;
  }
}

function optionalImplementationHandoff(platform, options = {}, builtIn = false) {
  if (!builtIn) return undefined;
  try {
    return buildMeetingPlatformImplementationHandoff(platform, options);
  } catch {
    return undefined;
  }
}

function optionalRuntimeProfile(platform, options = {}, builtIn = false) {
  if (!builtIn) return undefined;
  try {
    return buildMeetingPlatformRuntimeProfile(platform, options);
  } catch {
    return undefined;
  }
}

function normalizeSurface(surface) {
  return surface == null ? undefined : String(surface).replace(/-/g, '_');
}

function launchableSurface(surface) {
  const normalized = normalizeSurface(surface);
  if (normalized === 'desktop_observer') return 'native_detector';
  if (normalized === 'desktop_or_browser_observer') return 'native_detector';
  if (normalized === 'browser_extension_or_desktop_observer') return 'browser_extension';
  if (normalized === 'browser_extension_or_native_detector') return 'browser_extension';
  return normalized;
}

function recommendedFirstSurface(runtimeProfile = {}, fallback) {
  const surfaces = runtimeProfile.adapter_surfaces ?? {};
  const ordered = unique([
    ...(surfaces.recommended_order ?? []),
    surfaces.primary,
    fallback,
  ].map((surface) => launchableSurface(surface)));
  return ordered.find((surface) => surface && surface !== 'provider_reconcile') ?? fallback;
}

function docs(providerConnection = {}) {
  return (providerConnection.official_docs ?? []).map((doc) => compactObject({
    label: doc.label,
    url: doc.url,
    note: doc.note,
  }));
}

function eventRoles(providerConnection = {}) {
  const roles = new Map();
  for (const event of providerConnection.event_mapping ?? []) {
    const role = event.timeline_role ?? 'unknown';
    roles.set(role, (roles.get(role) ?? 0) + 1);
  }
  return [...roles.entries()].map(([role, count]) => ({ role, count }));
}

function postMeetingBackfill(providerConnection = {}, handoff = {}) {
  const transcript = providerConnection.integration?.post_meeting_transcript ?? {};
  return compactObject({
    realtime_dependency: false,
    transcript_blocks_realtime: false,
    strategy: transcript.strategy ?? handoff.adaptation_strategy?.post_meeting_transcript?.strategy,
    import_module: transcript.import_module,
    source: transcript.source,
    supported: Boolean(transcript.strategy || transcript.source || handoff.production_gaps),
  });
}

function surfaceFamily(surface) {
  const normalized = normalizeSurface(surface);
  if (normalized === 'browser_extension' || normalized === 'webview_preload') return 'browser_observer';
  if (normalized === 'native_detector' || normalized === 'host_sdk') return 'host_native_observer';
  if (normalized === 'provider_reconcile') return 'provider_reconcile';
  if (normalized?.includes('browser')) return 'browser_observer';
  if (normalized?.includes('native') || normalized?.includes('desktop')) return 'host_native_observer';
  return normalized ?? 'unknown';
}

function speakerStrategy(providerConnection = {}) {
  const speakerEvents = (providerConnection.event_mapping ?? [])
    .filter((event) => String(event.timeline_role ?? '').includes('speaker'));
  return {
    realtime_source: 'local_observer_or_host_detector',
    official_provider_realtime_supported: speakerEvents.length > 0,
    provider_event_count: speakerEvents.length,
    fallback: 'post_meeting_transcript_segments',
    blocks_realtime: false,
  };
}

function participantStrategy(providerConnection = {}) {
  const participantEvents = (providerConnection.event_mapping ?? [])
    .filter((event) => String(event.timeline_role ?? '').includes('participant'));
  return {
    realtime_source: participantEvents.length > 0 ? 'provider_or_local_observer_best_effort' : 'local_observer_best_effort',
    provider_event_count: participantEvents.length,
    required_for_realtime_annotation: false,
    blocks_realtime: false,
  };
}

function adapterStrategy(plan = {}, providerConnection = {}, handoff = {}, runtimeProfile = {}, firstSurface, backfill = {}) {
  const timestampField = plan.required_contracts?.timestamp_field ?? 'captured_at_ms';
  return compactObject({
    priority_tier: handoff.priority_tier,
    rank_hint: handoff.rank_hint,
    recommended_first_surface: firstSurface,
    first_surface_family: surfaceFamily(firstSurface),
    launch_context: runtimeProfile.adapter_surfaces?.launch_context,
    realtime_axis: {
      owner: 'local_observer',
      source: firstSurface,
      timestamp_field: timestampField,
      candidate_observation_required: true,
      may_start_before_provider_event: true,
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
    },
    provider_reconcile: compactObject({
      owner: 'official_provider_webhook_or_event_subscription',
      path: plan.provider_reconcile?.path,
      transport: plan.provider_reconcile?.transport,
      role: plan.provider_reconcile?.role ?? 'reconcile_or_backfill_after_local_axis_exists',
      required_for_realtime: false,
      required_for_production: Boolean(plan.provider_reconcile?.path),
      event_mapping_count: providerConnection.event_mapping?.length ?? 0,
      security_verifier: providerConnection.security?.verifier,
      blocks_realtime: false,
    }),
    speaker_positions: speakerStrategy(providerConnection),
    participant_positions: participantStrategy(providerConnection),
    post_meeting_artifacts: compactObject({
      transcript_supported: backfill.supported === true,
      transcript_strategy: backfill.strategy,
      transcript_source: backfill.source,
      transcript_blocks_realtime: false,
      recording_or_artifact_supported: Boolean(providerConnection.integration?.recording?.source),
      required_for_realtime: false,
    }),
    pilot_gate: {
      required: [
        'candidate_observation',
        'annotation_with_captured_at_ms',
        'manual_or_observed_meeting_end',
      ],
      provider_records_required: false,
      transcript_required: false,
    },
    production_gate: {
      required: [
        'meeting_app_evidence',
        'provider_replay_or_webhook_evidence',
        'nonblocking_transcript_import',
      ],
      provider_records_required: true,
      transcript_required: false,
      current_next_phase: handoff.next_phase,
    },
  });
}

function implementationState(handoff = {}) {
  return compactObject({
    priority_tier: handoff.priority_tier,
    rank_hint: handoff.rank_hint,
    implementation_ready: handoff.implementation_ready,
    pilot_ready: handoff.pilot_ready,
    production_ready: handoff.production_ready,
    recommended_first_surface: handoff.recommended_first_surface,
    next_phase: handoff.next_phase,
    next_action: handoff.next_action,
    production_gaps: handoff.production_gaps,
  });
}

function nextActions(plan = {}, providerConnection = {}, handoff = {}) {
  return unique([
    ...(plan.next_actions ?? []),
    ...(providerConnection.next_actions ?? []),
    ...(handoff.next_actions ?? []),
  ]);
}

export function buildMeetingPlatformAdapterPortfolioItem(platform, options = {}) {
  const plan = buildMeetingPlatformAdapterAuthoringPlan(platform, options);
  const providerConnection = optionalProviderConnection(plan.platform, options, plan.built_in);
  const handoff = optionalImplementationHandoff(plan.platform, options, plan.built_in);
  const runtimeProfile = optionalRuntimeProfile(plan.platform, options, plan.built_in);
  const firstSurface = recommendedFirstSurface(runtimeProfile, plan.recommended_first_surface);
  const browserMatches = plan.browser_surface?.matches ?? [];
  const providerDocs = docs(providerConnection);
  const backfill = postMeetingBackfill(providerConnection, handoff);
  const strategy = adapterStrategy(plan, providerConnection, handoff, runtimeProfile, firstSurface, backfill);
  return compactObject({
    type: 'meeting_platform_adapter_portfolio_item',
    schema: MEETING_PLATFORM_ADAPTER_PORTFOLIO_ITEM_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_PORTFOLIO_SCHEMA_VERSION,
    platform: plan.platform,
    display_name: plan.display_name,
    built_in: plan.built_in,
    adapter_status: plan.built_in ? 'built_in_adapter_available' : 'adapter_authoring_required',
    recommended_first_surface: firstSurface,
    adapter_surfaces: runtimeProfile?.adapter_surfaces,
    launch_requirements: runtimeProfile?.launch_requirements,
    evidence_thresholds: runtimeProfile?.evidence_thresholds,
    fallback_policy: runtimeProfile?.fallback_policy,
    adapter_strategy: strategy,
    p0_realtime_axis: {
      source: firstSurface,
      local_axis_first: true,
      browser_matches: browserMatches,
      browser_match_count: browserMatches.length,
      candidate_observation_required: true,
      candidate_observation_message_type: plan.required_contracts?.candidate_observation_message_type,
      timestamp_field: plan.required_contracts?.timestamp_field,
      surface_order: runtimeProfile?.adapter_surfaces?.recommended_order,
      launch_context: runtimeProfile?.adapter_surfaces?.launch_context,
      realtime_axis_surface: runtimeProfile?.adapter_surfaces?.realtime_axis_surface,
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
    },
    p1_provider_reconcile: compactObject({
      path: plan.provider_reconcile?.path,
      transport: plan.provider_reconcile?.transport,
      role: plan.provider_reconcile?.role,
      required_for_realtime: false,
      official_docs: providerDocs,
      official_doc_count: providerDocs.length,
      event_mapping_count: providerConnection?.event_mapping?.length,
      event_roles: eventRoles(providerConnection),
      security: providerConnection?.security,
      readiness: providerConnection?.readiness,
    }),
    p2_post_meeting_backfill: backfill,
    implementation: handoff ? implementationState(handoff) : undefined,
    evidence_requirements: plan.production_evidence,
    commands: compactObject({
      adapter_authoring: plan.commands?.authoring_plan,
      implementation_handoff: plan.commands?.implementation_handoff,
      provider_connection: plan.built_in ? `npm run meeting-platform:provider-connection -- --platforms=${plan.platform}` : undefined,
      handoff_readiness: plan.built_in ? `npm run meeting-platform:handoff-readiness -- --platforms=${plan.platform}` : undefined,
      sdk_test: plan.commands?.sdk_test,
    }),
    artifacts: options.includeArtifacts || options.include_artifacts
      ? compactObject({
        authoring_plan: plan,
        provider_connection: providerConnection,
        implementation_handoff: handoff,
      })
      : undefined,
    next_actions: nextActions(plan, providerConnection, handoff),
  });
}

export function buildMeetingPlatformAdapterPortfolio(options = {}) {
  const items = selectedPlatforms(options).map((platform) => buildMeetingPlatformAdapterPortfolioItem(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  const rows = items.map((item) => ({
    platform: item.platform,
    display_name: item.display_name,
    built_in: item.built_in,
    adapter_status: item.adapter_status,
    primary_surface: item.adapter_surfaces?.primary,
    surface_order: item.adapter_surfaces?.recommended_order,
    recommended_first_surface: item.recommended_first_surface,
    first_surface_family: item.adapter_strategy?.first_surface_family,
    priority_tier: item.adapter_strategy?.priority_tier,
    rank_hint: item.adapter_strategy?.rank_hint,
    browser_match_count: item.p0_realtime_axis?.browser_match_count ?? 0,
    provider_path: item.p1_provider_reconcile?.path,
    provider_transport: item.p1_provider_reconcile?.transport,
    official_doc_count: item.p1_provider_reconcile?.official_doc_count ?? 0,
    provider_required_for_realtime: item.adapter_strategy?.provider_reconcile?.required_for_realtime === true,
    pilot_provider_records_required: item.evidence_thresholds?.pilot?.provider_records_required,
    production_provider_records_required: item.evidence_thresholds?.production?.provider_records_required,
    transcript_supported: item.adapter_strategy?.post_meeting_artifacts?.transcript_supported === true,
    transcript_blocks_realtime: item.adapter_strategy?.post_meeting_artifacts?.transcript_blocks_realtime === true,
    speaker_realtime_source: item.adapter_strategy?.speaker_positions?.realtime_source,
    speaker_provider_realtime_supported: item.adapter_strategy?.speaker_positions?.official_provider_realtime_supported === true,
    implementation_ready: item.implementation?.implementation_ready === true,
    pilot_ready: item.implementation?.pilot_ready === true,
    production_ready: item.implementation?.production_ready === true,
    local_axis_first: item.p0_realtime_axis?.local_axis_first === true,
    provider_events_block_realtime: item.p0_realtime_axis?.provider_events_block_realtime === true,
    next_action: item.next_actions?.[0],
  }));
  return {
    type: 'meeting_platform_adapter_portfolio',
    schema: MEETING_PLATFORM_ADAPTER_PORTFOLIO_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_PORTFOLIO_SCHEMA_VERSION,
    platform_count: items.length,
    built_in_count: items.filter((item) => item.built_in).length,
    external_authoring_count: items.filter((item) => !item.built_in).length,
    browser_surface_ready_count: items.filter((item) => (item.p0_realtime_axis?.browser_match_count ?? 0) > 0).length,
    browser_first_count: items.filter((item) => item.adapter_strategy?.first_surface_family === 'browser_observer').length,
    native_first_count: items.filter((item) => item.adapter_strategy?.first_surface_family === 'host_native_observer').length,
    local_axis_first_count: items.filter((item) => item.adapter_strategy?.realtime_axis?.owner === 'local_observer').length,
    provider_reconcile_count: items.filter((item) => item.p1_provider_reconcile?.path).length,
    provider_required_for_realtime_count: items.filter((item) => item.adapter_strategy?.provider_reconcile?.required_for_realtime === true).length,
    post_meeting_transcript_count: items.filter((item) => item.adapter_strategy?.post_meeting_artifacts?.transcript_supported === true).length,
    transcript_blocking_count: items.filter((item) => item.adapter_strategy?.post_meeting_artifacts?.transcript_blocks_realtime === true).length,
    implementation_ready_count: items.filter((item) => item.implementation?.implementation_ready === true).length,
    pilot_ready_count: items.filter((item) => item.implementation?.pilot_ready === true).length,
    production_ready_count: items.filter((item) => item.implementation?.production_ready === true).length,
    platforms: items.map((item) => item.platform),
    rows,
    items,
    next_actions: unique(items.flatMap((item) => item.next_actions ?? [])),
  };
}
