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

function implementationState(handoff = {}) {
  return compactObject({
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
  const browserMatches = plan.browser_surface?.matches ?? [];
  const providerDocs = docs(providerConnection);
  return compactObject({
    type: 'meeting_platform_adapter_portfolio_item',
    schema: MEETING_PLATFORM_ADAPTER_PORTFOLIO_ITEM_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_PORTFOLIO_SCHEMA_VERSION,
    platform: plan.platform,
    display_name: plan.display_name,
    built_in: plan.built_in,
    adapter_status: plan.built_in ? 'built_in_adapter_available' : 'adapter_authoring_required',
    recommended_first_surface: plan.recommended_first_surface,
    p0_realtime_axis: {
      source: plan.recommended_first_surface,
      local_axis_first: true,
      browser_matches: browserMatches,
      browser_match_count: browserMatches.length,
      candidate_observation_required: true,
      candidate_observation_message_type: plan.required_contracts?.candidate_observation_message_type,
      timestamp_field: plan.required_contracts?.timestamp_field,
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
    p2_post_meeting_backfill: postMeetingBackfill(providerConnection, handoff),
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
    recommended_first_surface: item.recommended_first_surface,
    browser_match_count: item.p0_realtime_axis?.browser_match_count ?? 0,
    provider_path: item.p1_provider_reconcile?.path,
    provider_transport: item.p1_provider_reconcile?.transport,
    official_doc_count: item.p1_provider_reconcile?.official_doc_count ?? 0,
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
    provider_reconcile_count: items.filter((item) => item.p1_provider_reconcile?.path).length,
    implementation_ready_count: items.filter((item) => item.implementation?.implementation_ready === true).length,
    pilot_ready_count: items.filter((item) => item.implementation?.pilot_ready === true).length,
    production_ready_count: items.filter((item) => item.implementation?.production_ready === true).length,
    platforms: items.map((item) => item.platform),
    rows,
    items,
    next_actions: unique(items.flatMap((item) => item.next_actions ?? [])),
  };
}
