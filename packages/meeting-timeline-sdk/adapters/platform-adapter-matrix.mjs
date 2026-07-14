import { compactObject } from './internal-utils.mjs';
import {
  MEETING_PLATFORM_KEYS,
  normalizeMeetingPlatform,
  platformCapabilityContract,
} from './platform-setup.mjs';
import {
  buildMeetingAppRuntimeAdapterConfig,
} from './meeting-app-profile.mjs';
import {
  buildMeetingPlatformAdapterDecision,
} from './platform-adapter-decision.mjs';
import {
  buildMeetingPlatformProviderConnectionPack,
} from './platform-provider-connection.mjs';
import {
  buildMeetingPlatformRuntimeProfile,
} from './platform-runtime-profile.mjs';

export const MEETING_PLATFORM_ADAPTER_MATRIX_ROW_SCHEMA = 'meeting_platform_adapter_matrix_row';
export const MEETING_PLATFORM_ADAPTER_MATRIX_SCHEMA = 'meeting_platform_adapter_matrix';
export const MEETING_PLATFORM_ADAPTER_MATRIX_SCHEMA_VERSION = 1;

const DEFAULT_ADAPTER_MATRIX_PLATFORMS = Object.freeze([
  'google_meet',
  'microsoft_teams',
  'zoom',
  'lark',
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

function selectedPlatforms(input = {}, options = {}) {
  const fallback = options.includeLocalDetector === true || options.include_local_detector === true
    ? MEETING_PLATFORM_KEYS
    : DEFAULT_ADAPTER_MATRIX_PLATFORMS;
  return unique(asArray(firstNonEmpty(
    input.platforms,
    input.platform_keys,
    input.platformKeys,
    options.platforms,
    options.platform_keys,
    options.platformKeys,
    fallback,
  )).map((platform) => normalizeMeetingPlatform(platform)));
}

function inputByPlatform(platform, input = {}, options = {}) {
  const maps = [
    input.inputs,
    input.inputByPlatform,
    input.input_by_platform,
    input.snapshots,
    input.snapshotByPlatform,
    input.snapshot_by_platform,
    options.inputs,
    options.inputByPlatform,
    options.input_by_platform,
    options.snapshots,
    options.snapshotByPlatform,
    options.snapshot_by_platform,
  ].filter(Boolean);
  return maps.reduce((merged, map) => ({
    ...merged,
    ...(map?.[platform] ?? {}),
  }), {});
}

function rowKind(platform, primarySurface, runtimeProfile = {}) {
  if (platform === 'local_detector') return 'host_detector';
  if (primarySurface === 'browser_extension') return 'browser_first';
  if (primarySurface === 'native_detector') return 'native_first';
  if (primarySurface === 'provider_reconcile') return 'provider_reconcile_only';
  const primary = runtimeProfile.adapter_surfaces?.primary;
  if (String(primary || '').includes('desktop')) return 'desktop_first';
  if (String(primary || '').includes('browser') && String(primary || '').includes('native')) return 'hybrid_local_observer';
  return 'local_observer_first';
}

function providerReadiness(provider = {}, runtimeProfile = {}) {
  const readiness = runtimeProfile.provider_events?.readiness ?? provider.readiness ?? {};
  return compactObject({
    ready: readiness.ready,
    warning_count: readiness.warning_count,
    blocking_count: readiness.blocking_count,
    missing_env: runtimeProfile.provider_events?.missing_env ?? provider.missing_env,
    checks: readiness.checks,
  });
}

function browserIntegration(runtimeAdapter = {}) {
  return compactObject({
    matches: runtimeAdapter.extension?.matches ?? [],
    host_permissions: runtimeAdapter.extension?.host_permissions ?? [],
    permissions: runtimeAdapter.extension?.permissions ?? [],
    runtime_preset: runtimeAdapter.runtime_options?.runtimePreset ?? runtimeAdapter.runtime_options?.runtime_preset,
    bridge_source: runtimeAdapter.bridge_options?.source,
  });
}

function providerSummary(provider = {}, runtimeProfile = {}, decision = {}) {
  const route = decision.route_summary?.provider_reconcile ?? {};
  return compactObject({
    required_for_realtime: route.required_for_realtime ?? runtimeProfile.axis?.provider_reconcile?.required_for_realtime,
    required_for_production: route.required_for_production ?? runtimeProfile.axis?.provider_reconcile?.required_for_production_evidence,
    blocks_realtime: route.blocks_realtime_if_missing,
    transport: route.transport ?? runtimeProfile.axis?.provider_reconcile?.transport ?? provider.transport,
    endpoint: provider.endpoint ?? runtimeProfile.axis?.provider_reconcile?.endpoint,
    status_endpoint: provider.status_endpoint ?? runtimeProfile.axis?.provider_reconcile?.status_endpoint,
    start_events: runtimeProfile.provider_events?.start_events ?? [],
    end_events: runtimeProfile.provider_events?.end_events ?? [],
    participant_events: runtimeProfile.provider_events?.participant_events ?? [],
    artifact_events: runtimeProfile.provider_events?.artifact_events ?? [],
    lifecycle_events: runtimeProfile.provider_events?.lifecycle_events ?? [],
    readiness: providerReadiness(provider, runtimeProfile),
  });
}

function actionPlan(platform, decision = {}, runtimeProfile = {}) {
  const surface = decision.selected_surface ?? runtimeProfile.adapter_surfaces?.primary;
  return unique([
    'resolve_platform_adapter_matrix_row',
    surface === 'browser_extension' ? 'install_browser_extension_or_webview_preload' : undefined,
    surface === 'native_detector' ? 'wire_native_window_or_accessibility_detector' : undefined,
    platform === 'local_detector' ? 'wire_trusted_host_detector' : undefined,
    'run_current_window_or_candidate_preflight_before_first_mark',
    'open_local_axis_before_insert_annotation',
    'insert_annotations_with_captured_at_ms',
    runtimeProfile.speaker_markers?.enabled ? 'emit_speaker_position_markers_after_filtering' : undefined,
    runtimeProfile.axis?.provider_reconcile?.required_for_production_evidence ? 'ingest_provider_events_for_nonblocking_reconcile' : undefined,
    runtimeProfile.transcript?.availability === 'post_meeting' ? 'import_post_meeting_transcript_when_available' : undefined,
  ]);
}

export function buildMeetingPlatformAdapterMatrixRow(platform, input = {}, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const platformInput = {
    ...input,
    ...inputByPlatform(key, input, options),
    platform: key,
  };
  const runtimeProfile = buildMeetingPlatformRuntimeProfile(key, options);
  const runtimeAdapter = key === 'local_detector' ? {} : buildMeetingAppRuntimeAdapterConfig(key, options);
  const decision = buildMeetingPlatformAdapterDecision(platformInput, {
    ...options,
    platform: key,
  });
  const provider = key === 'local_detector' ? {} : buildMeetingPlatformProviderConnectionPack(key, options);
  const capabilities = platformCapabilityContract(key, options);
  const primarySurface = decision.selected_surface ?? runtimeProfile.adapter_surfaces?.primary;
  const speaker = runtimeProfile.speaker_markers ?? {};
  const transcript = runtimeProfile.transcript ?? {};
  return compactObject({
    type: 'meeting_platform_adapter_matrix_row',
    schema: MEETING_PLATFORM_ADAPTER_MATRIX_ROW_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_MATRIX_SCHEMA_VERSION,
    platform: key,
    display_name: runtimeProfile.display_name ?? decision.display_name ?? capabilities.display_name,
    adapter_kind: rowKind(key, primarySurface, runtimeProfile),
    status: decision.status,
    realtime_ready: decision.realtime_ready,
    accepted: decision.accepted,
    primary_surface: primarySurface,
    recommended_order: runtimeProfile.adapter_surfaces?.recommended_order ?? [],
    fallback_surfaces: decision.adaptation_strategy?.fallback_surfaces ?? [],
    runtime: compactObject({
      selected_observer_mode: decision.adaptation_strategy?.selected_observer_mode,
      runtime_factory: decision.adaptation_strategy?.runtime_factory,
      host_install_target: decision.adaptation_strategy?.host_install_target,
      open_session_message_type: 'meeting_timeline.open_session',
      runtime_target_message_type: 'meeting_timeline.runtime_target',
    }),
    realtime_axis: compactObject({
      source: runtimeProfile.axis?.primary_source,
      selected_route: decision.selected_route,
      timestamp_field: runtimeProfile.runtime_contract?.annotation_timestamp_field ?? 'captured_at_ms',
      start_create_on: runtimeProfile.axis?.start?.create_on,
      end_create_on: runtimeProfile.axis?.end?.create_on,
      end_fallbacks: runtimeProfile.axis?.end?.fallbacks ?? [],
      end_timeout_policy: runtimeProfile.axis?.end?.timeout_policy,
      can_insert_before_provider_start_event: runtimeProfile.runtime_contract?.can_insert_annotation_before_provider_start_event,
      provider_events_block_realtime: runtimeProfile.runtime_contract?.provider_events_block_realtime,
      transcript_blocks_realtime: runtimeProfile.runtime_contract?.transcript_blocks_realtime,
      per_meeting_annotation_isolation_required: runtimeProfile.runtime_contract?.per_meeting_annotation_isolation_required,
    }),
    browser_integration: browserIntegration(runtimeAdapter),
    provider_reconcile: providerSummary(provider, runtimeProfile, decision),
    speaker_positions: compactObject({
      enabled: speaker.enabled,
      source: speaker.primary_source,
      provider_support: speaker.provider_support,
      provider_realtime_required: speaker.provider_realtime_required,
      content_policy: speaker.content_policy,
      filter: speaker.filter,
      output_contract: speaker.marker_payload_contract,
      backfill: speaker.backfill,
    }),
    post_meeting: compactObject({
      transcript_available: transcript.availability === 'post_meeting',
      transcript_source: transcript.source,
      transcript_import_endpoint: transcript.import_endpoint,
      transcript_normalizer: transcript.normalizer,
      transcript_realtime_dependency: transcript.realtime_dependency,
    }),
    readiness: compactObject({
      host_surface_available: decision.host_surface_compatibility?.selected_surface_available !== false,
      first_blocked_step: decision.first_blocked_step,
      provider_ready: runtimeProfile.provider_events?.readiness?.ready,
      production_provider_required: runtimeProfile.axis?.provider_reconcile?.required_for_production_evidence,
    }),
    evidence_first: decision.adaptation_strategy?.evidence_to_collect_first ?? [],
    risks: capabilities.limitations ?? [],
    next_actions: unique([
      ...actionPlan(key, decision, runtimeProfile),
      ...(decision.next_actions ?? []),
    ]),
    decision: options.includeDecision === true || options.include_decision === true ? decision : undefined,
    runtime_profile: options.includeRuntimeProfile === true || options.include_runtime_profile === true ? runtimeProfile : undefined,
  });
}

export function buildMeetingPlatformAdapterMatrix(input = {}, options = {}) {
  const platforms = selectedPlatforms(input, options);
  const rows = platforms.map((platform) => buildMeetingPlatformAdapterMatrixRow(platform, input, options));
  const providerRows = rows.filter((row) => row.provider_reconcile?.required_for_production === true);
  return {
    type: 'meeting_platform_adapter_matrix',
    schema: MEETING_PLATFORM_ADAPTER_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_MATRIX_SCHEMA_VERSION,
    objective: 'compare_meeting_platform_adapters_for_realtime_annotation_timeline_integration',
    platform_count: rows.length,
    realtime_ready_count: rows.filter((row) => row.realtime_ready === true).length,
    accepted_count: rows.filter((row) => row.accepted === true).length,
    browser_first_count: rows.filter((row) => row.adapter_kind === 'browser_first').length,
    native_first_count: rows.filter((row) => row.adapter_kind === 'native_first' || row.adapter_kind === 'desktop_first').length,
    provider_reconcile_required_count: providerRows.length,
    speaker_position_enabled_count: rows.filter((row) => row.speaker_positions?.enabled === true).length,
    post_meeting_transcript_count: rows.filter((row) => row.post_meeting?.transcript_available === true).length,
    platforms: rows.map((row) => row.platform),
    rows: rows.map((row) => ({
      platform: row.platform,
      display_name: row.display_name,
      adapter_kind: row.adapter_kind,
      realtime_ready: row.realtime_ready,
      accepted: row.accepted,
      primary_surface: row.primary_surface,
      recommended_order: row.recommended_order,
      selected_observer_mode: row.runtime?.selected_observer_mode,
      start_create_on: row.realtime_axis?.start_create_on,
      end_create_on: row.realtime_axis?.end_create_on,
      provider_reconcile_required: row.provider_reconcile?.required_for_production,
      provider_ready: row.provider_reconcile?.readiness?.ready,
      speaker_positions_enabled: row.speaker_positions?.enabled,
      transcript_available: row.post_meeting?.transcript_available,
      first_evidence: row.evidence_first?.[0],
      first_next_action: row.next_actions?.[0],
    })),
    adapters: rows,
    next_actions: unique(rows.flatMap((row) => row.next_actions ?? [])),
  };
}
