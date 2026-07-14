import { compactObject } from '../index.mjs';
import {
  normalizeMeetingPlatform,
  platformCapabilityContract,
} from './platform-setup.mjs';
import {
  buildMeetingAppRuntimeAdapterConfig,
} from './meeting-app-profile.mjs';
import {
  buildMeetingPlatformAdapterRoute,
  verifyMeetingPlatformAdapterRouteReadiness,
} from './platform-adapter-route.mjs';
import {
  buildMeetingPlatformProviderConnectionPack,
} from './platform-provider-connection.mjs';
import {
  buildMeetingPlatformRuntimeProfile,
} from './platform-runtime-profile.mjs';

export const MEETING_PLATFORM_ADAPTER_BLUEPRINT_SCHEMA = 'meeting_platform_adapter_blueprint';
export const MEETING_PLATFORM_ADAPTER_BLUEPRINT_MATRIX_SCHEMA = 'meeting_platform_adapter_blueprint_matrix';
export const MEETING_PLATFORM_ADAPTER_BLUEPRINT_SCHEMA_VERSION = 1;

const DEFAULT_BLUEPRINT_PLATFORMS = Object.freeze([
  'google_meet',
  'microsoft_teams',
  'zoom',
  'webex',
  'lark',
]);

const SURFACE_CANDIDATE_CONTRACTS = Object.freeze({
  host_detector: Object.freeze({
    evidence_kind: 'host_detector_signal',
    required_fields: Object.freeze(['meeting_started_or_active_state', 'captured_at_ms']),
    optional_fields: Object.freeze(['meeting_id', 'meeting_url', 'participants', 'active_speaker', 'manual_controller_state']),
    start_signal: 'trusted_host_detector_reports_active_meeting',
    end_signal: 'trusted_host_detector_reports_meeting_ended_or_manual_stop',
  }),
  browser_extension: Object.freeze({
    evidence_kind: 'dom_snapshot',
    required_fields: Object.freeze(['url_or_tab_url', 'document_title_or_window_title', 'meeting_controls_or_active_call_text', 'captured_at_ms']),
    optional_fields: Object.freeze(['participants', 'active_speaker', 'meeting_title', 'meeting_id', 'page_visibility']),
    start_signal: 'supported_meeting_url_plus_active_meeting_dom',
    end_signal: 'tab_closed_or_meeting_dom_inactive',
  }),
  native_detector: Object.freeze({
    evidence_kind: 'native_window',
    required_fields: Object.freeze(['platform_or_process_hint', 'window_title', 'active_or_visible_window', 'in_meeting_or_audio_call_active', 'captured_at_ms']),
    optional_fields: Object.freeze(['bundle_id', 'meeting_id', 'participants', 'active_speaker', 'accessibility_controls']),
    start_signal: 'native_window_reports_active_meeting',
    end_signal: 'native_window_missing_or_call_controls_inactive',
  }),
  provider_reconcile: Object.freeze({
    evidence_kind: 'provider_event',
    required_fields: Object.freeze(['event_type', 'provider_meeting_id', 'occurred_at_ms']),
    optional_fields: Object.freeze(['participant_id', 'participant_name', 'artifact_url', 'raw']),
    start_signal: 'provider_meeting_started_event',
    end_signal: 'provider_meeting_ended_event',
  }),
  post_meeting_artifact: Object.freeze({
    evidence_kind: 'artifact',
    required_fields: Object.freeze(['meeting_id_or_artifact_id', 'artifact_kind']),
    optional_fields: Object.freeze(['transcript_entries', 'recording_url', 'speaker_segments']),
    start_signal: null,
    end_signal: null,
  }),
});

const SURFACE_ROLES = Object.freeze({
  host_detector: 'trusted_host_axis_for_device_or_desktop_runtime',
  browser_extension: 'low_latency_local_axis_and_browser_speaker_positions',
  native_detector: 'low_latency_desktop_axis_and_native_speaker_positions',
  provider_reconcile: 'authoritative_start_end_and_artifact_reconcile_after_local_axis',
  post_meeting_artifact: 'transcript_or_recording_backfill_after_meeting',
});

const PLATFORM_NOTES = Object.freeze({
  google_meet: Object.freeze({
    primary_reason: 'google_meet_is_browser_first_so_url_and_dom_evidence_are_the_lowest_friction_realtime_axis',
    provider_caution: 'workspace_events_should_reconcile_axis_bounds_but_must_not_delay_realtime_annotation_insertion',
    production_focus: Object.freeze(['chrome_extension_or_webview_preload', 'workspace_events_subscription', 'post_meeting_transcript_import']),
  }),
  microsoft_teams: Object.freeze({
    primary_reason: 'teams_often_runs_as_desktop_client_so_native_window_evidence_should_be_available_before_browser_fallback',
    provider_caution: 'graph_notifications_are_production_evidence_not_the_required_first_paint_axis',
    production_focus: Object.freeze(['native_window_detector', 'browser_fallback', 'graph_change_notification_reconcile']),
  }),
  zoom: Object.freeze({
    primary_reason: 'zoom_native_client_is_common_and_browser_dom_may_not_exist_for_the_active_call',
    provider_caution: 'zoom_webhooks_can_arrive_late_and_should_not_move_captured_marks',
    production_focus: Object.freeze(['native_detector', 'browser_fallback', 'zoom_webhook_reconcile']),
  }),
  webex: Object.freeze({
    primary_reason: 'webex_can_be_browser_or_desktop_so_both_local_observer_surfaces_should_stay_enabled',
    provider_caution: 'webex_webhooks_are_reconcile_and_artifact_signals_not_realtime_annotation_prerequisites',
    production_focus: Object.freeze(['browser_extension', 'native_detector', 'webex_webhook_reconcile']),
  }),
  lark: Object.freeze({
    primary_reason: 'lark_events_help_reconcile_but_local_observer_or_current_user_scan_is_needed_when_delivery_lags',
    provider_caution: 'long_connection_delivery_and_tenant_permissions_can_be_delayed_or_missing',
    production_focus: Object.freeze(['local_observer', 'lark_event_callback_or_long_connection', 'minutes_import']),
  }),
});

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
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, DEFAULT_BLUEPRINT_PLATFORMS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

function safeRuntimeAdapter(platform, options = {}) {
  try {
    return buildMeetingAppRuntimeAdapterConfig(platform, options);
  } catch {
    return {};
  }
}

function surfaceRecommended(surface, runtimeProfile = {}) {
  return (runtimeProfile.adapter_surfaces?.recommended_order ?? []).includes(surface);
}

function surfacePriority(surface, runtimeProfile = {}) {
  const index = (runtimeProfile.adapter_surfaces?.recommended_order ?? []).indexOf(surface);
  return index < 0 ? undefined : index + 1;
}

function browserSurface(platform, runtimeProfile = {}, runtimeAdapter = {}) {
  return compactObject({
    surface: 'browser_extension',
    recommended: surfaceRecommended('browser_extension', runtimeProfile),
    priority: surfacePriority('browser_extension', runtimeProfile),
    role: SURFACE_ROLES.browser_extension,
    evidence: SURFACE_CANDIDATE_CONTRACTS.browser_extension,
    matches: runtimeAdapter.extension?.matches ?? [],
    permissions: runtimeAdapter.extension?.permissions ?? [],
    host_permissions: runtimeAdapter.extension?.host_permissions ?? [],
    launch_when: 'candidate_url_matches_platform_and_dom_preflight_accepts_active_meeting',
    use_for: [
      'meeting_started_axis_when_active_meeting_dom_is_seen',
      'meeting_ended_fallback_when_tab_or_call_controls_disappear',
      'speaker_position_markers_when_active_speaker_can_be_sampled',
    ],
    platform,
  });
}

function hostDetectorSurface(platform, runtimeProfile = {}, route = {}) {
  return compactObject({
    surface: 'host_detector',
    recommended: surfaceRecommended('host_detector', runtimeProfile),
    priority: surfacePriority('host_detector', runtimeProfile),
    role: SURFACE_ROLES.host_detector,
    evidence: SURFACE_CANDIDATE_CONTRACTS.host_detector,
    entrypoint: route.entrypoints?.host_detector,
    launch_when: 'trusted_host_detector_provides_absolute_meeting_axis_signal',
    use_for: [
      'meeting_started_axis_when_host_detector_reports_active_meeting',
      'meeting_ended_axis_when_host_detector_reports_end_or_manual_stop',
      'speaker_position_markers_when_host_detector_provides_active_speaker_samples',
    ],
    platform,
  });
}

function nativeSurface(platform, runtimeProfile = {}, route = {}) {
  return compactObject({
    surface: 'native_detector',
    recommended: surfaceRecommended('native_detector', runtimeProfile)
      || surfaceRecommended('desktop_observer', runtimeProfile),
    priority: firstNonEmpty(
      surfacePriority('native_detector', runtimeProfile),
      surfacePriority('desktop_observer', runtimeProfile),
    ),
    role: SURFACE_ROLES.native_detector,
    evidence: SURFACE_CANDIDATE_CONTRACTS.native_detector,
    entrypoint: route.entrypoints?.native_detector,
    launch_when: 'native_candidate_has_platform_process_or_window_hint_and_active_meeting_state',
    use_for: [
      'meeting_started_axis_when_native_window_reports_active_call',
      'meeting_ended_fallback_when_window_or_call_controls_disappear',
      'speaker_position_markers_when_accessibility_or_audio_focus_can_sample_active_speaker',
    ],
    platform,
  });
}

function providerSurface(platform, runtimeProfile = {}, provider = {}, route = {}) {
  return compactObject({
    surface: 'provider_reconcile',
    recommended: surfaceRecommended('provider_reconcile', runtimeProfile),
    priority: surfacePriority('provider_reconcile', runtimeProfile),
    role: SURFACE_ROLES.provider_reconcile,
    evidence: SURFACE_CANDIDATE_CONTRACTS.provider_reconcile,
    transport: provider.transport,
    endpoint: provider.endpoint,
    status_endpoint: provider.status_endpoint,
    start_events: runtimeProfile.provider_events?.start_events ?? [],
    end_events: runtimeProfile.provider_events?.end_events ?? [],
    participant_events: runtimeProfile.provider_events?.participant_events ?? [],
    artifact_events: runtimeProfile.provider_events?.artifact_events ?? [],
    lifecycle_events: runtimeProfile.provider_events?.lifecycle_events ?? [],
    references: route.entrypoints?.provider_webhook?.references,
    required_for_realtime: false,
    blocks_realtime: false,
    use_for: [
      'reconcile_local_axis_start_end_after_delivery',
      'production_evidence',
      'post_meeting_artifact_discovery',
    ],
    platform,
  });
}

function artifactSurface(platform, runtimeProfile = {}, capabilities = {}, route = {}) {
  return compactObject({
    surface: 'post_meeting_artifact',
    recommended: true,
    role: SURFACE_ROLES.post_meeting_artifact,
    evidence: SURFACE_CANDIDATE_CONTRACTS.post_meeting_artifact,
    transcript: runtimeProfile.transcript,
    import_endpoint: capabilities.post_meeting_transcript?.import_endpoint,
    transcript_source: capabilities.post_meeting_transcript?.source,
    route: route.entrypoints?.post_meeting_artifact,
    required_for_realtime: false,
    blocks_realtime: false,
    platform,
  });
}

function issue(code, message, details = {}) {
  return compactObject({ code, message, ...details });
}

function verifyBlueprint(blueprint = {}) {
  const surfaces = blueprint.surfaces ?? {};
  const runtime = blueprint.runtime_contract ?? {};
  const route = blueprint.adapter_route_readiness ?? {};
  const provider = surfaces.provider_reconcile;
  const localSurfaces = [surfaces.host_detector, surfaces.browser_extension, surfaces.native_detector].filter((surface) => surface?.recommended === true);
  const issues = [
    blueprint.schema === MEETING_PLATFORM_ADAPTER_BLUEPRINT_SCHEMA
      ? undefined
      : issue('invalid_blueprint_schema', 'Blueprint must use meeting_platform_adapter_blueprint schema.'),
    route.ready === true
      ? undefined
      : issue('adapter_route_not_ready', 'Adapter route readiness must pass before exposing the blueprint.', {
        missing: route.missing,
      }),
    runtime.annotations_use_absolute_captured_at_ms === true
      ? undefined
      : issue('annotations_must_use_captured_at_ms', 'Timeline marks must be inserted by absolute captured_at_ms.'),
    runtime.provider_events_block_realtime === false
      ? undefined
      : issue('provider_must_not_block_realtime', 'Provider events must not block realtime annotation insertion.'),
    runtime.transcript_blocks_realtime === false
      ? undefined
      : issue('transcript_must_not_block_realtime', 'Transcript import must not block realtime annotation insertion.'),
    localSurfaces.length > 0
      ? undefined
      : issue('missing_local_realtime_surface', 'At least one local observer surface must be recommended.'),
    !provider || provider.blocks_realtime === false
      ? undefined
      : issue('provider_surface_blocks_realtime', 'Provider reconcile surface must be non-blocking for realtime.'),
  ].filter(Boolean);
  return compactObject({
    type: 'meeting_platform_adapter_blueprint_readiness',
    platform: blueprint.platform,
    ready: issues.length === 0,
    missing: issues.map((item) => item.code),
    issues,
  });
}

export function buildMeetingPlatformAdapterBlueprint(platformOrOptions = {}, options = {}) {
  const objectInput = platformOrOptions && typeof platformOrOptions === 'object' && !Array.isArray(platformOrOptions);
  const rawPlatform = objectInput
    ? firstNonEmpty(platformOrOptions.platform, platformOrOptions.provider, platformOrOptions.key)
    : platformOrOptions;
  const merged = objectInput ? { ...platformOrOptions, ...options } : { ...options, platform: rawPlatform };
  const platform = normalizeMeetingPlatform(firstNonEmpty(rawPlatform, merged.platform));
  const capabilities = platformCapabilityContract(platform, merged);
  const runtimeProfile = buildMeetingPlatformRuntimeProfile(platform, merged);
  const route = buildMeetingPlatformAdapterRoute(platform, merged);
  const routeReadiness = verifyMeetingPlatformAdapterRouteReadiness(route);
  const provider = buildMeetingPlatformProviderConnectionPack(platform, merged);
  const runtimeAdapter = platform === 'local_detector' ? {} : safeRuntimeAdapter(platform, merged);
  const notes = PLATFORM_NOTES[platform] ?? {};
  const blueprint = compactObject({
    type: 'meeting_platform_adapter_blueprint',
    schema: MEETING_PLATFORM_ADAPTER_BLUEPRINT_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_BLUEPRINT_SCHEMA_VERSION,
    platform,
    display_name: capabilities.display_name,
    objective: 'let_a_host_project_insert_realtime_human_marks_on_a_meeting_timeline_across_meeting_apps',
    recommended_mode: route.recommended_mode,
    primary_surface: runtimeProfile.adapter_surfaces?.primary,
    surface_order: runtimeProfile.adapter_surfaces?.recommended_order ?? [],
    surfaces: {
      host_detector: platform === 'local_detector' ? hostDetectorSurface(platform, runtimeProfile, route) : undefined,
      browser_extension: platform === 'local_detector' ? undefined : browserSurface(platform, runtimeProfile, runtimeAdapter),
      native_detector: platform === 'local_detector' ? undefined : nativeSurface(platform, runtimeProfile, route),
      provider_reconcile: platform === 'local_detector' ? undefined : providerSurface(platform, runtimeProfile, provider, route),
      post_meeting_artifact: artifactSurface(platform, runtimeProfile, capabilities, route),
    },
    realtime_axis_contract: {
      create_on: runtimeProfile.axis?.start?.create_on,
      end_on: runtimeProfile.axis?.end?.create_on,
      primary_source: runtimeProfile.axis?.primary_source,
      timestamp_field: runtimeProfile.runtime_contract?.annotation_timestamp_field,
      timebase: runtimeProfile.runtime_contract?.axis_timebase,
      may_insert_before_provider_start_event: runtimeProfile.runtime_contract?.can_insert_annotation_before_provider_start_event,
      provider_reconcile_required_for_realtime: false,
      transcript_required_for_realtime: false,
    },
    annotation_contract: {
      timestamp_field: 'captured_at_ms',
      required_fields: runtimeProfile.annotations?.required_fields ?? [],
      recommended_fields: runtimeProfile.annotations?.recommended_fields ?? [],
      per_meeting_isolation_required: runtimeProfile.runtime_contract?.per_meeting_annotation_isolation_required,
      out_of_order_policy: runtimeProfile.annotations?.out_of_order_policy,
    },
    speaker_marker_contract: runtimeProfile.speaker_markers,
    runtime_contract: {
      annotations_use_absolute_captured_at_ms: true,
      provider_events_block_realtime: runtimeProfile.runtime_contract?.provider_events_block_realtime,
      transcript_blocks_realtime: runtimeProfile.runtime_contract?.transcript_blocks_realtime,
      local_observer_first: platform !== 'local_detector',
      provider_reconcile_after_local_axis: platform !== 'local_detector',
    },
    adapter_route_readiness: routeReadiness,
    platform_notes: notes,
    implementation_sequence: [
      'install_primary_local_surface_from_surface_order',
      'observe_candidates_and_open_axis_when_local_evidence_is_active',
      'insert_human_marks_immediately_by_captured_at_ms',
      'emit_speaker_position_markers_from_filtered_local_speaker_samples',
      'ingest_provider_events_as_non_blocking_reconcile',
      'import_post_meeting_artifacts_for_transcript_or_recording_backfill',
    ],
    acceptance_gates: {
      realtime_pilot: [
        'local_candidate_preflight_accepts_active_meeting',
        'annotation_insert_by_captured_at_ms_before_provider_event',
        'meeting_end_fallback_or_manual_stop_observed',
      ],
      production: [
        'live_snapshot_or_native_window_evidence_corpus',
        'provider_reconcile_records_when_available',
        'per_meeting_annotation_isolation_verified',
        'speaker_marker_filter_thresholds_validated',
      ],
    },
    risks: runtimeProfile.risks ?? [],
    next_actions: unique([
      ...(runtimeProfile.next_actions ?? []),
      'wire_platform_adapter_blueprint_into_host_project',
      'capture_real_candidate_evidence_for_primary_surface',
      'run_candidate_preflight_against_real_meetings',
    ]),
  });
  return {
    ...blueprint,
    readiness: verifyBlueprint(blueprint),
  };
}

export function verifyMeetingPlatformAdapterBlueprint(blueprintOrPlatform = {}, options = {}) {
  if (blueprintOrPlatform?.schema === MEETING_PLATFORM_ADAPTER_BLUEPRINT_SCHEMA) {
    return verifyBlueprint(blueprintOrPlatform);
  }
  return verifyBlueprint(buildMeetingPlatformAdapterBlueprint(blueprintOrPlatform, options));
}

export function assertMeetingPlatformAdapterBlueprint(blueprintOrPlatform = {}, options = {}) {
  const blueprint = blueprintOrPlatform?.schema === MEETING_PLATFORM_ADAPTER_BLUEPRINT_SCHEMA
    ? blueprintOrPlatform
    : buildMeetingPlatformAdapterBlueprint(blueprintOrPlatform, options);
  const readiness = verifyBlueprint(blueprint);
  if (!readiness.ready) {
    const error = new Error('Meeting platform adapter blueprint is not ready');
    error.name = 'MeetingTimelineSdkError';
    error.details = { readiness, blueprint };
    throw error;
  }
  return blueprint;
}

export function buildMeetingPlatformAdapterBlueprintMatrix(options = {}) {
  const blueprints = selectedPlatforms(options).map((platform) => buildMeetingPlatformAdapterBlueprint(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_adapter_blueprint_matrix',
    schema: MEETING_PLATFORM_ADAPTER_BLUEPRINT_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_BLUEPRINT_SCHEMA_VERSION,
    platform_count: blueprints.length,
    ready_count: blueprints.filter((blueprint) => blueprint.readiness.ready).length,
    browser_primary_count: blueprints.filter((blueprint) => blueprint.primary_surface === 'browser_extension').length,
    native_primary_count: blueprints.filter((blueprint) => blueprint.primary_surface === 'native_detector' || blueprint.primary_surface === 'desktop_or_browser_observer').length,
    provider_non_blocking_count: blueprints.filter((blueprint) => blueprint.runtime_contract.provider_events_block_realtime === false).length,
    transcript_non_blocking_count: blueprints.filter((blueprint) => blueprint.runtime_contract.transcript_blocks_realtime === false).length,
    platforms: blueprints.map((blueprint) => blueprint.platform),
    rows: blueprints.map((blueprint) => ({
      platform: blueprint.platform,
      display_name: blueprint.display_name,
      ready: blueprint.readiness.ready,
      primary_surface: blueprint.primary_surface,
      surface_order: blueprint.surface_order,
      browser_recommended: blueprint.surfaces.browser_extension?.recommended,
      native_recommended: blueprint.surfaces.native_detector?.recommended,
      provider_blocks_realtime: blueprint.runtime_contract.provider_events_block_realtime,
      transcript_blocks_realtime: blueprint.runtime_contract.transcript_blocks_realtime,
      start_create_on: blueprint.realtime_axis_contract.create_on,
      end_on: blueprint.realtime_axis_contract.end_on,
      first_acceptance_gate: blueprint.acceptance_gates.realtime_pilot[0],
    })),
    blueprints,
  };
}

export function assertMeetingPlatformAdapterBlueprintMatrix(options = {}) {
  const matrix = options?.schema === MEETING_PLATFORM_ADAPTER_BLUEPRINT_MATRIX_SCHEMA
    ? options
    : buildMeetingPlatformAdapterBlueprintMatrix(options);
  const failed = matrix.blueprints.filter((blueprint) => blueprint.readiness.ready !== true);
  if (failed.length > 0) {
    const error = new Error('Meeting platform adapter blueprint matrix is not ready');
    error.name = 'MeetingTimelineSdkError';
    error.details = { failed: failed.map((blueprint) => blueprint.readiness), matrix };
    throw error;
  }
  return matrix;
}
