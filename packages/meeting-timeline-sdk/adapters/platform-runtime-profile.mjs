import { compactObject } from '../index.mjs';
import {
  MEETING_PLATFORM_KEYS,
  buildPlatformIntegrationPlan,
  normalizeMeetingPlatform,
  platformCapabilityContract,
} from './platform-setup.mjs';
import { buildMeetingPlatformProviderConnectionPack } from './platform-provider-connection.mjs';
import { buildMeetingPlatformAdaptationStrategy } from './platform-strategy.mjs';

export const MEETING_PLATFORM_RUNTIME_PROFILE_SCHEMA = 'meeting_platform_runtime_profile';
export const MEETING_PLATFORM_RUNTIME_PROFILE_MATRIX_SCHEMA = 'meeting_platform_runtime_profile_matrix';
export const MEETING_PLATFORM_RUNTIME_PROFILE_SCHEMA_VERSION = 1;

const SPEAKER_FILTER_DEFAULTS = Object.freeze({
  local_detector: Object.freeze({
    sample_interval_ms: 200,
    min_stable_ms: 300,
    switch_stable_ms: 500,
    end_idle_ms: 1_200,
    dedupe_window_ms: 600,
    merge_gap_ms: 900,
    min_segment_ms: 500,
    emit_speaker_end: true,
  }),
  lark: Object.freeze({
    sample_interval_ms: 300,
    min_stable_ms: 700,
    switch_stable_ms: 1_000,
    end_idle_ms: 2_000,
    dedupe_window_ms: 900,
    merge_gap_ms: 1_500,
    min_segment_ms: 700,
    emit_speaker_end: true,
  }),
  google_meet: Object.freeze({
    sample_interval_ms: 300,
    min_stable_ms: 700,
    switch_stable_ms: 1_000,
    end_idle_ms: 2_000,
    dedupe_window_ms: 900,
    merge_gap_ms: 1_500,
    min_segment_ms: 700,
    emit_speaker_end: true,
  }),
  microsoft_teams: Object.freeze({
    sample_interval_ms: 300,
    min_stable_ms: 800,
    switch_stable_ms: 1_200,
    end_idle_ms: 2_200,
    dedupe_window_ms: 1_000,
    merge_gap_ms: 1_600,
    min_segment_ms: 800,
    emit_speaker_end: true,
  }),
  zoom: Object.freeze({
    sample_interval_ms: 250,
    min_stable_ms: 600,
    switch_stable_ms: 900,
    end_idle_ms: 1_800,
    dedupe_window_ms: 800,
    merge_gap_ms: 1_200,
    min_segment_ms: 650,
    emit_speaker_end: true,
  }),
  webex: Object.freeze({
    sample_interval_ms: 300,
    min_stable_ms: 700,
    switch_stable_ms: 1_000,
    end_idle_ms: 2_000,
    dedupe_window_ms: 900,
    merge_gap_ms: 1_500,
    min_segment_ms: 700,
    emit_speaker_end: true,
  }),
});

const END_FALLBACKS = Object.freeze({
  local_detector: Object.freeze([
    'trusted_host_meeting_ended_signal',
    'manual_stop_or_explicit_session_close',
  ]),
  lark: Object.freeze([
    'local_observer_detects_meeting_window_inactive_or_closed',
    'current_user_leaves_call_or_call_controls_disappear',
    'manual_stop_when_provider_end_event_is_missing',
  ]),
  google_meet: Object.freeze([
    'browser_tab_leaves_meet_url_or_meeting_dom_inactive',
    'native_window_or_pwa_call_controls_disappear',
    'manual_stop_when_workspace_event_is_late',
  ]),
  microsoft_teams: Object.freeze([
    'teams_call_stage_leaves_active_meeting',
    'native_window_call_controls_disappear',
    'manual_stop_when_graph_notification_is_late',
  ]),
  zoom: Object.freeze([
    'zoom_meeting_window_closes_or_leave_button_state_changes',
    'browser_or_native_observer_reports_meeting_inactive',
    'manual_stop_when_zoom_webhook_is_late',
  ]),
  webex: Object.freeze([
    'webex_meeting_window_closes_or_call_controls_disappear',
    'browser_or_native_observer_reports_meeting_inactive',
    'manual_stop_when_webex_webhook_is_late',
  ]),
});

const SURFACE_RECOMMENDATIONS = Object.freeze({
  local_detector: Object.freeze({
    primary: 'host_detector',
    recommended_order: Object.freeze(['host_detector', 'manual_controller']),
    launch_context: 'trusted_host_or_device_runtime',
    realtime_axis_surface: 'host_detector',
    provider_reconcile_surface: null,
    candidate_detection: Object.freeze(['explicit_host_signal']),
    rationale: 'host_detector_is_the_meeting_axis_source_and_does_not_need_a_browser_or_provider_adapter',
  }),
  lark: Object.freeze({
    primary: 'browser_extension_or_desktop_observer',
    recommended_order: Object.freeze(['browser_extension', 'desktop_observer', 'provider_reconcile']),
    launch_context: 'browser_tab_or_desktop_client',
    realtime_axis_surface: 'local_observer',
    provider_reconcile_surface: 'lark_long_connection_or_event_callback',
    candidate_detection: Object.freeze(['url_match', 'tabs_permission', 'dom_preflight', 'current_user_meeting_scan_optional']),
    rationale: 'lark_events_are_useful_for_reconcile_but_local_observer_should_create_the_live_axis_first',
  }),
  google_meet: Object.freeze({
    primary: 'browser_extension',
    recommended_order: Object.freeze(['browser_extension', 'desktop_observer', 'provider_reconcile']),
    launch_context: 'browser_tab_or_pwa',
    realtime_axis_surface: 'local_observer',
    provider_reconcile_surface: 'google_workspace_events_pubsub',
    candidate_detection: Object.freeze(['url_match', 'tabs_permission', 'dom_preflight']),
    rationale: 'meet_is_browser_first_and_workspace_events_should_reconcile_after_the_local_axis_exists',
  }),
  microsoft_teams: Object.freeze({
    primary: 'desktop_or_browser_observer',
    recommended_order: Object.freeze(['desktop_observer', 'browser_extension', 'provider_reconcile']),
    launch_context: 'desktop_client_or_browser_tab',
    realtime_axis_surface: 'local_observer',
    provider_reconcile_surface: 'microsoft_graph_change_notifications',
    candidate_detection: Object.freeze(['native_window_match', 'url_match', 'tabs_permission', 'dom_preflight']),
    rationale: 'teams_usage_often_moves_between_desktop_and_browser_so_the_host_should_keep_both_observer_surfaces_available',
  }),
  zoom: Object.freeze({
    primary: 'native_detector',
    recommended_order: Object.freeze(['native_detector', 'browser_extension', 'provider_reconcile']),
    launch_context: 'desktop_client_or_browser_tab',
    realtime_axis_surface: 'local_observer',
    provider_reconcile_surface: 'zoom_meeting_webhooks',
    candidate_detection: Object.freeze(['native_window_match', 'url_match', 'tabs_permission']),
    rationale: 'zoom_native_client_is_common_so_native_window_detection_should_be_available_before_falling_back_to_browser_observation',
  }),
  webex: Object.freeze({
    primary: 'browser_extension_or_native_detector',
    recommended_order: Object.freeze(['browser_extension', 'native_detector', 'provider_reconcile']),
    launch_context: 'browser_tab_or_desktop_client',
    realtime_axis_surface: 'local_observer',
    provider_reconcile_surface: 'webex_webhooks',
    candidate_detection: Object.freeze(['url_match', 'native_window_match', 'tabs_permission', 'dom_preflight']),
    rationale: 'webex_can_run_in_browser_or_desktop_client_and_provider_webhooks_should_not_block_the_live_axis',
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
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_PLATFORM_KEYS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

function eventMappingFor(pack = {}, normalizedSignal) {
  return (pack.event_mapping ?? [])
    .filter((item) => item.normalized_signal === normalizedSignal)
    .map((item) => item.provider_event);
}

function eventMappingByRole(pack = {}, role) {
  return (pack.event_mapping ?? [])
    .filter((item) => item.timeline_role === role || String(item.timeline_role ?? '').includes(role))
    .map((item) => item.provider_event);
}

function speakerFilterFor(platform, options = {}) {
  const configured = firstNonEmpty(
    options.speakerFilter,
    options.speaker_filter,
    options.activeSpeakerFilter,
    options.active_speaker_filter,
  ) ?? {};
  const byPlatform = configured[platform] ?? {};
  const platformKeys = new Set(MEETING_PLATFORM_KEYS);
  const globalConfig = Object.fromEntries(
    Object.entries(configured).filter(([key]) => !platformKeys.has(key)),
  );
  return {
    ...SPEAKER_FILTER_DEFAULTS[platform],
    ...globalConfig,
    ...byPlatform,
  };
}

function axisStartSource(platform) {
  if (platform === 'local_detector') return 'local_detector_meeting_started';
  return 'local_observer_active_meeting_detected';
}

function axisEndSource(platform) {
  if (platform === 'local_detector') return 'local_detector_meeting_ended';
  return 'local_observer_meeting_inactive_or_manual_stop';
}

function endTimeoutPolicy(platform) {
  if (platform === 'local_detector') {
    return {
      mode: 'host_authoritative',
      grace_ms: 0,
      provider_reconcile_required: false,
    };
  }
  return {
    mode: 'local_end_first_provider_reconcile_later',
    observer_inactive_grace_ms: 3_000,
    manual_stop_allowed: true,
    provider_reconcile_required_for_production_evidence: true,
  };
}

function surfaceRecommendationFor(platform) {
  return SURFACE_RECOMMENDATIONS[platform] ?? SURFACE_RECOMMENDATIONS.local_detector;
}

function launchRequirements(platform, surface = {}, strategy = {}) {
  const localPlatform = platform === 'local_detector';
  return compactObject({
    primary_surface: surface.primary,
    required_before_realtime_annotation: localPlatform
      ? ['host_detector_absolute_started_at_ms']
      : ['candidate_observation', 'local_axis_start_or_pending_axis'],
    required_message_types: localPlatform
      ? []
      : [
        'meeting_timeline.observe_candidates',
        'meeting_timeline.preflight_current_window',
        'meeting_timeline.candidate_launch_plan',
        'meeting_timeline.open_candidate_session',
      ],
    browser_permissions: localPlatform ? [] : ['tabs', 'storage'],
    timestamp_field: 'captured_at_ms',
    may_insert_before_provider_start_event: true,
    provider_events_block_launch: false,
    transcript_blocks_launch: false,
    preflight: {
      required_for_auto_launch: !localPlatform,
      accepted_when: localPlatform
        ? 'host_detector_selected'
        : 'candidate_has_supported_platform_and_local_observer_can_emit_absolute_timestamps',
      recommended_options: {
        requireSpeakerTrack: false,
        requireProviderConnection: false,
      },
    },
    local_observer: {
      required: true,
      source: strategy.local_observer?.implementation,
      evidence_input: strategy.local_observer?.evidence_input,
    },
  });
}

function evidenceThresholds(platform, provider = {}, strategy = {}) {
  const localPlatform = platform === 'local_detector';
  return compactObject({
    pilot: {
      required: localPlatform
        ? ['host_detector_meeting_started', 'host_detector_meeting_ended_or_manual_stop']
        : ['active_meeting_observed', 'annotation_with_captured_at_ms', 'meeting_end_fallback_observed_or_manual_stop'],
      provider_records_required: false,
      transcript_required: false,
      pass_condition: 'ready_for_realtime_annotations === true',
    },
    production: {
      required: localPlatform
        ? ['local_detector_start_end_records']
        : ['meetingAppRecordSet', 'providerRecords', 'end_fallback_or_provider_end_evidence'],
      provider_records_required: !localPlatform,
      transcript_required: false,
      pass_condition: 'production_ready === true',
    },
    provider_connection: {
      required_for_realtime: false,
      required_for_production_evidence: !localPlatform,
      readiness: provider.readiness,
      missing_env: provider.security?.missing_env ?? [],
    },
    sdk_gates: [
      '@ai-annotation/meeting-timeline-sdk/adapters/platform-field-capture',
      '@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-package',
    ],
    current_strategy_status: {
      rollout_status: strategy.rollout_status,
      production_ready: strategy.production_ready,
      ready_for_realtime_annotations: strategy.ready_for_realtime_annotations,
    },
  });
}

function fallbackPolicy(platform, surface = {}) {
  const localPlatform = platform === 'local_detector';
  return {
    start: localPlatform
      ? ['manual_controller_can_create_axis_if_host_signal_is_missing']
      : ['create_pending_axis_from_annotation_capture_time', 'promote_axis_when_local_observer_confirms_active_meeting'],
    end: END_FALLBACKS[platform] ?? [],
    provider_late_or_missing: localPlatform
      ? 'no_provider_dependency'
      : 'keep_realtime_marks_on_captured_at_ms_and_reconcile_axis_bounds_later',
    speaker_track_missing: 'omit_speaker_markers_or_backfill_from_post_meeting_artifacts',
    transcript_missing: 'do_not_block_realtime_annotations',
    next_surface_on_preflight_failure: surface.recommended_order?.[1] ?? null,
  };
}

export function buildMeetingPlatformRuntimeProfile(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const capabilities = platformCapabilityContract(key, options);
  const integration = buildPlatformIntegrationPlan(key, options);
  const strategy = buildMeetingPlatformAdaptationStrategy(key, options);
  const provider = buildMeetingPlatformProviderConnectionPack(key, options);
  const surface = surfaceRecommendationFor(key);
  const startEvents = eventMappingFor(provider, 'meeting_started');
  const endEvents = eventMappingFor(provider, 'meeting_ended');
  const participantEvents = [
    ...eventMappingFor(provider, 'participant_joined'),
    ...eventMappingFor(provider, 'participant_left'),
    ...eventMappingFor(provider, 'participant_joined_or_left'),
  ];
  const artifactEvents = eventMappingByRole(provider, 'post_meeting');
  const speakerFilter = speakerFilterFor(key, options);

  return compactObject({
    type: 'meeting_platform_runtime_profile',
    schema: MEETING_PLATFORM_RUNTIME_PROFILE_SCHEMA,
    schema_version: MEETING_PLATFORM_RUNTIME_PROFILE_SCHEMA_VERSION,
    platform: key,
    display_name: capabilities.display_name ?? integration.display_name,
    objective: 'insert_human_annotations_to_the_current_meeting_axis_with_low_latency',
    runtime_contract: {
      annotation_timestamp_field: 'captured_at_ms',
      axis_timebase: 'absolute_unix_ms',
      relative_mark_time_allowed: false,
      per_meeting_annotation_isolation_required: true,
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      can_insert_annotation_before_provider_start_event: true,
    },
    axis: {
      primary_source: strategy.realtime_axis?.primary_source,
      start: {
        create_on: axisStartSource(key),
        provider_reconcile_events: startEvents,
        fallback_allowed: true,
      },
      end: {
        create_on: axisEndSource(key),
        provider_reconcile_events: endEvents,
        fallbacks: END_FALLBACKS[key] ?? [],
        timeout_policy: endTimeoutPolicy(key),
      },
      provider_reconcile: {
        role: provider.provider_role,
        transport: provider.transport,
        endpoint: provider.endpoint,
        status_endpoint: provider.status_endpoint,
        required_for_realtime: false,
        required_for_production_evidence: key !== 'local_detector',
      },
    },
    adapter_surfaces: {
      primary: surface.primary,
      recommended_order: [...surface.recommended_order],
      launch_context: surface.launch_context,
      realtime_axis_surface: surface.realtime_axis_surface,
      provider_reconcile_surface: surface.provider_reconcile_surface,
      candidate_detection: [...surface.candidate_detection],
      rationale: surface.rationale,
    },
    launch_requirements: launchRequirements(key, surface, strategy),
    evidence_thresholds: evidenceThresholds(key, provider, strategy),
    fallback_policy: fallbackPolicy(key, surface),
    annotations: {
      insert_policy: 'always_place_by_captured_at_ms_on_active_axis',
      out_of_order_policy: 'keep_mark_time_then_recompute_axis_bounds_when_provider_events_arrive',
      required_fields: ['captured_at_ms', 'kind_or_type', 'source'],
      recommended_fields: ['id', 'text_candidates', 'target_region', 'strokes'],
    },
    speaker_markers: {
      enabled: true,
      content_policy: 'speaker_position_only_no_transcript_text_required',
      primary_source: strategy.speaker_activity?.realtime_primary,
      provider_support: strategy.speaker_activity?.provider_support,
      provider_realtime_required: false,
      event_types: ['speaker_started', 'speaker_ended'],
      sdk_module: '@ai-annotation/meeting-timeline-sdk/adapters/active-speaker',
      filter: speakerFilter,
      marker_payload_contract: {
        kind: 'speaker',
        timestamp_field: 'captured_at_ms',
        stable_start_field: 'occurred_at_ms',
        optional_fields: ['speaker_id', 'speaker_name', 'participant_id', 'participant_name'],
      },
      backfill: strategy.speaker_activity?.backfill,
    },
    provider_events: {
      start_events: startEvents,
      end_events: endEvents,
      participant_events: participantEvents,
      artifact_events: artifactEvents,
      lifecycle_events: provider.subscription?.lifecycle_event_types ?? [],
      readiness: provider.readiness,
      missing_env: provider.security?.missing_env ?? [],
    },
    transcript: {
      realtime_dependency: false,
      availability: capabilities.post_meeting_transcript?.availability,
      source: capabilities.post_meeting_transcript?.source,
      import_endpoint: capabilities.post_meeting_transcript?.import_endpoint,
      normalizer: capabilities.post_meeting_transcript?.sdk_normalizer,
    },
    implementation_order: [
      'open_or_detect_meeting_axis_from_local_observer',
      'insert_annotations_immediately_with_captured_at_ms',
      'emit_speaker_position_markers_after_active_speaker_filter',
      'ingest_provider_events_for_start_end_participant_reconcile',
      'import_transcript_or_recording_after_meeting_only',
    ],
    risks: unique([
      ...(capabilities.limitations ?? []),
      provider.readiness?.ready === true ? undefined : 'provider_connection_not_ready',
      'provider_event_latency_must_not_move_realtime_marks_off_their_captured_time',
    ]),
    next_actions: unique([
      ...(strategy.next_actions ?? []),
      ...(provider.next_actions ?? []),
      'validate_runtime_profile_with_real_meeting_capture',
    ]),
  });
}

export function buildMeetingPlatformRuntimeProfileMatrix(options = {}) {
  const profiles = selectedPlatforms(options).map((platform) => buildMeetingPlatformRuntimeProfile(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_runtime_profile_matrix',
    schema: MEETING_PLATFORM_RUNTIME_PROFILE_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_RUNTIME_PROFILE_SCHEMA_VERSION,
    platform_count: profiles.length,
    non_blocking_provider_count: profiles.filter((profile) => profile.runtime_contract.provider_events_block_realtime === false).length,
    transcript_non_blocking_count: profiles.filter((profile) => profile.runtime_contract.transcript_blocks_realtime === false).length,
    speaker_marker_enabled_count: profiles.filter((profile) => profile.speaker_markers.enabled === true).length,
    platforms: profiles.map((profile) => profile.platform),
    rows: profiles.map((profile) => ({
      platform: profile.platform,
      display_name: profile.display_name,
      primary_axis_source: profile.axis.primary_source,
      primary_surface: profile.adapter_surfaces.primary,
      surface_order: profile.adapter_surfaces.recommended_order,
      provider_reconcile_surface: profile.adapter_surfaces.provider_reconcile_surface,
      start_create_on: profile.axis.start.create_on,
      end_create_on: profile.axis.end.create_on,
      provider_start_events: profile.axis.start.provider_reconcile_events,
      provider_end_events: profile.axis.end.provider_reconcile_events,
      pilot_provider_records_required: profile.evidence_thresholds.pilot.provider_records_required,
      production_provider_records_required: profile.evidence_thresholds.production.provider_records_required,
      next_surface_on_preflight_failure: profile.fallback_policy.next_surface_on_preflight_failure,
      speaker_primary_source: profile.speaker_markers.primary_source,
      speaker_min_stable_ms: profile.speaker_markers.filter.min_stable_ms,
      speaker_switch_stable_ms: profile.speaker_markers.filter.switch_stable_ms,
      provider_events_block_realtime: profile.runtime_contract.provider_events_block_realtime,
      transcript_blocks_realtime: profile.runtime_contract.transcript_blocks_realtime,
    })),
    profiles,
  };
}
