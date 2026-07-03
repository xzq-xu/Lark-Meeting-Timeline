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

export function buildMeetingPlatformRuntimeProfile(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const capabilities = platformCapabilityContract(key, options);
  const integration = buildPlatformIntegrationPlan(key, options);
  const strategy = buildMeetingPlatformAdaptationStrategy(key, options);
  const provider = buildMeetingPlatformProviderConnectionPack(key, options);
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
      start_create_on: profile.axis.start.create_on,
      end_create_on: profile.axis.end.create_on,
      provider_start_events: profile.axis.start.provider_reconcile_events,
      provider_end_events: profile.axis.end.provider_reconcile_events,
      speaker_primary_source: profile.speaker_markers.primary_source,
      speaker_min_stable_ms: profile.speaker_markers.filter.min_stable_ms,
      speaker_switch_stable_ms: profile.speaker_markers.filter.switch_stable_ms,
      provider_events_block_realtime: profile.runtime_contract.provider_events_block_realtime,
      transcript_blocks_realtime: profile.runtime_contract.transcript_blocks_realtime,
    })),
    profiles,
  };
}
