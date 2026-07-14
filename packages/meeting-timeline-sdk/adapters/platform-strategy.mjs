import { compactObject } from './internal-utils.mjs';
import {
  MEETING_PLATFORM_KEYS,
  buildPlatformIntegrationPlan,
  normalizeMeetingPlatform,
  platformCapabilityContract,
} from './platform-setup.mjs';
import { buildMeetingPlatformRolloutPlan } from './platform-rollout.mjs';

export const MEETING_PLATFORM_ADAPTATION_STRATEGY_SCHEMA = 'meeting_platform_adaptation_strategy';
export const MEETING_PLATFORM_ADAPTATION_STRATEGY_SCHEMA_VERSION = 1;

const PROVIDER_LATENCY_PROFILES = Object.freeze({
  local_detector: {
    class: 'host_realtime',
    realtime_blocking: false,
    policy: 'host_detector_must_emit_absolute_timestamps_as_soon_as_the_meeting_state_changes',
  },
  lark: {
    class: 'provider_event_best_effort',
    realtime_blocking: false,
    policy: 'provider_events_can_lag_or_fail_open_platform_delivery_checks_so_local_axis_should_be_created_first',
  },
  google_meet: {
    class: 'provider_event_best_effort',
    realtime_blocking: false,
    policy: 'workspace_events_and_pubsub_are_authoritative_for_reconcile_but_not_the_low_latency_annotation_clock',
  },
  microsoft_teams: {
    class: 'provider_event_best_effort',
    realtime_blocking: false,
    policy: 'graph_change_notifications_are_authoritative_for_reconcile_but_subscription_delivery_and_renewal_must_not_block_marks',
  },
  zoom: {
    class: 'provider_event_best_effort',
    realtime_blocking: false,
    policy: 'zoom_webhooks_reconcile_meeting_lifecycle_and_recording_artifacts_after_the_local_axis_exists',
  },
  webex: {
    class: 'provider_event_best_effort',
    realtime_blocking: false,
    policy: 'webex_webhooks_reconcile_meeting_lifecycle_and_artifacts_after_the_local_axis_exists',
  },
});

const PROVIDER_INTEGRATION_PATHS = Object.freeze({
  local_detector: {
    path: 'host_detector_runtime',
    transport: 'host_runtime_signal',
    permission_risk: 'none',
    setup_risk: 'timestamp_quality_and_observer_coverage',
  },
  lark: {
    path: 'lark_long_connection_or_event_callback',
    transport: 'long_connection_or_http_callback',
    permission_risk: 'tenant_app_permission_and_event_subscription',
    setup_risk: 'event_delivery_depends_on_app_publication_and_tenant_scope',
  },
  google_meet: {
    path: 'google_workspace_events_pubsub',
    transport: 'pubsub_push_or_pull',
    permission_risk: 'workspace_scope_and_subscription_target',
    setup_risk: 'workspace_event_subscription_scope_and_pubsub_delivery',
  },
  microsoft_teams: {
    path: 'microsoft_graph_change_notifications',
    transport: 'graph_webhook_subscription',
    permission_risk: 'tenant_admin_consent_and_subscription_renewal',
    setup_risk: 'graph_subscription_lifecycle_and_policy_limits',
  },
  zoom: {
    path: 'zoom_meeting_webhooks',
    transport: 'webhook',
    permission_risk: 'account_scope_and_event_subscription',
    setup_risk: 'cloud_recording_or_transcript_settings_for_artifacts',
  },
  webex: {
    path: 'webex_webhooks',
    transport: 'webhook',
    permission_risk: 'workspace_scope_and_admin_scope_variant',
    setup_risk: 'webhook_resource_coverage_and_rest_enrichment',
  },
});

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  return Array.isArray(value) ? value : value == null ? [] : [value];
}

function uniqueList(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function selectedPlatforms(options = {}) {
  return asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_PLATFORM_KEYS));
}

function localObserverStrategy(platform, rollout = {}) {
  if (platform === 'local_detector') {
    return {
      role: 'primary_axis_source',
      required_for_pilot: true,
      implementation: 'trusted_host_detector_or_manual_controller',
      evidence_input: 'providerRecords_or_local_detector_records',
      gate: '@ai-annotation/meeting-timeline-sdk/adapters/platform-gate',
      production_ready: rollout.provider_events?.production_ready === true,
      evidence_level: rollout.provider_events?.evidence_level,
    };
  }
  return {
    role: 'primary_low_latency_axis_source',
    required_for_pilot: true,
    implementation: 'browser_extension_desktop_observer_or_native_app_observer',
    evidence_input: 'meetingAppRecordSet',
    gate: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-gate',
    runtime_ready: rollout.local_observer?.runtime_ready === true,
    production_ready: rollout.local_observer?.production_ready === true,
    evidence_level: rollout.local_observer?.evidence_level,
    missing_required_coverage: rollout.local_observer?.missing_required_coverage,
  };
}

function providerStrategy(platform, integration = {}, rollout = {}) {
  if (platform === 'local_detector') return undefined;
  const providerEvents = integration.provider_events ?? {};
  return {
    role: 'non_blocking_reconcile_and_backfill',
    required_for_pilot: false,
    required_for_production: true,
    latency: PROVIDER_LATENCY_PROFILES[platform],
    transport: providerEvents.transport,
    endpoint: providerEvents.endpoint,
    status_endpoint: providerEvents.status_endpoint,
    event_types: providerEvents.event_types ?? [],
    lifecycle_event_types: providerEvents.lifecycle_event_types ?? [],
    evidence_input: 'providerRecords',
    gate: '@ai-annotation/meeting-timeline-sdk/adapters/platform-gate',
    setup_ready: rollout.provider_events?.setup_ready === true,
    production_ready: rollout.provider_events?.production_ready === true,
    evidence_level: rollout.provider_events?.evidence_level,
    missing_required_coverage: rollout.provider_events?.missing_required_coverage,
  };
}

function speakerStrategy(capabilities = {}, integration = {}) {
  const speaker = capabilities.speaker_activity ?? {};
  const providerSupportsRealtime = !String(speaker.status ?? '').startsWith('not_supported');
  return {
    realtime_primary: providerSupportsRealtime ? 'provider_or_local_observer' : 'local_observer_or_detector',
    provider_support: speaker.status,
    realtime_provider_blocking: false,
    production_gate_required: false,
    signal_types: speaker.signal_types ?? ['speaker_started', 'speaker_ended'],
    backfill: speaker.fallback,
    active_speaker_module: integration.speaker_activity?.active_speaker_module,
  };
}

function transcriptStrategy(capabilities = {}, integration = {}) {
  const transcript = capabilities.post_meeting_transcript ?? {};
  return {
    realtime_dependency: false,
    realtime_blocking: false,
    strategy: integration.post_meeting_transcript?.strategy,
    availability: transcript.availability,
    source: transcript.source,
    import_endpoint: transcript.import_endpoint,
    sdk_normalizer: transcript.sdk_normalizer,
    import_module: integration.post_meeting_transcript?.import_module,
  };
}

function phaseStatus(condition, pending = 'pending_evidence') {
  return condition ? 'ready' : pending;
}

function providerPath(platform) {
  return PROVIDER_INTEGRATION_PATHS[platform] ?? PROVIDER_INTEGRATION_PATHS.local_detector;
}

function supportsPostMeetingBackfill(transcript = {}) {
  const status = String(transcript.status ?? '');
  return status !== 'not_applicable' && status !== 'not_supported';
}

function adaptationPhases(platform, capabilities = {}, integration = {}, rollout = {}) {
  const transcript = capabilities.post_meeting_transcript ?? {};
  const localPlatform = platform === 'local_detector';
  const providerEvents = integration.provider_events ?? {};
  const provider = providerPath(platform);
  const realtimeReady = rollout.ready_for_realtime_annotations === true;
  const providerReady = rollout.provider_events?.production_ready === true;
  const productionReady = rollout.production_ready === true;
  const postMeetingBackfillSupported = supportsPostMeetingBackfill(transcript);

  return [
    compactObject({
      id: 'axis_bootstrap',
      priority: 'P0',
      purpose: 'create_low_latency_meeting_axis_before_any_annotation_is_inserted',
      owner: localPlatform ? 'host_detector' : 'local_observer',
      source: localPlatform ? 'local_detector' : 'browser_extension_desktop_observer_or_native_app_observer',
      required_for_pilot: true,
      required_for_production: true,
      status: phaseStatus(realtimeReady),
      done_when: localPlatform
        ? 'host_emits_meeting_started_and_meeting_ended_with_absolute_ms'
        : 'observer_can_detect_active_meeting_window_and_emit_absolute_start_time',
      evidence_input: localPlatform ? 'local_detector_records' : 'meetingAppRecordSet',
    }),
    compactObject({
      id: 'realtime_annotation_intake',
      priority: 'P0',
      purpose: 'insert_marks_against_the_current_axis_with_device_capture_time',
      owner: 'annotation_host',
      source: 'device_or_browser_runtime',
      required_for_pilot: true,
      required_for_production: true,
      status: phaseStatus(realtimeReady),
      timestamp_field: integration.realtime_annotations?.required_field ?? 'captured_at_ms',
      invariant: integration.realtime_axis?.invariant,
      sdk_method: integration.realtime_annotations?.sdk_method,
    }),
    compactObject({
      id: 'provider_reconcile',
      priority: localPlatform ? 'P3' : 'P1',
      purpose: localPlatform
        ? 'not_needed_for_host_detector_only_axis'
        : 'reconcile_or_backfill_meeting_lifecycle_and_participant_events_without_blocking_marks',
      owner: localPlatform ? 'none' : 'provider_adapter',
      source: localPlatform ? 'none' : provider.path,
      transport: localPlatform ? undefined : provider.transport,
      required_for_pilot: false,
      required_for_production: !localPlatform,
      non_blocking_for_realtime: true,
      status: localPlatform ? 'not_applicable' : phaseStatus(providerReady, 'pending_provider_evidence'),
      event_types: localPlatform ? undefined : providerEvents.event_types,
      lifecycle_event_types: localPlatform ? undefined : providerEvents.lifecycle_event_types,
    }),
    compactObject({
      id: 'speaker_participant_tracks',
      priority: 'P1',
      purpose: 'draw_speaker_or_participant_positions_on_the_same_axis_when_available',
      owner: 'track_pipeline',
      source: capabilities.speaker_activity?.status?.startsWith('not_supported')
        ? 'local_observer_or_post_meeting_transcript'
        : 'provider_events_or_local_observer',
      required_for_pilot: false,
      required_for_production: false,
      status: 'optional_progressive_enhancement',
      participant_support: capabilities.participant_track?.status,
      speaker_support: capabilities.speaker_activity?.status,
      fallback: capabilities.speaker_activity?.fallback,
    }),
    compactObject({
      id: 'post_meeting_artifacts',
      priority: 'P2',
      purpose: 'import_transcript_recording_or_minutes_after_meeting_end_for_backfill',
      owner: 'artifact_importer',
      source: transcript.source,
      required_for_pilot: false,
      required_for_production: false,
      realtime_dependency: false,
      status: postMeetingBackfillSupported ? 'available_after_meeting' : 'provider_specific_or_not_applicable',
      availability: transcript.availability,
      import_endpoint: transcript.import_endpoint,
      sdk_normalizer: transcript.sdk_normalizer,
    }),
    compactObject({
      id: 'production_evidence_gate',
      priority: 'P3',
      purpose: 'prove_real_platform_coverage_before_treating_the_adapter_as_production_ready',
      owner: 'platform_adapter_owner',
      source: 'meeting_platform_evidence_package',
      required_for_pilot: false,
      required_for_production: true,
      status: phaseStatus(productionReady, 'pending_production_evidence'),
      gate: '@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-package',
      requires: localPlatform
        ? ['local_detector_start_end_records']
        : ['meetingAppRecordSet', 'providerRecords'],
    }),
  ];
}

function nextPhaseFrom(phases = []) {
  return phases.find((phase) => (
    phase.required_for_pilot === true
    && phase.status !== 'ready'
    && phase.status !== 'not_applicable'
  )) ?? phases.find((phase) => (
    phase.required_for_production === true
    && phase.status !== 'ready'
    && phase.status !== 'not_applicable'
  )) ?? null;
}

function riskProfile(platform, capabilities = {}, rollout = {}) {
  const provider = providerPath(platform);
  const transcript = capabilities.post_meeting_transcript ?? {};
  const speakerUnsupported = String(capabilities.speaker_activity?.status ?? '').startsWith('not_supported');
  const localPlatform = platform === 'local_detector';
  return {
    provider_path: provider.path,
    provider_transport: provider.transport,
    permission_risk: provider.permission_risk,
    setup_risk: provider.setup_risk,
    event_latency_risk: localPlatform ? 'low' : 'medium_provider_delivery_best_effort',
    provider_axis_risk: localPlatform ? 'not_applicable' : 'must_not_block_realtime_axis',
    speaker_realtime_gap: speakerUnsupported,
    transcript_availability_risk: supportsPostMeetingBackfill(transcript)
      ? 'post_meeting_only'
      : 'provider_specific_or_unavailable',
    current_rollout_status: rollout.status,
    mitigations: uniqueList([
      localPlatform ? 'validate_host_absolute_timestamps' : 'create_local_axis_before_provider_event_arrives',
      speakerUnsupported ? 'use_local_observer_for_realtime_speaker_markers' : undefined,
      supportsPostMeetingBackfill(transcript) ? 'import_transcript_after_meeting_for_backfill' : undefined,
      platform === 'microsoft_teams' ? 'monitor_subscription_renewal_and_lifecycle_events' : undefined,
      platform === 'google_meet' ? 'monitor_workspace_subscription_lifecycle_events' : undefined,
    ]),
  };
}

function evidenceContract(platform, rollout = {}) {
  const productionRequires = platform === 'local_detector'
    ? ['local_detector_start_end_records']
    : ['meetingAppRecordSet', 'providerRecords'];
  return {
    pilot: {
      condition: 'ready_for_realtime_annotations === true',
      accepts_statuses: ['production_ready', 'realtime_ready_provider_pending'],
      requires: platform === 'local_detector'
        ? ['local_detector_start_event']
        : ['meetingAppRecordSet_with_active_and_ended_snapshots_or_provider_start_end_records'],
      current_passed: rollout.ready_for_realtime_annotations === true,
    },
    production: {
      condition: 'production_ready === true',
      requires: productionRequires,
      current_passed: rollout.production_ready === true,
    },
    handoff_package: {
      schema: 'meeting_platform_evidence_package',
      builder: '@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-package',
      verifier: 'verifyMeetingPlatformEvidencePackage',
      current_status: rollout.status,
    },
  };
}

function recommendationFor(platform, rollout = {}) {
  if (rollout.production_ready) return 'enable_pilot_with_provider_reconcile_and_monitoring';
  if (rollout.ready_for_realtime_annotations) return 'enable_local_observer_pilot_collect_provider_evidence';
  if (platform === 'local_detector') return 'capture_host_detector_start_end_events';
  return 'capture_live_local_observer_evidence_then_provider_events';
}

export function buildMeetingPlatformAdaptationStrategy(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const capabilities = platformCapabilityContract(key, options);
  const integration = buildPlatformIntegrationPlan(key, options);
  const rollout = buildMeetingPlatformRolloutPlan(key, options);
  const phases = adaptationPhases(key, capabilities, integration, rollout);
  const nextPhase = nextPhaseFrom(phases);
  return compactObject({
    schema: MEETING_PLATFORM_ADAPTATION_STRATEGY_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTATION_STRATEGY_SCHEMA_VERSION,
    type: 'meeting_platform_adaptation_strategy',
    platform: key,
    display_name: integration.display_name,
    rollout_status: rollout.status,
    production_ready: rollout.production_ready,
    ready_for_realtime_annotations: rollout.ready_for_realtime_annotations,
    recommendation: recommendationFor(key, rollout),
    source_priority: integration.source_priority,
    realtime_axis: {
      primary_source: integration.realtime_axis?.primary,
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      timestamp_invariant: integration.realtime_axis?.invariant,
      annotation_time_field: integration.realtime_annotations?.required_field ?? 'captured_at_ms',
      signal_types: integration.realtime_axis?.signal_types ?? ['meeting_started', 'meeting_ended'],
      delay_policy: PROVIDER_LATENCY_PROFILES[key]?.policy,
    },
    local_observer: localObserverStrategy(key, rollout),
    provider_events: providerStrategy(key, integration, rollout),
    speaker_activity: speakerStrategy(capabilities, integration),
    post_meeting_transcript: transcriptStrategy(capabilities, integration),
    adaptation_playbook: {
      phases,
      next_phase: nextPhase?.id,
      next_phase_priority: nextPhase?.priority,
      integration_path: providerPath(key),
      risk_profile: riskProfile(key, capabilities, rollout),
    },
    evidence_contract: evidenceContract(key, rollout),
    next_actions: uniqueList(rollout.next_actions ?? []),
    rollout_plan: options.includeRolloutPlan === true || options.include_rollout_plan === true
      ? rollout
      : undefined,
  });
}

export function buildAllMeetingPlatformAdaptationStrategies(options = {}) {
  return selectedPlatforms(options).map((platform) => buildMeetingPlatformAdaptationStrategy(platform, options));
}

export function buildMeetingPlatformAdaptationStrategyMatrix(options = {}) {
  const strategies = buildAllMeetingPlatformAdaptationStrategies(options);
  return {
    type: 'meeting_platform_adaptation_strategy_matrix',
    schema: MEETING_PLATFORM_ADAPTATION_STRATEGY_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTATION_STRATEGY_SCHEMA_VERSION,
    strategy_count: strategies.length,
    production_ready_count: strategies.filter((item) => item.production_ready).length,
    realtime_ready_count: strategies.filter((item) => item.ready_for_realtime_annotations).length,
    local_first_count: strategies.filter((item) => item.realtime_axis?.primary_source === 'local_observer').length,
    non_blocking_provider_count: strategies.filter((item) => item.provider_events?.latency?.realtime_blocking === false).length,
    provider_reconcile_required_count: strategies.filter((item) => item.provider_events?.required_for_production === true).length,
    speaker_local_fallback_count: strategies.filter((item) => item.speaker_activity?.realtime_primary === 'local_observer_or_detector').length,
    post_meeting_backfill_count: strategies.filter((item) => (
      item.adaptation_playbook?.risk_profile?.transcript_availability_risk === 'post_meeting_only'
    )).length,
    platforms: strategies.map((item) => item.platform),
    rows: strategies.map((item) => ({
      platform: item.platform,
      display_name: item.display_name,
      rollout_status: item.rollout_status,
      recommendation: item.recommendation,
      next_phase: item.adaptation_playbook?.next_phase,
      provider_path: item.adaptation_playbook?.integration_path?.path,
      permission_risk: item.adaptation_playbook?.risk_profile?.permission_risk,
      primary_axis_source: item.realtime_axis?.primary_source,
      provider_role: item.provider_events?.role ?? item.local_observer?.role,
      provider_reconcile_required: item.provider_events?.required_for_production === true,
      provider_blocks_realtime: item.realtime_axis?.provider_events_block_realtime,
      transcript_blocks_realtime: item.realtime_axis?.transcript_blocks_realtime,
      speaker_realtime_primary: item.speaker_activity?.realtime_primary,
      speaker_realtime_gap: item.adaptation_playbook?.risk_profile?.speaker_realtime_gap,
      post_meeting_backfill_supported: item.adaptation_playbook?.risk_profile?.transcript_availability_risk === 'post_meeting_only',
      production_ready: item.production_ready,
      ready_for_realtime_annotations: item.ready_for_realtime_annotations,
    })),
    strategies,
  };
}
