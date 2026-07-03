import { compactObject } from '../index.mjs';
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
    platforms: strategies.map((item) => item.platform),
    rows: strategies.map((item) => ({
      platform: item.platform,
      display_name: item.display_name,
      rollout_status: item.rollout_status,
      recommendation: item.recommendation,
      primary_axis_source: item.realtime_axis?.primary_source,
      provider_role: item.provider_events?.role ?? item.local_observer?.role,
      provider_blocks_realtime: item.realtime_axis?.provider_events_block_realtime,
      transcript_blocks_realtime: item.realtime_axis?.transcript_blocks_realtime,
      speaker_realtime_primary: item.speaker_activity?.realtime_primary,
      production_ready: item.production_ready,
      ready_for_realtime_annotations: item.ready_for_realtime_annotations,
    })),
    strategies,
  };
}
