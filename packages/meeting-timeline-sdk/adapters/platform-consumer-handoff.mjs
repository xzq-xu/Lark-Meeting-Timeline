import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  normalizeMeetingPlatform,
} from './platform-setup.mjs';
import {
  buildMeetingPlatformAdaptationPackageMatrix,
} from './platform-adaptation-package.mjs';
import {
  buildMeetingPlatformConformanceReport,
} from './platform-conformance.mjs';
import {
  buildMeetingPlatformHandoffReadinessMatrix,
} from './platform-handoff-readiness.mjs';
import {
  buildMeetingPlatformHostIntegrationPlan,
} from './platform-host-integration.mjs';

export const MEETING_PLATFORM_CONSUMER_HANDOFF_SCHEMA = 'meeting_platform_consumer_handoff';
export const MEETING_PLATFORM_CONSUMER_HANDOFF_SCHEMA_VERSION = 1;

const DEFAULT_CONSUMER_PLATFORMS = Object.freeze([
  'lark',
  'google_meet',
  'microsoft_teams',
  'zoom',
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

function selectedPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(
    options.platforms,
    options.platform_keys,
    DEFAULT_CONSUMER_PLATFORMS,
  )).map((platform) => normalizeMeetingPlatform(platform)));
}

function byPlatform(rows = []) {
  return Object.fromEntries(asArray(rows).map((row) => [row.platform, row]));
}

function issue(severity, code, message, details = {}) {
  return compactObject({ severity, code, message, ...details });
}

function boolOption(options = {}, camel, snake, defaultValue = false) {
  return firstNonEmpty(options[camel], options[snake], defaultValue) === true;
}

function commandWithBase(name, options = {}, extra = '') {
  const baseUrl = firstNonEmpty(options.baseUrl, options.base_url);
  const args = [
    baseUrl ? `--base-url=${baseUrl}` : '',
    String(extra).trim(),
  ].filter(Boolean).join(' ');
  return args ? `npm run ${name} -- ${args}` : `npm run ${name}`;
}

function buildConsumerCommands(options = {}) {
  return {
    consumer_handoff: commandWithBase('meeting-platform:consumer-handoff', options),
    app_adapter_manifest: commandWithBase('meeting-app:adapter-manifest', options),
    conformance: commandWithBase('meeting-platform:conformance', options),
    host_integration: commandWithBase('meeting-platform:host-integration', options),
    integration_runtime_manifest: commandWithBase('meeting-platform:integration-runtime-manifest', options),
    runtime_bundles: commandWithBase('meeting-platform:runtime-bundle', options),
    adapter_routes: commandWithBase('meeting-platform:adapter-route', options),
    runtime_event_plans: commandWithBase('meeting-platform:runtime-event-plan', options),
    adaptation_packages: commandWithBase('meeting-platform:adaptation-package', options),
    handoff_readiness: commandWithBase('meeting-platform:handoff-readiness', options),
    runtime_host_replay: commandWithBase('meeting-platform:runtime-host-replay', options),
  };
}

function buildBootOrder(hostPlan = {}, commands = {}) {
  return [
    {
      step: 1,
      action: 'run_static_consumer_handoff',
      command: commands.consumer_handoff,
      output_schema: MEETING_PLATFORM_CONSUMER_HANDOFF_SCHEMA,
    },
    {
      step: 2,
      action: 'create_host_or_kit_runtime',
      sdk_modules: [
        hostPlan.sdk?.kit_module,
        hostPlan.sdk?.integration_runtime_module,
        hostPlan.sdk?.host_integration_module,
      ].filter(Boolean),
      required_methods: [
        'createMeetingPlatformTimelineKit',
        'createMeetingPlatformIntegrationRuntime',
        'createMeetingPlatformHost',
      ],
    },
    {
      step: 3,
      action: 'wire_local_observer_candidate_binding',
      endpoint: hostPlan.endpoints?.platform_candidate_observation,
      runtime_event_action: 'observe_platform_candidates',
      required_before: 'insert_realtime_annotation',
    },
    {
      step: 4,
      action: 'insert_realtime_annotations_by_capture_time',
      endpoint: hostPlan.endpoints?.annotations,
      runtime_event_endpoint: hostPlan.endpoints?.runtime_events,
      required_field: 'captured_at_ms',
    },
    {
      step: 5,
      action: 'emit_speaker_and_participant_positions',
      endpoints: [
        hostPlan.endpoints?.runtime_events,
      ].filter(Boolean),
      source_priority: ['local_observer_samples', 'provider_reconcile', 'post_meeting_artifacts'],
    },
    {
      step: 6,
      action: 'connect_provider_events_as_reconcile_backfill',
      endpoint: hostPlan.endpoints?.platform_events,
      blocks_realtime_annotation: false,
    },
    {
      step: 7,
      action: 'collect_real_evidence_and_run_handoff_readiness',
      command: commands.handoff_readiness,
      required_for: ['pilot', 'production'],
    },
  ];
}

function buildEntrypoints(hostPlan = {}, options = {}) {
  return {
    package: '@ai-annotation/meeting-timeline-sdk',
    primary_modules: {
      platform_kit: hostPlan.sdk?.kit_module,
      integration_runtime: hostPlan.sdk?.integration_runtime_module,
      host_integration: hostPlan.sdk?.host_integration_module,
      app_adapter_manifest: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-manifest',
      app_adapter_spec: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-spec',
      app_adapter_runtime_config: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-runtime-config',
      consumer_handoff: '@ai-annotation/meeting-timeline-sdk/adapters/platform-consumer-handoff',
      conformance: hostPlan.sdk?.platform_conformance_module,
      runtime_event: hostPlan.sdk?.runtime_event_module,
      live_adapter: hostPlan.sdk?.live_adapter_module,
    },
    kit_methods: [
      'platformConsumerHandoff',
      'assertPlatformConsumerHandoff',
      'meetingAppAdapterManifest',
      'meetingAppAdapterManifestMatrix',
      'meetingAppAdapterSpec',
      'meetingAppAdapterSpecMatrix',
      'meetingAppAdapterRuntimeConfig',
      'meetingAppAdapterRuntimeConfigMatrix',
      'platformConformance',
      'platformHostIntegrationPlan',
      'platformRuntimeBundleMatrix',
      'platformAdapterRouteMatrix',
      'platformRuntimeEventPlanMatrix',
      'platformHandoffReadinessMatrix',
    ],
    host_methods: [
      'platformConformance',
      'consumerHandoff',
      'integrationRuntimeManifest',
      'observePlatformCandidates',
      'insertAnnotation',
      'speakerTrack',
      'participantTrack',
      'runHandoffReadiness',
    ],
    http_endpoints: hostPlan.endpoints,
    commands: buildConsumerCommands(options),
  };
}

function buildContracts(hostPlan = {}) {
  return {
    timestamp_field: 'captured_at_ms',
    realtime_axis_source_order: ['local_observer_candidate', 'provider_reconcile', 'post_meeting_transcript_or_artifact'],
    local_observer_may_start_axis: true,
    candidate_observation_required_before_realtime_mark: true,
    provider_events_block_realtime: false,
    transcript_blocks_realtime: false,
    per_meeting_annotation_isolation_required: true,
    speaker_track_required_for_realtime_timeline: true,
    participant_track_required_for_realtime_timeline: true,
    meeting_end_must_close_current_axis: true,
    host_runtime_contract: hostPlan.runtime_contract,
  };
}

function hostPlanSummary(hostPlan = {}) {
  return compactObject({
    type: hostPlan.type,
    schema: hostPlan.schema,
    schema_version: hostPlan.schema_version,
    base_url: hostPlan.base_url,
    base_path: hostPlan.base_path,
    platforms: hostPlan.platforms,
    sdk: hostPlan.sdk,
    runtime_contract: hostPlan.runtime_contract,
    endpoints: hostPlan.endpoints,
    commands: hostPlan.commands,
    candidate_observation_contract: hostPlan.candidate_observation_contract,
    meeting_track_contract: hostPlan.meeting_track_contract,
    platform_conformance_summary: hostPlan.platform_conformance_report
      ? {
        accepted: hostPlan.platform_conformance_report.accepted,
        accepted_count: hostPlan.platform_conformance_report.accepted_count,
        blocking_count: hostPlan.platform_conformance_report.blocking_count,
      }
      : undefined,
    next_actions: hostPlan.next_actions,
  });
}

function platformRow(platform, inputs = {}) {
  const {
    conformanceRows,
    runtimeRows,
    routeRows,
    adaptationRows,
    handoffRows,
    trackRows,
    requireHandoffReady,
    requireProductionReady,
  } = inputs;
  const conformance = conformanceRows[platform] ?? {};
  const runtime = runtimeRows[platform] ?? {};
  const route = routeRows[platform] ?? {};
  const adaptation = adaptationRows[platform] ?? {};
  const handoff = handoffRows[platform] ?? {};
  const track = trackRows[platform] ?? {};
  const candidateReady = runtime.candidate_observation_ready === true
    || conformance.candidate_observation_ready === true
    || handoff.candidate_observation_ready === true;
  const speakerReady = track.speaker_track_ready === true || runtime.speaker_track_ready === true;
  const participantReady = track.participant_track_ready === true || runtime.participant_track_ready === true;
  const issues = [
    conformance.accepted === true
      ? undefined
      : issue('error', 'platform_conformance_not_accepted', 'Static SDK conformance must pass before consumer handoff.', {
        platform,
        blocking_count: conformance.blocking_count,
      }),
    runtime.runtime_ready === true || conformance.runtime_ready === true
      ? undefined
      : issue('error', 'runtime_bundle_not_ready', 'Runtime bundle must be ready for the consumer host.', { platform }),
    route.ready === true || route.adapter_route_ready === true || conformance.adapter_route_ready === true
      ? undefined
      : issue('error', 'adapter_route_not_ready', 'Adapter route must expose the local observer first path.', { platform }),
    candidateReady
      ? undefined
      : issue('error', 'candidate_observation_not_ready', 'Consumer host must be able to observe active meeting candidates before realtime marks.', { platform }),
    speakerReady
      ? undefined
      : issue('error', 'speaker_track_not_ready', 'Consumer host must be able to place speaker positions without provider/transcript blocking.', { platform }),
    participantReady
      ? undefined
      : issue('error', 'participant_track_not_ready', 'Consumer host must be able to place participant positions without provider/transcript blocking.', { platform }),
    runtime.provider_blocks_realtime === true || conformance.provider_required_for_realtime === true
      ? issue('error', 'provider_blocks_realtime', 'Provider events must not be required for realtime annotation placement.', { platform })
      : undefined,
    runtime.transcript_blocks_realtime === true || conformance.transcript_blocks_realtime === true
      ? issue('error', 'transcript_blocks_realtime', 'Transcript import must not block realtime annotation placement.', { platform })
      : undefined,
    requireHandoffReady && handoff.handoff_ready !== true
      ? issue('error', 'handoff_readiness_required', 'requireHandoffReady=true requires real handoff readiness evidence.', { platform })
      : undefined,
    requireProductionReady && handoff.production_ready !== true
      ? issue('error', 'production_ready_required', 'requireProductionReady=true requires production evidence.', { platform })
      : undefined,
    handoff.production_ready === true
      ? undefined
      : issue('warning', 'production_evidence_pending', 'Static consumer handoff is usable, but production evidence is still pending.', {
        platform,
        first_next_action: handoff.first_next_action,
      }),
  ].filter(Boolean);
  const blocking = issues.filter((item) => item.severity === 'error');
  return compactObject({
    platform,
    display_name: conformance.display_name ?? runtime.display_name ?? adaptation.display_name ?? handoff.display_name,
    consumer_ready: blocking.length === 0,
    conformance_accepted: conformance.accepted === true,
    conformance_blocking_count: conformance.blocking_count,
    runtime_ready: runtime.runtime_ready === true || conformance.runtime_ready === true,
    sdk_wiring_ready: runtime.sdk_wiring_ready === true || adaptation.sdk_wiring_ready === true,
    candidate_observation_ready: candidateReady,
    candidate_observer_message_type: runtime.candidate_observer_message_type ?? handoff.candidate_observer_message_type,
    adapter_route_ready: route.ready === true || route.adapter_route_ready === true || conformance.adapter_route_ready === true,
    adapter_first_route: route.first_route ?? route.adapter_first_route ?? conformance.adapter_first_route,
    adapter_recommended_mode: route.recommended_mode ?? runtime.adapter_recommended_mode,
    speaker_track_ready: speakerReady,
    speaker_min_stable_ms: track.speaker_min_stable_ms ?? runtime.speaker_min_stable_ms,
    speaker_switch_stable_ms: track.speaker_switch_stable_ms ?? runtime.speaker_switch_stable_ms,
    speaker_end_idle_ms: track.speaker_end_idle_ms ?? runtime.speaker_end_idle_ms,
    participant_track_ready: participantReady,
    participant_duplicate_window_ms: track.participant_duplicate_window_ms ?? runtime.participant_duplicate_window_ms,
    participant_leave_stable_ms: track.participant_leave_stable_ms ?? runtime.participant_leave_stable_ms,
    provider_blocks_realtime: runtime.provider_blocks_realtime === true || conformance.provider_required_for_realtime === true,
    transcript_blocks_realtime: runtime.transcript_blocks_realtime === true || conformance.transcript_blocks_realtime === true,
    handoff_ready: handoff.handoff_ready === true,
    pilot_ready: handoff.pilot_ready === true,
    production_ready: handoff.production_ready === true,
    first_next_action: handoff.first_next_action ?? runtime.first_next_action ?? adaptation.first_next_action,
    issue_count: issues.length,
    blocking_count: blocking.length,
    issues,
  });
}

export function buildMeetingPlatformConsumerHandoff(options = {}) {
  const platforms = selectedPlatforms(options);
  const baseUrl = firstNonEmpty(options.baseUrl, options.base_url, 'http://localhost:8787');
  const basePath = firstNonEmpty(options.basePath, options.base_path, '/api/platform-events');
  const sharedOptions = {
    ...options,
    baseUrl,
    basePath,
    platforms,
  };
  const hostPlan = buildMeetingPlatformHostIntegrationPlan(sharedOptions);
  const conformanceReport = buildMeetingPlatformConformanceReport(sharedOptions);
  const adaptationPackageMatrix = buildMeetingPlatformAdaptationPackageMatrix(sharedOptions);
  const handoffReadinessMatrix = firstNonEmpty(
    options.handoffReadinessMatrix,
    options.handoff_readiness_matrix,
  ) ?? buildMeetingPlatformHandoffReadinessMatrix({}, sharedOptions);
  const runtimeRows = byPlatform(hostPlan.runtime_bundle_matrix?.rows);
  const routeRows = byPlatform(hostPlan.adapter_route_matrix?.rows);
  const conformanceRows = byPlatform(conformanceReport.rows);
  const adaptationRows = byPlatform(adaptationPackageMatrix.rows);
  const handoffRows = byPlatform(handoffReadinessMatrix.rows);
  const trackRows = byPlatform(hostPlan.meeting_track_contract?.rows);
  const requireHandoffReady = boolOption(options, 'requireHandoffReady', 'require_handoff_ready', false);
  const requireProductionReady = boolOption(options, 'requireProductionReady', 'require_production_ready', false);
  const includeDetails = boolOption(options, 'includeDetails', 'include_details', false);
  const rows = platforms.map((platform) => platformRow(platform, {
    conformanceRows,
    runtimeRows,
    routeRows,
    adaptationRows,
    handoffRows,
    trackRows,
    requireHandoffReady,
    requireProductionReady,
  }));
  const rowIssues = rows.flatMap((row) => row.issues ?? []);
  const structuralIssues = [
    hostPlan.candidate_observation_contract?.all_ready === true
      ? undefined
      : issue('error', 'candidate_observation_contract_not_ready', 'Every selected platform must support candidate observation.', {
        missing_count: hostPlan.candidate_observation_contract?.missing_count,
      }),
    hostPlan.meeting_track_contract?.all_ready === true
      ? undefined
      : issue('error', 'meeting_track_contract_not_ready', 'Every selected platform must support realtime speaker and participant tracks.', {
        missing_count: hostPlan.meeting_track_contract?.missing_count,
      }),
    conformanceReport.accepted === true
      ? undefined
      : issue('error', 'platform_conformance_report_not_accepted', 'Static platform conformance must pass for consumer handoff.', {
        blocking_count: conformanceReport.blocking_count,
      }),
  ].filter(Boolean);
  const issues = [...structuralIssues, ...rowIssues];
  const blocking = issues.filter((item) => item.severity === 'error');
  const commands = buildConsumerCommands(sharedOptions);
  return {
    type: 'meeting_platform_consumer_handoff',
    schema: MEETING_PLATFORM_CONSUMER_HANDOFF_SCHEMA,
    schema_version: MEETING_PLATFORM_CONSUMER_HANDOFF_SCHEMA_VERSION,
    accepted: blocking.length === 0,
    base_url: baseUrl,
    base_path: basePath,
    platform_count: platforms.length,
    consumer_ready_count: rows.filter((row) => row.consumer_ready === true).length,
    conformance_accepted_count: rows.filter((row) => row.conformance_accepted === true).length,
    runtime_ready_count: rows.filter((row) => row.runtime_ready === true).length,
    adapter_route_ready_count: rows.filter((row) => row.adapter_route_ready === true).length,
    candidate_observer_count: rows.filter((row) => row.candidate_observation_ready === true).length,
    speaker_track_ready_count: rows.filter((row) => row.speaker_track_ready === true).length,
    participant_track_ready_count: rows.filter((row) => row.participant_track_ready === true).length,
    handoff_ready_count: rows.filter((row) => row.handoff_ready === true).length,
    pilot_ready_count: rows.filter((row) => row.pilot_ready === true).length,
    production_ready_count: rows.filter((row) => row.production_ready === true).length,
    blocking_count: blocking.length,
    warning_count: issues.filter((item) => item.severity === 'warning').length,
    require_handoff_ready: requireHandoffReady,
    require_production_ready: requireProductionReady,
    platforms,
    entrypoints: buildEntrypoints(hostPlan, sharedOptions),
    hard_contracts: buildContracts(hostPlan),
    boot_order: buildBootOrder(hostPlan, commands),
    endpoints: hostPlan.endpoints,
    commands,
    rows,
    issues,
    next_actions: unique([
      ...blocking.map((item) => item.code),
      ...rows.map((row) => row.first_next_action),
      ...(hostPlan.next_actions ?? []),
      ...(handoffReadinessMatrix.next_actions ?? []),
    ]),
    host_integration_plan: includeDetails ? hostPlan : hostPlanSummary(hostPlan),
    conformance_report: conformanceReport,
    adaptation_package_matrix: {
      ...adaptationPackageMatrix,
      packages: undefined,
    },
    handoff_readiness_matrix: {
      ...handoffReadinessMatrix,
      reports: undefined,
    },
  };
}

export function assertMeetingPlatformConsumerHandoff(handoffOrOptions = {}, options = {}) {
  const handoff = handoffOrOptions?.schema === MEETING_PLATFORM_CONSUMER_HANDOFF_SCHEMA
    ? handoffOrOptions
    : buildMeetingPlatformConsumerHandoff({
      ...handoffOrOptions,
      ...options,
    });
  if (handoff.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting platform consumer handoff is not ready', {
      code: 'meeting_platform_consumer_handoff_not_ready',
      issues: handoff.issues,
      next_actions: handoff.next_actions,
    });
  }
  return handoff;
}
