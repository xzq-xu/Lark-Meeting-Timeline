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
import {
  buildMeetingPlatformConnectorHub,
} from './meeting-platform-connector.mjs';

export const MEETING_PLATFORM_CONSUMER_HANDOFF_SCHEMA = 'meeting_platform_consumer_handoff';
export const MEETING_PLATFORM_CONSUMER_HANDOFF_SCHEMA_VERSION = 1;

const DEFAULT_CONSUMER_PLATFORMS = Object.freeze([
  'lark',
  'google_meet',
  'microsoft_teams',
  'zoom',
  'webex',
]);

const DEFAULT_PRIORITY_ORDER = Object.freeze([
  'google_meet',
  'zoom',
  'microsoft_teams',
  'webex',
  'lark',
  'local_detector',
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

function priorityOrder(options = {}) {
  return unique(asArray(firstNonEmpty(
    options.priorityPlatformOrder,
    options.priority_platform_order,
    options.priorityPlatforms,
    options.priority_platforms,
    DEFAULT_PRIORITY_ORDER,
  )).map((platform) => normalizeMeetingPlatform(platform)));
}

function priorityIndex(platform, order = []) {
  const index = order.indexOf(platform);
  return index >= 0 ? index : order.length + 10;
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
      action: 'choose_lightweight_connector_or_full_integration_runtime',
      lightweight_entrypoint: 'createMeetingPlatformConnectorHub',
      content_script_bridge: 'installMeetingPlatformConnectorContentScriptBridge',
      when_to_use: 'browser_extension_or_webview_only_needs_runtime_events',
      full_runtime_entrypoint: 'createMeetingPlatformIntegrationRuntime',
    },
    {
      step: 4,
      action: 'wire_local_observer_candidate_binding',
      endpoint: hostPlan.endpoints?.platform_candidate_observation,
      runtime_event_action: 'observe_platform_candidates',
      required_before: 'insert_realtime_annotation',
    },
    {
      step: 5,
      action: 'insert_realtime_annotations_by_capture_time',
      endpoint: hostPlan.endpoints?.annotations,
      runtime_event_endpoint: hostPlan.endpoints?.runtime_events,
      required_field: 'captured_at_ms',
    },
    {
      step: 6,
      action: 'emit_speaker_and_participant_positions',
      endpoints: [
        hostPlan.endpoints?.runtime_events,
      ].filter(Boolean),
      source_priority: ['local_observer_samples', 'provider_reconcile', 'post_meeting_artifacts'],
    },
    {
      step: 7,
      action: 'connect_provider_events_as_reconcile_backfill',
      endpoint: hostPlan.endpoints?.platform_events,
      blocks_realtime_annotation: false,
    },
    {
      step: 8,
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
      implementation_handoff: '@ai-annotation/meeting-timeline-sdk/adapters/platform-implementation-handoff',
      meeting_platform_connector: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-connector',
      conformance: hostPlan.sdk?.platform_conformance_module,
      runtime_event: hostPlan.sdk?.runtime_event_module,
      live_adapter: hostPlan.sdk?.live_adapter_module,
    },
    kit_methods: [
      'platformConsumerHandoff',
      'assertPlatformConsumerHandoff',
      'platformImplementationHandoff',
      'platformImplementationHandoffMatrix',
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
      'platformConnector',
      'platformConnectorMatrix',
      'platformConnectorHub',
      'resolvePlatformConnector',
      'createPlatformConnectorRuntime',
      'createPlatformConnectorHub',
      'createPlatformConnectorBrowserRuntime',
      'createPlatformConnectorContentScriptBridge',
      'installPlatformConnectorContentScriptBridge',
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

function buildLightweightConnectorHandoff(connectorHub = {}, hostPlan = {}) {
  return compactObject({
    type: 'meeting_platform_lightweight_connector_handoff',
    accepted: connectorHub.accepted === true,
    platform_count: connectorHub.platform_count,
    accepted_count: connectorHub.accepted_count,
    realtime_ready_count: connectorHub.realtime_ready_count,
    candidate_observer_count: connectorHub.candidate_observer_count,
    platforms: connectorHub.platforms,
    runtime_event_endpoint: connectorHub.runtime_event_endpoint,
    module: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-connector',
    factories: {
      build_connector: 'buildMeetingPlatformConnector',
      create_runtime: 'createMeetingPlatformConnectorRuntime',
      build_hub: 'buildMeetingPlatformConnectorHub',
      create_hub: 'createMeetingPlatformConnectorHub',
      create_browser_runtime: 'createMeetingPlatformConnectorBrowserRuntime',
      create_content_script_bridge: 'createMeetingPlatformConnectorContentScriptBridge',
      install_content_script_bridge: 'installMeetingPlatformConnectorContentScriptBridge',
    },
    kit_methods: [
      'platformConnector',
      'platformConnectorMatrix',
      'platformConnectorHub',
      'resolvePlatformConnector',
      'createPlatformConnectorRuntime',
      'createPlatformConnectorHub',
      'createPlatformConnectorBrowserRuntime',
      'createPlatformConnectorContentScriptBridge',
      'installPlatformConnectorContentScriptBridge',
    ],
    content_script_bridge: {
      supported_surfaces: ['browser_extension_content_script', 'electron_webview_preload', 'mobile_webview'],
      route_source_priority: [
        'explicit platform/provider',
        'current location.href',
        'message payload url',
        'tab/window candidate url',
      ],
      message_types: [
        'meeting_timeline.insert_mark',
        'meeting_timeline.insert_marks',
        'meeting_timeline.sample',
        'meeting_timeline.sample_tracks',
        'meeting_timeline.preflight_current_window',
        'meeting_timeline.provider_event',
        'meeting_timeline.observe_candidates',
      ],
      output_runtime_actions: [
        'insert_annotation',
        'observe_meeting_app',
        'observe_platform_candidates',
        'provider_event',
        'speaker_track',
        'participant_track',
      ],
    },
    host_requirements: {
      fetch_required_in_browser_context: true,
      runtime_event_endpoint: hostPlan.endpoints?.runtime_events ?? connectorHub.runtime_event_endpoint,
      annotation_endpoint: hostPlan.endpoints?.annotations,
      candidate_observation_endpoint: hostPlan.endpoints?.platform_candidate_observation,
      timestamp_field: 'captured_at_ms',
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
    },
    readiness: connectorHub.readiness,
    issues: connectorHub.issues,
  });
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
    lightweight_connector_bridge_supported: true,
    connector_bridge_message_types: [
      'meeting_timeline.insert_mark',
      'meeting_timeline.sample',
      'meeting_timeline.sample_tracks',
      'meeting_timeline.preflight_current_window',
      'meeting_timeline.provider_event',
    ],
    host_runtime_contract: hostPlan.runtime_contract,
  };
}

function buildSdkFacadeHandoff(platforms = [], inputs = {}) {
  const {
    baseUrl,
    hostPlan = {},
    adaptationRows = {},
    runtimeRows = {},
    routeRows = {},
    connectorHub = {},
  } = inputs;
  const platformRows = platforms.map((platform) => {
    const adaptation = adaptationRows[platform] ?? {};
    const runtime = runtimeRows[platform] ?? {};
    const route = routeRows[platform] ?? {};
    return compactObject({
      platform,
      display_name: adaptation.display_name ?? runtime.display_name,
      next_phase: adaptation.next_phase,
      provider_path: adaptation.provider_path,
      provider_permission_risk: adaptation.provider_permission_risk,
      adapter_first_route: route.first_route ?? runtime.adapter_first_route,
      runtime_bundle_ready: runtime.runtime_ready,
      lightweight_connector_ready: connectorHub.platforms?.includes(platform) && connectorHub.accepted === true,
      facade_methods: {
        adaptation_package: `sdk.platformAdaptationPackage('${platform}')`,
        runtime_bundle: `sdk.platformRuntimeBundle('${platform}')`,
        adapter_route: `sdk.platformAdapterRoute('${platform}')`,
        adaptation_strategy: `sdk.platformAdaptationStrategy('${platform}')`,
      },
    });
  });
  return compactObject({
    type: 'meeting_platform_sdk_facade_handoff',
    package: '@ai-annotation/meeting-timeline-sdk',
    create_function: 'createMeetingAppTimelineSdk',
    constructor_options: compactObject({
      baseUrl,
      platforms,
    }),
    required_facade_methods: [
      'detect',
      'observePlatformCandidates',
      'observeMeetingApp',
      'insertAnnotation',
      'speakerTrack',
      'participantTrack',
      'ingestProvider',
      'platformAdaptationPackage',
      'platformConsumerHandoff',
      'platformImplementationHandoff',
      'platformImplementationHandoffMatrix',
      'platformRuntimeBundle',
      'platformAdapterRoute',
      'platformAdaptationStrategy',
      'platformConnectorHub',
    ],
    minimal_realtime_flow: [
      {
        step: 1,
        method: 'createMeetingAppTimelineSdk({ baseUrl, platforms })',
        output: 'meeting_app_timeline_sdk',
      },
      {
        step: 2,
        method: 'sdk.platformConsumerHandoff()',
        output: 'meeting_platform_consumer_handoff',
        purpose: 'validate_static_contracts_before_runtime_install',
      },
      {
        step: 3,
        method: 'sdk.observePlatformCandidates({ tabs/windows })',
        output: 'local_observer_axis_or_candidate_binding',
        purpose: 'create_or_bind_the_current_meeting_axis_before_marks',
      },
      {
        step: 4,
        method: 'sdk.insertAnnotation(platform, { captured_at_ms, ...mark }, { remote: true })',
        output: 'timeline_annotation',
        purpose: 'insert_realtime_marks_by_device_capture_time',
      },
      {
        step: 5,
        method: 'sdk.speakerTrack(platform, sample) / sdk.participantTrack(platform, sample)',
        output: 'speaker_or_participant_position_markers',
        purpose: 'draw_positions_without_realtime_transcript_text',
      },
      {
        step: 6,
        method: 'sdk.ingestProvider(platform, providerEvent)',
        output: 'reconcile_or_backfill_signal',
        purpose: 'official_events_reconcile_lifecycle_without_blocking_marks',
      },
    ],
    surface_wiring: {
      browser_extension: {
        install_method: 'sdk.platformRuntimeBundle(platform).runtime.lightweight_connector_bridge',
        manifest_source: 'sdk.platformRuntimeBundle(platform).extension.content_script_manifest',
        first_background_call: 'sdk.observePlatformCandidates({ tabs })',
        content_script_messages: [
          'meeting_timeline.sample',
          'meeting_timeline.sample_tracks',
          'meeting_timeline.preflight_current_window',
          'meeting_timeline.preflight_candidates',
          'meeting_timeline.insert_mark',
          'meeting_timeline.provider_event',
          'meeting_timeline.observe_candidates',
        ],
      },
      electron_webview: {
        install_method: 'sdk.platformRuntimeBundle(platform).runtime.lightweight_connector_bridge.options with windowMessaging=true',
        first_preload_call: 'sdk.observeMeetingApp(platform, webviewSnapshot)',
        host_bridge: 'postMessage meeting_timeline.* envelope to preload bridge',
      },
      native_detector: {
        install_method: 'host calls sdk.observePlatformCandidates() and sdk.insertAnnotation() directly',
        first_detector_call: 'sdk.observePlatformCandidates({ windows, processes, activeWindow })',
        timestamp_source: 'device_absolute_captured_at_ms',
      },
      provider_adapter: {
        install_method: 'host webhook/long-connection handler calls sdk.ingestProvider(platform, payload)',
        realtime_blocking: false,
        role: 'reconcile_and_backfill_after_local_axis',
      },
    },
    host_endpoints: hostPlan.endpoints,
    runtime_event_endpoint: hostPlan.endpoints?.runtime_events,
    timestamp_field: 'captured_at_ms',
    provider_events_block_realtime: false,
    transcript_blocks_realtime: false,
    platform_rows: platformRows,
  });
}

function buildSurfaceCoverageMatrix(platforms = [], inputs = {}) {
  const {
    adaptationRows = {},
    runtimeRows = {},
    routeRows = {},
    handoffRows = {},
  } = inputs;
  const rows = platforms.map((platform) => {
    const adaptation = adaptationRows[platform] ?? {};
    const runtime = runtimeRows[platform] ?? {};
    const route = routeRows[platform] ?? {};
    const handoff = handoffRows[platform] ?? {};
    const browserReady = (runtime.browser_match_count ?? adaptation.browser_match_count ?? 0) > 0
      && (runtime.candidate_observation_ready === true || adaptation.candidate_observation_ready === true);
    const webviewReady = browserReady && runtime.lightweight_connector_ready === true;
    const nativeDetectorReady = (route.route_ready === true || route.adapter_route_ready === true || handoff.pilot_ready === true)
      && runtime.provider_required_for_realtime !== true
      && runtime.transcript_blocks_realtime !== true;
    const providerReady = Boolean(adaptation.provider_transport ?? runtime.provider_transport)
      && runtime.provider_required_for_realtime !== true;
    const backfillSupported = adaptation.post_meeting_backfill_supported === true
      || runtime.transcript_blocks_realtime === false;
    return {
      platform,
      display_name: adaptation.display_name ?? runtime.display_name ?? handoff.display_name,
      browser_extension: {
        ready: browserReady,
        browser_match_count: runtime.browser_match_count ?? adaptation.browser_match_count ?? 0,
        candidate_observation_ready: runtime.candidate_observation_ready === true || adaptation.candidate_observation_ready === true,
        message_type: runtime.candidate_observer_message_type ?? adaptation.candidate_observer_message_type,
        permission: runtime.candidate_observer_permission ?? adaptation.candidate_observer_permission,
        install_method: 'sdk.platformRuntimeBundle(platform).runtime.lightweight_connector_bridge',
      },
      webview_preload: {
        ready: webviewReady,
        window_messaging_supported: true,
        install_method: 'sdk.platformRuntimeBundle(platform).runtime.lightweight_connector_bridge.options with windowMessaging=true',
      },
      native_detector: {
        ready: nativeDetectorReady,
        first_call: 'sdk.observePlatformCandidates({ windows, processes, activeWindow })',
        adapter_first_route: route.first_route ?? runtime.adapter_first_route,
        timestamp_field: 'captured_at_ms',
      },
      provider_reconcile: {
        ready: providerReady,
        transport: adaptation.provider_transport ?? runtime.provider_transport,
        provider_path: adaptation.provider_path,
        realtime_blocking: false,
        required_for_realtime: runtime.provider_required_for_realtime === true,
        permission_risk: adaptation.provider_permission_risk,
      },
      post_meeting_backfill: {
        supported: backfillSupported,
        transcript_blocks_realtime: runtime.transcript_blocks_realtime === true,
        source: 'transcript_or_recording_artifact_after_meeting_end',
      },
      lightweight_connector: {
        ready: runtime.lightweight_connector_ready === true,
        install_function: runtime.lightweight_connector_install_function,
      },
      speaker_track: {
        ready: handoff.speaker_track_ready === true || runtime.observer_scheduler_ready === true,
        min_stable_ms: handoff.speaker_min_stable_ms ?? runtime.speaker_min_stable_ms,
      },
      participant_track: {
        ready: handoff.participant_track_ready === true || runtime.observer_scheduler_ready === true,
        duplicate_window_ms: handoff.participant_duplicate_window_ms,
      },
    };
  });
  return {
    type: 'meeting_platform_surface_coverage_matrix',
    schema: 'meeting_platform_surface_coverage_matrix',
    schema_version: 1,
    platform_count: rows.length,
    browser_extension_ready_count: rows.filter((row) => row.browser_extension.ready).length,
    webview_preload_ready_count: rows.filter((row) => row.webview_preload.ready).length,
    native_detector_ready_count: rows.filter((row) => row.native_detector.ready).length,
    provider_reconcile_ready_count: rows.filter((row) => row.provider_reconcile.ready).length,
    post_meeting_backfill_supported_count: rows.filter((row) => row.post_meeting_backfill.supported).length,
    lightweight_connector_ready_count: rows.filter((row) => row.lightweight_connector.ready).length,
    speaker_track_ready_count: rows.filter((row) => row.speaker_track.ready).length,
    participant_track_ready_count: rows.filter((row) => row.participant_track.ready).length,
    platforms,
    rows,
  };
}

function recommendedSurface(coverageRow = {}) {
  if (coverageRow.browser_extension?.ready === true) return 'browser_extension';
  if (coverageRow.webview_preload?.ready === true) return 'webview_preload';
  if (coverageRow.native_detector?.ready === true) return 'native_detector';
  if (coverageRow.provider_reconcile?.ready === true) return 'provider_reconcile';
  return 'manual_or_local_detector';
}

function surfaceReady(surface, coverageRow = {}) {
  if (surface === 'manual_or_local_detector') return false;
  return coverageRow[surface]?.ready === true;
}

function priorityTierRank(tier) {
  if (tier === 'production') return 0;
  if (tier === 'pilot') return 1;
  return 2;
}

function roadmapReasons(row = {}, coverageRow = {}, adaptation = {}) {
  return [
    coverageRow.browser_extension?.ready === true ? 'browser_extension_ready_for_local_axis' : undefined,
    coverageRow.lightweight_connector?.ready === true ? 'lightweight_connector_ready' : undefined,
    coverageRow.provider_reconcile?.ready === true ? `provider_reconcile_path:${coverageRow.provider_reconcile.provider_path}` : undefined,
    coverageRow.post_meeting_backfill?.supported === true ? 'post_meeting_backfill_supported' : undefined,
    row.provider_blocks_realtime === false ? 'provider_does_not_block_realtime_marks' : undefined,
    row.transcript_blocks_realtime === false ? 'transcript_does_not_block_realtime_marks' : undefined,
    adaptation.speaker_realtime_gap === true ? 'speaker_track_uses_local_observer_or_backfill' : undefined,
  ].filter(Boolean);
}

function productionGaps(row = {}, adaptation = {}, coverageRow = {}) {
  const gaps = [];
  if (row.production_ready !== true) gaps.push('production_evidence_pending');
  if (row.handoff_ready !== true) gaps.push('handoff_readiness_evidence_pending');
  if (adaptation.ready_for_realtime_annotations !== true) gaps.push(adaptation.next_phase ?? 'axis_bootstrap_pending');
  if (coverageRow.provider_reconcile?.ready !== true) gaps.push('provider_reconcile_not_ready');
  return unique(gaps);
}

function buildAdaptationRoadmap(platforms = [], inputs = {}) {
  const {
    rowsByPlatform = {},
    adaptationRows = {},
    surfaceCoverageMatrix = {},
    options = {},
  } = inputs;
  const order = priorityOrder(options);
  const coverageRows = byPlatform(surfaceCoverageMatrix.rows);
  const roadmapRows = platforms.map((platform) => {
    const row = rowsByPlatform[platform] ?? {};
    const adaptation = adaptationRows[platform] ?? {};
    const coverage = coverageRows[platform] ?? {};
    const surface = recommendedSurface(coverage);
    const gaps = productionGaps(row, adaptation, coverage);
    const productionReady = row.production_ready === true;
    const pilotReady = row.consumer_ready === true
      && surfaceReady(surface, coverage)
      && row.provider_blocks_realtime === false
      && row.transcript_blocks_realtime === false;
    return compactObject({
      platform,
      display_name: row.display_name ?? adaptation.display_name,
      rank_hint: priorityIndex(platform, order) + 1,
      priority_tier: productionReady ? 'production' : pilotReady ? 'pilot' : 'blocked',
      recommended_first_surface: surface,
      next_phase: adaptation.next_phase ?? row.first_next_action,
      next_action: row.first_next_action ?? adaptation.next_phase ?? gaps[0],
      provider_path: adaptation.provider_path ?? coverage.provider_reconcile?.provider_path,
      provider_permission_risk: adaptation.provider_permission_risk ?? coverage.provider_reconcile?.permission_risk,
      pilot_ready: pilotReady,
      production_ready: productionReady,
      production_gaps: gaps,
      reasons: roadmapReasons(row, coverage, adaptation),
    });
  }).sort((left, right) => (
    priorityTierRank(left.priority_tier) - priorityTierRank(right.priority_tier)
    || Number(left.rank_hint ?? 999) - Number(right.rank_hint ?? 999)
  ));
  return {
    type: 'meeting_platform_adaptation_roadmap',
    schema: 'meeting_platform_adaptation_roadmap',
    schema_version: 1,
    platform_count: roadmapRows.length,
    pilot_ready_count: roadmapRows.filter((row) => row.pilot_ready).length,
    production_ready_count: roadmapRows.filter((row) => row.production_ready).length,
    recommended_first_platform: roadmapRows[0]?.platform,
    recommended_first_surface: roadmapRows[0]?.recommended_first_surface,
    priority_order: order,
    rows: roadmapRows,
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
  const connectorHub = buildMeetingPlatformConnectorHub(sharedOptions);
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
  const surfaceCoverageMatrix = buildSurfaceCoverageMatrix(platforms, {
    adaptationRows,
    runtimeRows,
    routeRows,
    handoffRows,
  });
  const rowsByPlatform = byPlatform(rows);
  const adaptationRoadmap = buildAdaptationRoadmap(platforms, {
    rowsByPlatform,
    adaptationRows,
    surfaceCoverageMatrix,
    options: sharedOptions,
  });
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
    connectorHub.accepted === true
      ? undefined
      : issue('error', 'lightweight_connector_handoff_not_ready', 'Lightweight connector hub must be ready for browser/WebView consumer handoff.', {
        blocking_count: connectorHub.blocking_count,
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
    lightweight_connector_ready: connectorHub.accepted === true,
    lightweight_connector_platform_count: connectorHub.platform_count,
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
    lightweight_connector_handoff: buildLightweightConnectorHandoff(connectorHub, hostPlan),
    sdk_facade_handoff: buildSdkFacadeHandoff(platforms, {
      baseUrl,
      hostPlan,
      adaptationRows,
      runtimeRows,
      routeRows,
      connectorHub,
    }),
    surface_coverage_matrix: surfaceCoverageMatrix,
    adaptation_roadmap: adaptationRoadmap,
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
