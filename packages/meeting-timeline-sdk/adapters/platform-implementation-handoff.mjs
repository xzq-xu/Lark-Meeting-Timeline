import { compactObject } from '../index.mjs';
import {
  buildMeetingPlatformConsumerHandoff,
} from './platform-consumer-handoff.mjs';
import {
  buildMeetingPlatformRuntimeBundle,
  buildMeetingPlatformRuntimeBundleMatrix,
} from './platform-runtime-bundle.mjs';
import {
  buildMeetingPlatformAdaptationStrategy,
} from './platform-strategy.mjs';
import {
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_IMPLEMENTATION_HANDOFF_SCHEMA = 'meeting_platform_implementation_handoff';
export const MEETING_PLATFORM_IMPLEMENTATION_HANDOFF_MATRIX_SCHEMA = 'meeting_platform_implementation_handoff_matrix';
export const MEETING_PLATFORM_IMPLEMENTATION_HANDOFF_SCHEMA_VERSION = 1;

const DEFAULT_IMPLEMENTATION_PLATFORMS = Object.freeze([
  'google_meet',
  'zoom',
  'microsoft_teams',
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
    DEFAULT_IMPLEMENTATION_PLATFORMS,
  )).map((platform) => normalizeMeetingPlatform(platform)));
}

function byPlatform(rows = []) {
  return Object.fromEntries(asArray(rows).map((row) => [row.platform, row]));
}

function priorityOrder(options = {}) {
  return unique(asArray(firstNonEmpty(
    options.priorityPlatformOrder,
    options.priority_platform_order,
    DEFAULT_IMPLEMENTATION_PLATFORMS,
  )).map((platform) => normalizeMeetingPlatform(platform)));
}

function commandWithBase(name, options = {}, extra = '') {
  const baseUrl = firstNonEmpty(options.baseUrl, options.base_url);
  const args = [
    baseUrl ? `--base-url=${baseUrl}` : '',
    String(extra).trim(),
  ].filter(Boolean).join(' ');
  return args ? `npm run ${name} -- ${args}` : `npm run ${name}`;
}

function flowSteps(platform, surface, bundle = {}) {
  return [
    {
      step: 1,
      id: 'install_sdk',
      action: 'add_package_and_create_sdk_facade',
      method: "createMeetingAppTimelineSdk({ baseUrl, platforms: [platform] })",
      output: 'meeting_app_timeline_sdk',
    },
    {
      step: 2,
      id: 'install_surface_runtime',
      action: surface === 'browser_extension'
        ? 'install_browser_content_script_or_extension_runtime'
        : surface === 'webview_preload'
          ? 'install_webview_preload_bridge'
          : 'install_native_or_host_detector_bridge',
      install_function: bundle.runtime?.lightweight_connector_bridge?.install_function,
      output: 'meeting_timeline.* messages_delivered_to_host_runtime_endpoint',
    },
    {
      step: 3,
      id: 'bind_current_axis',
      action: 'observe_platform_candidates_before_the_first_mark',
      message_type: bundle.messaging?.candidate_observation?.message_type,
      sdk_method: 'sdk.observePlatformCandidates({ tabs/windows })',
      required_before: 'insert_annotation',
    },
    {
      step: 4,
      id: 'insert_realtime_annotations',
      action: 'write_marks_with_device_capture_time',
      sdk_method: "sdk.insertAnnotation(platform, { captured_at_ms, ...mark })",
      endpoint: bundle.host?.endpoints?.insertMark,
      required_field: 'captured_at_ms',
    },
    {
      step: 5,
      id: 'emit_speaker_positions',
      action: 'optionally_emit_speaker_or_participant_track_markers_without_transcript_text',
      sdk_method: 'sdk.speakerTrack(platform, sample) / sdk.participantTrack(platform, sample)',
      output: 'speaker_or_participant_position_markers',
    },
    {
      step: 6,
      id: 'provider_reconcile',
      action: 'ingest_provider_events_only_as_reconcile_or_backfill',
      sdk_method: 'sdk.ingestProvider(platform, providerEvent)',
      realtime_blocking: false,
    },
    {
      step: 7,
      id: 'verify_handoff',
      action: 'run_static_handoff_and_runtime_bundle_checks_before_real_evidence',
      commands: [
        commandWithBase('meeting-platform:consumer-handoff', { platforms: [platform] }, `--platforms=${platform}`),
        commandWithBase('meeting-platform:runtime-bundle', { platforms: [platform] }, `--platforms=${platform}`),
      ],
    },
  ];
}

function acceptanceCommands(platform, options = {}) {
  return {
    implementation_handoff: commandWithBase('meeting-platform:implementation-handoff', options, `--platforms=${platform}`),
    consumer_handoff: commandWithBase('meeting-platform:consumer-handoff', options, `--platforms=${platform}`),
    runtime_bundle: commandWithBase('meeting-platform:runtime-bundle', options, `--platforms=${platform}`),
    adapter_route: commandWithBase('meeting-platform:adapter-route', options, `--platforms=${platform}`),
    handoff_readiness: commandWithBase('meeting-platform:handoff-readiness', options, `--platforms=${platform}`),
    runtime_host_verify: `npm run meeting-platform:runtime-host-verify -- --platforms=${platform}`,
  };
}

function recommendedSurface(row = {}, bundle = {}) {
  return firstNonEmpty(
    row.recommended_first_surface,
    bundle.readiness?.lightweight_connector_ready === true ? 'browser_extension' : undefined,
    bundle.readiness?.runtime_ready === true ? 'native_detector' : undefined,
    'manual_or_local_detector',
  );
}

function runtimeEventActions(bundle = {}) {
  return unique((bundle.messaging?.runtime_event?.plan?.actions ?? []).map((item) => item.action));
}

function buildSingleHandoff(platform, inputs = {}) {
  const {
    options = {},
    roadmapRow = {},
    consumerRow = {},
    bundle,
  } = inputs;
  const key = normalizeMeetingPlatform(platform);
  const runtimeBundle = bundle ?? buildMeetingPlatformRuntimeBundle(key, options);
  const strategy = buildMeetingPlatformAdaptationStrategy(key, options);
  const surface = recommendedSurface(roadmapRow, runtimeBundle);
  const provider = runtimeBundle.provider_reconcile ?? {};
  const productionGaps = unique([
    ...(roadmapRow.production_gaps ?? []),
    ...(runtimeBundle.readiness?.missing_items ?? []),
  ]);
  const runtimeReady = runtimeBundle.readiness?.runtime_ready === true
    && runtimeBundle.readiness?.provider_required_for_realtime !== true
    && runtimeBundle.readiness?.transcript_blocks_realtime !== true;
  const bridgeReady = runtimeBundle.readiness?.lightweight_connector_ready === true
    || surface === 'native_detector';
  const implementationReady = consumerRow.consumer_ready === true && runtimeReady && bridgeReady;

  return compactObject({
    type: 'meeting_platform_implementation_handoff',
    schema: MEETING_PLATFORM_IMPLEMENTATION_HANDOFF_SCHEMA,
    schema_version: MEETING_PLATFORM_IMPLEMENTATION_HANDOFF_SCHEMA_VERSION,
    platform: key,
    display_name: runtimeBundle.display_name ?? roadmapRow.display_name,
    objective: 'external_project_can_install_a_realtime_meeting_annotation_timeline_adapter',
    priority_tier: roadmapRow.priority_tier,
    rank_hint: roadmapRow.rank_hint,
    implementation_ready: implementationReady,
    pilot_ready: roadmapRow.pilot_ready === true,
    production_ready: roadmapRow.production_ready === true,
    recommended_first_surface: surface,
    next_phase: roadmapRow.next_phase,
    next_action: roadmapRow.next_action,
    package_entrypoints: {
      package: '@ai-annotation/meeting-timeline-sdk',
      root_create_function: 'createMeetingAppTimelineSdk',
      kit_create_function: 'createMeetingPlatformTimelineKit',
      runtime_bundle_module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-bundle',
      implementation_handoff_module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-implementation-handoff',
    },
    install_surface: {
      surface,
      browser_matches: runtimeBundle.browser?.matches ?? [],
      browser_permissions: runtimeBundle.browser?.permissions ?? [],
      host_permissions: runtimeBundle.browser?.host_permissions ?? [],
      content_script_manifest: runtimeBundle.browser?.manifest,
      runtime_preset: runtimeBundle.runtime?.preset,
      start_options: runtimeBundle.runtime?.start_options,
      lightweight_connector_bridge: runtimeBundle.runtime?.lightweight_connector_bridge,
      observation_loop: runtimeBundle.runtime?.observation_loop,
      observer_scheduler: runtimeBundle.runtime?.observer_scheduler,
      runtime_host: runtimeBundle.runtime?.runtime_host,
    },
    runtime_events: {
      endpoint: runtimeBundle.messaging?.runtime_event?.endpoint,
      schema: runtimeBundle.messaging?.runtime_event?.schema,
      accepted_methods: runtimeBundle.messaging?.accepted_methods ?? [],
      message_types: runtimeBundle.messaging?.lightweight_connector_message_types ?? [],
      actions: runtimeEventActions(runtimeBundle),
      candidate_observation: runtimeBundle.messaging?.candidate_observation,
      examples: runtimeBundle.messaging?.examples,
    },
    provider_reconcile: {
      path: roadmapRow.provider_path ?? provider.provider_path,
      transport: provider.transport,
      permission_risk: roadmapRow.provider_permission_risk ?? provider.permission_risk,
      required_for_realtime: provider.required_for_realtime === true,
      realtime_blocking: false,
    },
    contracts: {
      timestamp_field: 'captured_at_ms',
      local_axis_first: true,
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      per_meeting_annotation_isolation_required: true,
      meeting_end_must_close_current_axis: true,
      speaker_track_text_required: false,
      participant_track_text_required: false,
    },
    implementation_flow: flowSteps(key, surface, runtimeBundle),
    acceptance: {
      static_ready_condition: 'implementation_ready === true',
      pilot_condition: 'pilot_ready === true && collect_real_meeting_axis_evidence',
      production_condition: 'production_ready === true after handoff readiness evidence',
      commands: acceptanceCommands(key, options),
    },
    readiness: {
      consumer_ready: consumerRow.consumer_ready === true,
      runtime_ready: runtimeBundle.readiness?.runtime_ready === true,
      observer_plan_ready: runtimeBundle.readiness?.observer_plan_ready === true,
      observer_scheduler_ready: runtimeBundle.readiness?.observer_scheduler_ready === true,
      runtime_host_ready: runtimeBundle.readiness?.runtime_host_ready === true,
      lightweight_connector_ready: runtimeBundle.readiness?.lightweight_connector_ready === true,
      provider_required_for_realtime: runtimeBundle.readiness?.provider_required_for_realtime === true,
      transcript_blocks_realtime: runtimeBundle.readiness?.transcript_blocks_realtime === true,
    },
    adaptation_strategy: {
      recommended_mode: strategy.recommended_mode,
      current_phase: strategy.current_phase,
      phases: strategy.phases,
    },
    production_gaps: productionGaps,
    next_actions: unique([
      roadmapRow.next_action,
      ...productionGaps,
      ...(runtimeBundle.next_actions ?? []),
    ]),
  });
}

export function buildMeetingPlatformImplementationHandoff(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const sharedOptions = {
    ...options,
    platforms: [key],
  };
  const consumerHandoff = buildMeetingPlatformConsumerHandoff(sharedOptions);
  const roadmapRow = consumerHandoff.adaptation_roadmap?.rows?.[0] ?? {};
  const consumerRow = consumerHandoff.rows?.[0] ?? {};
  return buildSingleHandoff(key, {
    options: sharedOptions,
    roadmapRow,
    consumerRow,
  });
}

export function buildMeetingPlatformImplementationHandoffMatrix(options = {}) {
  const platforms = selectedPlatforms(options);
  const order = priorityOrder(options);
  const sharedOptions = {
    ...options,
    platforms,
  };
  const consumerHandoff = buildMeetingPlatformConsumerHandoff(sharedOptions);
  const runtimeBundleMatrix = buildMeetingPlatformRuntimeBundleMatrix(sharedOptions);
  const roadmapRows = byPlatform(consumerHandoff.adaptation_roadmap?.rows);
  const consumerRows = byPlatform(consumerHandoff.rows);
  const bundles = byPlatform(runtimeBundleMatrix.bundles);
  const handoffs = platforms.map((platform) => buildSingleHandoff(platform, {
    options: {
      ...sharedOptions,
      platforms: [platform],
    },
    roadmapRow: roadmapRows[platform],
    consumerRow: consumerRows[platform],
    bundle: bundles[platform],
  })).sort((left, right) => Number(left.rank_hint ?? 999) - Number(right.rank_hint ?? 999));

  return {
    type: 'meeting_platform_implementation_handoff_matrix',
    schema: MEETING_PLATFORM_IMPLEMENTATION_HANDOFF_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_IMPLEMENTATION_HANDOFF_SCHEMA_VERSION,
    platform_count: handoffs.length,
    implementation_ready_count: handoffs.filter((handoff) => handoff.implementation_ready).length,
    pilot_ready_count: handoffs.filter((handoff) => handoff.pilot_ready).length,
    production_ready_count: handoffs.filter((handoff) => handoff.production_ready).length,
    recommended_first_platform: handoffs[0]?.platform,
    recommended_first_surface: handoffs[0]?.recommended_first_surface,
    priority_order: order,
    platforms: handoffs.map((handoff) => handoff.platform),
    rows: handoffs.map((handoff) => ({
      platform: handoff.platform,
      display_name: handoff.display_name,
      rank_hint: handoff.rank_hint,
      priority_tier: handoff.priority_tier,
      recommended_first_surface: handoff.recommended_first_surface,
      implementation_ready: handoff.implementation_ready,
      pilot_ready: handoff.pilot_ready,
      production_ready: handoff.production_ready,
      provider_path: handoff.provider_reconcile?.path,
      provider_permission_risk: handoff.provider_reconcile?.permission_risk,
      runtime_event_endpoint: handoff.runtime_events?.endpoint,
      browser_match_count: handoff.install_surface?.browser_matches?.length ?? 0,
      production_gap_count: handoff.production_gaps?.length ?? 0,
      next_action: handoff.next_action,
    })),
    handoffs,
    next_actions: unique(handoffs.flatMap((handoff) => handoff.next_actions ?? [])),
  };
}
