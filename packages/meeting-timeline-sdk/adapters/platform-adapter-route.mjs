import { compactObject } from './internal-utils.mjs';
import {
  buildPlatformIntegrationPlan,
  normalizeMeetingPlatform,
  platformCapabilityContract,
} from './platform-setup.mjs';
import {
  buildMeetingPlatformRuntimeProfile,
} from './platform-runtime-profile.mjs';
import {
  buildMeetingPlatformProviderConnectionPack,
} from './platform-provider-connection.mjs';
import {
  buildMeetingPlatformRuntimeEventPlan,
} from './platform-runtime-event.mjs';
import {
  buildMeetingAppRuntimeAdapterConfig,
  buildMeetingAppRuntimeObserverPlan,
} from './meeting-app-profile.mjs';

export const MEETING_PLATFORM_ADAPTER_ROUTE_SCHEMA = 'meeting_platform_adapter_route';
export const MEETING_PLATFORM_ADAPTER_ROUTE_MATRIX_SCHEMA = 'meeting_platform_adapter_route_matrix';
export const MEETING_PLATFORM_ADAPTER_ROUTE_SCHEMA_VERSION = 1;

const DEFAULT_ROUTE_PLATFORMS = Object.freeze([
  'lark',
  'google_meet',
  'microsoft_teams',
  'zoom',
  'webex',
]);

const NATIVE_ENTRYPOINTS = Object.freeze({
  lark: Object.freeze({
    app_names: ['Lark', 'Feishu'],
    process_hints: ['Lark', 'Feishu'],
    window_title_hints: ['Lark', 'Feishu', 'Meeting', 'Video Meeting'],
  }),
  google_meet: Object.freeze({
    app_names: ['Google Chrome', 'Chrome PWA', 'Google Meet PWA'],
    process_hints: ['Google Chrome', 'Chrome', 'Google Meet'],
    window_title_hints: ['Google Meet', 'Meet'],
  }),
  microsoft_teams: Object.freeze({
    app_names: ['Microsoft Teams', 'Teams'],
    process_hints: ['Microsoft Teams', 'Teams', 'ms-teams'],
    window_title_hints: ['Microsoft Teams', 'Teams Meeting', 'Meeting'],
  }),
  zoom: Object.freeze({
    app_names: ['Zoom Workplace', 'zoom.us', 'Zoom'],
    process_hints: ['zoom.us', 'Zoom Workplace', 'Zoom'],
    window_title_hints: ['Zoom Meeting', 'Zoom Workplace', 'Meeting'],
  }),
  webex: Object.freeze({
    app_names: ['Cisco Webex', 'Webex'],
    process_hints: ['Cisco Webex', 'Webex'],
    window_title_hints: ['Cisco Webex', 'Webex Meeting', 'Meeting'],
  }),
});

const PROVIDER_REFERENCES = Object.freeze({
  lark: Object.freeze({
    vendor: 'Feishu/Lark Open Platform',
    realtime_events: 'vc.meeting started/ended/join/leave events',
    docs: [],
  }),
  google_meet: Object.freeze({
    vendor: 'Google Workspace Events API',
    realtime_events: 'conference, participant, recording, transcript, and smart note events',
    docs: [
      'https://developers.google.com/workspace/events/guides/events-meet',
      'https://developers.google.com/workspace/events/guides/auth',
    ],
  }),
  microsoft_teams: Object.freeze({
    vendor: 'Microsoft Graph change notifications',
    realtime_events: 'meetingCallEvents call started, call ended, and roster updated notifications',
    docs: [
      'https://learn.microsoft.com/en-us/graph/changenotifications-for-onlinemeeting',
      'https://learn.microsoft.com/en-us/graph/teams-changenotifications-callrecording-and-calltranscript',
    ],
  }),
  zoom: Object.freeze({
    vendor: 'Zoom Meeting webhooks',
    realtime_events: 'meeting.started, meeting.ended, participant, and recording events',
    docs: [
      'https://developers.zoom.us/docs/api/meetings/events/',
      'https://developers.zoom.us/docs/api/webhooks/',
    ],
  }),
  webex: Object.freeze({
    vendor: 'Webex webhooks',
    realtime_events: 'meetings started/ended, meetingParticipants joined/left, recordings, and transcripts',
    docs: [
      'https://developer.webex.com/messaging/docs/api/guides/webhooks',
      'https://developer.webex.com/docs/api/v1/webhooks/create-a-webhook',
    ],
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

function routeHasRequiredRealtimeStep(adapterRoute = {}, routeName) {
  return asArray(adapterRoute.routes).some((route) => route.route === routeName && route.blocks_realtime_if_missing === true);
}

function issue(code, message, details = {}) {
  return compactObject({
    code,
    message,
    ...details,
  });
}

function selectedPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, DEFAULT_ROUTE_PLATFORMS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

function safeRuntimeAdapter(platform, options = {}) {
  try {
    return buildMeetingAppRuntimeAdapterConfig(platform, options);
  } catch {
    return {};
  }
}

function safeObserverPlan(platform, options = {}) {
  try {
    return buildMeetingAppRuntimeObserverPlan({ platform }, options);
  } catch {
    return {};
  }
}

function routePolicy(platform) {
  if (platform === 'local_detector') {
    return {
      recommended_mode: 'trusted_host_detector_only',
      route_order: ['host_detector_axis', 'annotation_insert', 'post_meeting_import_if_provider_known'],
      pilot_gate: 'local_detector_start_end_records',
      production_gate: 'host_detector_replay_evidence',
    };
  }
  return {
    recommended_mode: 'local_observer_first_provider_reconcile',
    route_order: ['local_observer_axis', 'annotation_insert', 'speaker_position_markers', 'provider_reconcile', 'post_meeting_artifact_import'],
    pilot_gate: 'meetingAppRecordSet_with_active_and_ended_snapshots',
    production_gate: 'meetingAppRecordSet_plus_providerRecords',
  };
}

function localObserverRoute(platform, runtimeProfile = {}, runtimeAdapter = {}, observerPlan = {}) {
  const localDetector = platform === 'local_detector';
  return compactObject({
    route: localDetector ? 'host_detector_axis' : 'local_observer_axis',
    priority: 1,
    role: localDetector ? 'authoritative_realtime_axis' : 'primary_low_latency_axis',
    required_for_realtime: true,
    blocks_realtime_if_missing: true,
    source: localDetector ? 'trusted_host_detector' : 'browser_extension_or_native_detector',
    surfaces: localDetector ? ['host_sdk'] : ['browser_extension', 'native_detector'],
    sdk_modules: compactObject({
      browser_runtime: localDetector ? undefined : '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-browser-runtime',
      native_runtime: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-track-runtime',
      runtime_host: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-runtime-host',
      realtime_annotation: '@ai-annotation/meeting-timeline-sdk/adapters/platform-realtime-annotation',
    }),
    browser_matches: runtimeAdapter.extension?.matches ?? [],
    native_entrypoint: NATIVE_ENTRYPOINTS[platform],
    observer_runtime: observerPlan.observer_runtime,
    start_condition: runtimeProfile.axis?.start?.create_on,
    end_condition: runtimeProfile.axis?.end?.create_on,
    timestamp_field: runtimeProfile.runtime_contract?.annotation_timestamp_field ?? 'captured_at_ms',
    evidence_input: localDetector ? 'providerRecords_or_local_detector_records' : 'meetingAppRecordSet',
  });
}

function annotationRoute(runtimeProfile = {}, runtimeEventPlan = {}) {
  return {
    route: 'annotation_insert',
    priority: 2,
    role: 'place_human_marks_on_current_axis',
    required_for_realtime: true,
    blocks_realtime_if_missing: true,
    timestamp_field: runtimeProfile.runtime_contract?.annotation_timestamp_field ?? 'captured_at_ms',
    invariant: 'always_insert_by_absolute_captured_at_ms_not_provider_delivery_time',
    runtime_event_endpoint: runtimeEventPlan.endpoint,
    supported_actions: runtimeEventPlan.supported_actions ?? [],
  };
}

function speakerRoute(runtimeProfile = {}) {
  return compactObject({
    route: 'speaker_position_markers',
    priority: 3,
    role: 'optional_timeline_positions_only',
    required_for_realtime: false,
    blocks_realtime_if_missing: false,
    content_policy: runtimeProfile.speaker_markers?.content_policy,
    primary_source: runtimeProfile.speaker_markers?.primary_source,
    filter: runtimeProfile.speaker_markers?.filter,
    output: runtimeProfile.speaker_markers?.marker_payload_contract,
  });
}

function providerRoute(platform, provider = {}, runtimeProfile = {}) {
  if (platform === 'local_detector') return undefined;
  return compactObject({
    route: 'provider_reconcile',
    priority: 4,
    role: 'authoritative_reconcile_and_production_evidence',
    required_for_realtime: false,
    required_for_production: true,
    blocks_realtime_if_missing: false,
    transport: provider.transport,
    endpoint: provider.endpoint,
    status_endpoint: provider.status_endpoint,
    start_events: runtimeProfile.provider_events?.start_events ?? [],
    end_events: runtimeProfile.provider_events?.end_events ?? [],
    participant_events: runtimeProfile.provider_events?.participant_events ?? [],
    lifecycle_events: runtimeProfile.provider_events?.lifecycle_events ?? [],
    security: provider.security,
    readiness: provider.readiness,
    evidence_input: 'providerRecords',
    references: PROVIDER_REFERENCES[platform],
  });
}

function artifactRoute(capabilities = {}, runtimeProfile = {}) {
  return compactObject({
    route: 'post_meeting_artifact_import',
    priority: 5,
    role: 'transcript_recording_backfill_after_meeting',
    required_for_realtime: false,
    blocks_realtime_if_missing: false,
    transcript: runtimeProfile.transcript,
    artifact_events: runtimeProfile.provider_events?.artifact_events ?? [],
    import_endpoint: capabilities.post_meeting_transcript?.import_endpoint,
    normalizer: capabilities.post_meeting_transcript?.sdk_normalizer,
  });
}

function buildRoutes(platform, capabilities, provider, runtimeProfile, runtimeAdapter, runtimeEventPlan, observerPlan) {
  return [
    localObserverRoute(platform, runtimeProfile, runtimeAdapter, observerPlan),
    annotationRoute(runtimeProfile, runtimeEventPlan),
    speakerRoute(runtimeProfile),
    providerRoute(platform, provider, runtimeProfile),
    artifactRoute(capabilities, runtimeProfile),
  ].filter(Boolean);
}

export function buildMeetingPlatformAdapterRoute(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const capabilities = platformCapabilityContract(key, options);
  const integration = buildPlatformIntegrationPlan(key, options);
  const runtimeProfile = buildMeetingPlatformRuntimeProfile(key, options);
  const provider = buildMeetingPlatformProviderConnectionPack(key, options);
  const runtimeEventPlan = buildMeetingPlatformRuntimeEventPlan(key, options);
  const runtimeAdapter = key === 'local_detector' ? {} : safeRuntimeAdapter(key, options);
  const observerPlan = key === 'local_detector' ? {} : safeObserverPlan(key, options);
  const policy = routePolicy(key);
  const routes = buildRoutes(key, capabilities, provider, runtimeProfile, runtimeAdapter, runtimeEventPlan, observerPlan);

  return compactObject({
    type: 'meeting_platform_adapter_route',
    schema: MEETING_PLATFORM_ADAPTER_ROUTE_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_ROUTE_SCHEMA_VERSION,
    platform: key,
    display_name: capabilities.display_name ?? integration.display_name,
    objective: 'choose_the_low_latency_adapter_path_for_realtime_meeting_timeline_annotations',
    recommended_mode: policy.recommended_mode,
    route_order: policy.route_order,
    route_count: routes.length,
    routes,
    adapter_surfaces: runtimeProfile.adapter_surfaces,
    launch_requirements: runtimeProfile.launch_requirements,
    evidence_thresholds: runtimeProfile.evidence_thresholds,
    fallback_policy: runtimeProfile.fallback_policy,
    entrypoints: {
      browser_extension: key === 'local_detector' ? undefined : {
        matches: runtimeAdapter.extension?.matches ?? [],
        permissions: runtimeAdapter.extension?.permissions ?? [],
        runtime_preset: runtimeAdapter.runtime_options?.runtimePreset,
        capture_profile: runtimeAdapter.capture_options?.captureProfile,
      },
      native_detector: NATIVE_ENTRYPOINTS[key],
      provider_webhook: key === 'local_detector' ? undefined : {
        endpoint: provider.endpoint,
        transport: provider.transport,
        references: PROVIDER_REFERENCES[key],
      },
      post_meeting_artifact: {
        transcript_source: capabilities.post_meeting_transcript?.source,
        import_endpoint: capabilities.post_meeting_transcript?.import_endpoint,
      },
    },
    realtime_invariants: {
      primary_axis_must_be_created_before_provider_reconcile: key !== 'local_detector',
      annotations_use_absolute_captured_at_ms: true,
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      per_meeting_annotation_isolation_required: true,
    },
    gates: {
      pilot: policy.pilot_gate,
      production: policy.production_gate,
      runtime_event_plan: runtimeEventPlan.schema,
      handoff_readiness: '@ai-annotation/meeting-timeline-sdk/adapters/platform-handoff-readiness',
    },
    implementation_order: runtimeProfile.implementation_order,
    next_actions: unique([
      ...(runtimeProfile.next_actions ?? []),
      key === 'local_detector' ? 'wire_host_detector_start_end_events' : 'capture_browser_or_native_observer_snapshots',
      key === 'local_detector' ? undefined : 'wire_provider_webhook_as_non_blocking_reconcile',
      'validate_annotations_land_by_captured_at_ms',
    ]),
  });
}

export function summarizeMeetingPlatformAdapterRoute(adapterRoute = {}) {
  if (!adapterRoute || typeof adapterRoute !== 'object') return undefined;
  const invariants = adapterRoute.realtime_invariants ?? {};
  return compactObject({
    platform: adapterRoute.platform,
    recommended_mode: adapterRoute.recommended_mode,
    primary_surface: adapterRoute.adapter_surfaces?.primary,
    surface_order: adapterRoute.adapter_surfaces?.recommended_order,
    first_route: adapterRoute.routes?.[0]?.route,
    route_count: adapterRoute.route_count,
    provider_events_block_realtime: invariants.provider_events_block_realtime,
    transcript_blocks_realtime: invariants.transcript_blocks_realtime,
    annotations_use_absolute_captured_at_ms: invariants.annotations_use_absolute_captured_at_ms,
    per_meeting_annotation_isolation_required: invariants.per_meeting_annotation_isolation_required,
  });
}

export function verifyMeetingPlatformAdapterRouteReadiness(adapterRoute = {}) {
  const invariants = adapterRoute.realtime_invariants ?? {};
  const axisRoute = adapterRoute.platform === 'local_detector' ? 'host_detector_axis' : 'local_observer_axis';
  const issues = [
    adapterRoute.schema === MEETING_PLATFORM_ADAPTER_ROUTE_SCHEMA
      ? undefined
      : issue('invalid_adapter_route_schema', 'Adapter route must use meeting_platform_adapter_route schema.'),
    adapterRoute.route_count > 0
      ? undefined
      : issue('missing_adapter_routes', 'Adapter route must contain at least one route.'),
    routeHasRequiredRealtimeStep(adapterRoute, axisRoute)
      ? undefined
      : issue('missing_primary_axis_route', 'Adapter route must include the realtime axis route.', { required_route: axisRoute }),
    routeHasRequiredRealtimeStep(adapterRoute, 'annotation_insert')
      ? undefined
      : issue('missing_annotation_insert_route', 'Adapter route must include the realtime annotation insert route.'),
    invariants.annotations_use_absolute_captured_at_ms === true
      ? undefined
      : issue('annotations_must_use_captured_at_ms', 'Annotations must use absolute captured_at_ms timestamps.'),
    invariants.provider_events_block_realtime === false
      ? undefined
      : issue('provider_events_block_realtime', 'Provider events must not block realtime annotations.'),
    invariants.transcript_blocks_realtime === false
      ? undefined
      : issue('transcript_blocks_realtime', 'Transcript import must not block realtime annotations.'),
    invariants.per_meeting_annotation_isolation_required === true
      ? undefined
      : issue('missing_per_meeting_annotation_isolation', 'Adapter route must require per-meeting annotation isolation.'),
  ].filter(Boolean);
  return compactObject({
    type: 'meeting_platform_adapter_route_readiness',
    platform: adapterRoute.platform,
    ready: issues.length === 0,
    required_axis_route: axisRoute,
    missing: issues.map((item) => item.code),
    issues,
    summary: summarizeMeetingPlatformAdapterRoute(adapterRoute),
  });
}

export function buildMeetingPlatformAdapterRouteMatrix(options = {}) {
  const routes = selectedPlatforms(options).map((platform) => buildMeetingPlatformAdapterRoute(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  const routeReadiness = routes.map((route) => verifyMeetingPlatformAdapterRouteReadiness(route));
  return {
    type: 'meeting_platform_adapter_route_matrix',
    schema: MEETING_PLATFORM_ADAPTER_ROUTE_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_ROUTE_SCHEMA_VERSION,
    platform_count: routes.length,
    route_ready_count: routeReadiness.filter((readiness) => readiness.ready).length,
    local_observer_first_count: routes.filter((route) => route.recommended_mode === 'local_observer_first_provider_reconcile').length,
    provider_non_blocking_count: routes.filter((route) => route.realtime_invariants.provider_events_block_realtime === false).length,
    transcript_non_blocking_count: routes.filter((route) => route.realtime_invariants.transcript_blocks_realtime === false).length,
    platforms: routes.map((route) => route.platform),
    rows: routes.map((route) => ({
      platform: route.platform,
      display_name: route.display_name,
      recommended_mode: route.recommended_mode,
      primary_surface: route.adapter_surfaces?.primary,
      surface_order: route.adapter_surfaces?.recommended_order,
      provider_reconcile_surface: route.adapter_surfaces?.provider_reconcile_surface,
      first_route: route.routes[0]?.route,
      route_ready: routeReadiness.find((readiness) => readiness.platform === route.platform)?.ready,
      route_missing: routeReadiness.find((readiness) => readiness.platform === route.platform)?.missing ?? [],
      browser_match_count: route.entrypoints.browser_extension?.matches?.length ?? 0,
      provider_transport: route.entrypoints.provider_webhook?.transport,
      pilot_provider_records_required: route.evidence_thresholds?.pilot?.provider_records_required,
      production_provider_records_required: route.evidence_thresholds?.production?.provider_records_required,
      provider_blocks_realtime: route.realtime_invariants.provider_events_block_realtime,
      transcript_blocks_realtime: route.realtime_invariants.transcript_blocks_realtime,
      production_gate: route.gates.production,
    })),
    routes,
  };
}
