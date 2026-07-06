import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  detectMeetingFromUrl,
} from './meeting-url.mjs';
import {
  buildMeetingAppAdapterCapabilityReport,
  buildMeetingAppAdapterExecutionPlan,
} from './meeting-app-adapter-capability.mjs';
import {
  buildMeetingPlatformAdapterRoute,
  verifyMeetingPlatformAdapterRouteReadiness,
} from './platform-adapter-route.mjs';
import {
  MEETING_PLATFORM_KEYS,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_ADAPTER_DECISION_SCHEMA = 'meeting_platform_adapter_decision';
export const MEETING_PLATFORM_ADAPTER_DECISION_MATRIX_SCHEMA = 'meeting_platform_adapter_decision_matrix';
export const MEETING_PLATFORM_ADAPTER_DECISION_SCHEMA_VERSION = 1;

const DEFAULT_DECISION_PLATFORMS = Object.freeze([
  'google_meet',
  'microsoft_teams',
  'zoom',
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

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function normalizeOptionalPlatform(value) {
  if (value == null || value === '') return undefined;
  return normalizeMeetingPlatform(value);
}

function safeDetectMeeting(input = {}) {
  try {
    return detectMeetingFromUrl(input);
  } catch {
    return undefined;
  }
}

function inputUrl(input = {}) {
  if (typeof input === 'string' || input instanceof URL) return String(input);
  return firstNonEmpty(
    input.url,
    input.meeting_url,
    input.meetingUrl,
    input.join_url,
    input.joinUrl,
    input.location?.href,
    input.window?.url,
    input.browser?.url,
    input.tab?.url,
    input.tabs?.find?.((tab) => tab?.active)?.url,
    input.tabs?.[0]?.url,
  );
}

function inputTitle(input = {}) {
  if (typeof input === 'string' || input instanceof URL) return undefined;
  return firstNonEmpty(
    input.title,
    input.meeting_title,
    input.meetingTitle,
    input.window?.title,
    input.tab?.title,
    input.tabs?.find?.((tab) => tab?.active)?.title,
    input.tabs?.[0]?.title,
  );
}

function platformFromInput(input = {}, options = {}) {
  const explicit = firstNonEmpty(
    options.platform,
    options.provider,
    options.adapter_key,
    options.adapterKey,
    input.platform,
    input.provider,
    input.adapter_key,
    input.adapterKey,
  );
  if (explicit) return {
    platform: normalizeMeetingPlatform(explicit),
    source: 'explicit',
  };
  const detected = safeDetectMeeting(input);
  if (detected?.platform) return {
    platform: normalizeMeetingPlatform(detected.platform),
    source: 'url_detection',
    detected,
  };
  const fallback = firstNonEmpty(options.defaultPlatform, options.default_platform, options.fallbackPlatform, options.fallback_platform);
  if (fallback) return {
    platform: normalizeMeetingPlatform(fallback),
    source: 'fallback',
  };
  return {
    platform: undefined,
    source: 'missing',
    detected,
  };
}

function normalizeSurface(value) {
  if (!value) return undefined;
  return String(value)
    .trim()
    .toLowerCase()
    .replaceAll('-', '_')
    .replaceAll(' ', '_');
}

function surfaceFromInput(input = {}, options = {}) {
  const explicit = normalizeSurface(firstNonEmpty(
    options.surface,
    options.adapterSurface,
    options.adapter_surface,
    input.surface,
    input.adapterSurface,
    input.adapter_surface,
  ));
  if (explicit) return {
    surface: explicit,
    source: 'explicit',
  };
  const rawUrl = inputUrl(input);
  if (rawUrl && /^https?:\/\//i.test(String(rawUrl))) return {
    surface: 'browser_extension',
    source: 'url',
  };
  if (input.window || input.process || input.app || input.native || input.native_detector) return {
    surface: 'native_detector',
    source: 'native_context',
  };
  if (input.provider_event || input.providerEvent || input.provider_events || input.providerEvents || options.providerOnly === true || options.provider_only === true) {
    return {
      surface: 'provider_reconcile',
      source: 'provider_event',
    };
  }
  return {
    surface: 'browser_extension',
    source: 'default',
  };
}

function inputForCapability(input = {}, options = {}) {
  if (options.input || options.snapshot || options.sample) return undefined;
  if (typeof input === 'string' || input instanceof URL) return {
    url: String(input),
  };
  return input;
}

function selectedPlatforms(input = {}, options = {}) {
  return unique(asArray(firstNonEmpty(
    options.platforms,
    options.platform_keys,
    input.platforms,
    input.platform_keys,
    DEFAULT_DECISION_PLATFORMS,
  )).map((platform) => normalizeMeetingPlatform(platform)));
}

function inputForPlatform(platform, input = {}, options = {}) {
  const sources = [
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
  ];
  const keys = unique([
    platform,
    platform.replaceAll('_', '-'),
    platform === 'microsoft_teams' ? 'teams' : undefined,
    platform === 'google_meet' ? 'google-meet' : undefined,
  ]);
  for (const source of sources) {
    if (!isPlainObject(source)) continue;
    for (const key of keys) {
      if (source[key] != null) return source[key];
    }
  }
  return input.input ?? input.snapshot ?? input.sample ?? options.input ?? options.snapshot ?? options.sample ?? {};
}

function routeByName(routes = [], name) {
  return routes.find((route) => route.route === name);
}

function runtimeActions(platform, surface, capability = {}, route = {}) {
  const actions = [
    {
      action: 'observe_platform_candidates',
      when: surface === 'browser_extension' ? 'before_first_realtime_annotation' : undefined,
      required: surface === 'browser_extension',
      sdk_method: 'observePlatformCandidates',
    },
    {
      action: 'observe_meeting_app',
      when: surface !== 'provider_reconcile' ? 'when_local_window_or_webview_snapshot_is_available' : undefined,
      required: surface !== 'provider_reconcile',
      sdk_method: 'observeMeetingApp',
    },
    {
      action: 'insert_annotation',
      when: 'on_human_mark_captured',
      required: true,
      sdk_method: 'insertAnnotation',
      timestamp_field: 'captured_at_ms',
    },
    {
      action: 'speaker_track',
      when: capability.timeline_capabilities?.speaker_track?.local_ready === true ? 'on_active_speaker_change' : undefined,
      required: false,
      sdk_method: 'speakerTrack',
    },
    {
      action: 'participant_track',
      when: capability.timeline_capabilities?.participant_track?.local_ready === true ? 'on_participant_or_roster_change' : undefined,
      required: false,
      sdk_method: 'participantTrack',
    },
    {
      action: 'provider_event',
      when: platform !== 'local_detector' ? 'when_provider_webhook_or_long_connection_arrives' : undefined,
      required: false,
      blocks_realtime_annotation: route.realtime_invariants?.provider_events_block_realtime === true,
      sdk_method: 'ingestProvider',
    },
  ];
  return actions.filter((action) => action.when || action.required);
}

function evidenceRequirements(decision = {}) {
  return {
    realtime_minimum: [
      'candidate_observation_or_local_detector_axis',
      'annotation_captured_at_ms',
    ],
    pilot: [
      'live_meeting_app_snapshot',
      'annotation_insert_current_axis_sample',
    ],
    production: [
      'provider_meeting_started_event',
      'provider_meeting_ended_event',
      'provider_events_do_not_block_realtime',
      'post_meeting_transcript_import_is_nonblocking',
    ],
    current_missing: decision.realtime_ready
      ? ['production_provider_evidence']
      : ['local_axis_or_provider_axis_signal'],
  };
}

export function buildMeetingPlatformAdapterDecision(input = {}, options = {}) {
  const objectInput = typeof input === 'string' || input instanceof URL
    ? { url: String(input) }
    : (input ?? {});
  const platformResolution = platformFromInput(objectInput, options);
  if (!platformResolution.platform) {
    return {
      type: 'meeting_platform_adapter_decision',
      schema: MEETING_PLATFORM_ADAPTER_DECISION_SCHEMA,
      schema_version: MEETING_PLATFORM_ADAPTER_DECISION_SCHEMA_VERSION,
      accepted: false,
      realtime_ready: false,
      status: 'missing_platform',
      platform: undefined,
      detected: platformResolution.detected,
      issues: [{
        code: 'missing_platform',
        message: 'Adapter decision requires an explicit platform or a detectable meeting URL.',
      }],
      next_actions: ['provide_platform_or_supported_meeting_url'],
    };
  }
  const platform = platformResolution.platform;
  const surface = surfaceFromInput(objectInput, options);
  const route = buildMeetingPlatformAdapterRoute(platform, options);
  const routeReadiness = verifyMeetingPlatformAdapterRouteReadiness(route);
  const capability = buildMeetingAppAdapterCapabilityReport(platform, {
    ...options,
    input: inputForCapability(objectInput, options),
  });
  const executionPlan = buildMeetingAppAdapterExecutionPlan(capability, options);
  const firstRoute = route.routes?.[0];
  const providerRoute = routeByName(route.routes, 'provider_reconcile');
  const transcriptRoute = routeByName(route.routes, 'post_meeting_artifact_import');
  const realtimeReady = routeReadiness.ready === true && executionPlan.realtime_ready === true;
  const decision = {
    realtime_ready: realtimeReady,
  };
  const accepted = realtimeReady && route.realtime_invariants?.provider_events_block_realtime === false;
  return compactObject({
    type: 'meeting_platform_adapter_decision',
    schema: MEETING_PLATFORM_ADAPTER_DECISION_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_DECISION_SCHEMA_VERSION,
    accepted,
    realtime_ready: realtimeReady,
    status: accepted ? 'ready_for_realtime_annotation' : 'needs_adapter_setup',
    platform,
    display_name: route.display_name ?? capability.display_name,
    detected: platformResolution.detected,
    platform_source: platformResolution.source,
    input: compactObject({
      url: inputUrl(objectInput),
      title: inputTitle(objectInput),
    }),
    selected_surface: surface.surface,
    surface_source: surface.source,
    selected_route: firstRoute?.route,
    recommended_mode: route.recommended_mode,
    first_blocked_step: executionPlan.first_blocked_step,
    contracts: {
      timestamp_field: 'captured_at_ms',
      provider_events_block_realtime: route.realtime_invariants?.provider_events_block_realtime,
      transcript_blocks_realtime: route.realtime_invariants?.transcript_blocks_realtime,
      local_axis_first: route.realtime_invariants?.primary_axis_must_be_created_before_provider_reconcile,
      per_meeting_annotation_isolation_required: route.realtime_invariants?.per_meeting_annotation_isolation_required,
    },
    timeline_capabilities: capability.timeline_capabilities,
    route_summary: {
      route_count: route.route_count,
      first_route: firstRoute?.route,
      provider_reconcile: providerRoute ? {
        required_for_realtime: providerRoute.required_for_realtime,
        required_for_production: providerRoute.required_for_production,
        blocks_realtime_if_missing: providerRoute.blocks_realtime_if_missing,
        transport: providerRoute.transport,
      } : undefined,
      post_meeting_artifact: transcriptRoute ? {
        required_for_realtime: transcriptRoute.required_for_realtime,
        blocks_realtime_if_missing: transcriptRoute.blocks_realtime_if_missing,
        import_endpoint: transcriptRoute.import_endpoint,
        normalizer: transcriptRoute.normalizer,
      } : undefined,
    },
    runtime_actions: runtimeActions(platform, surface.surface, capability, route),
    evidence_requirements: evidenceRequirements(decision),
    reports: options.includeReports === true || options.include_reports === true ? {
      route,
      route_readiness: routeReadiness,
      capability,
      execution_plan: executionPlan,
    } : undefined,
    next_actions: unique([
      ...(routeReadiness.issues ?? []).map((item) => `adapter_route:${item.code}`),
      ...(capability.next_actions ?? []),
      ...(executionPlan.next_actions ?? []),
      accepted ? 'wire_runtime_actions_into_host_surface' : 'fix_adapter_decision_before_realtime_rollout',
    ]),
  });
}

export function buildMeetingPlatformAdapterDecisionMatrix(input = {}, options = {}) {
  const platforms = selectedPlatforms(input, options);
  const decisions = platforms.map((platform) => buildMeetingPlatformAdapterDecision({
    ...inputForPlatform(platform, input, options),
    platform,
  }, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_adapter_decision_matrix',
    schema: MEETING_PLATFORM_ADAPTER_DECISION_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_DECISION_SCHEMA_VERSION,
    platform_count: decisions.length,
    accepted_count: decisions.filter((decision) => decision.accepted).length,
    realtime_ready_count: decisions.filter((decision) => decision.realtime_ready).length,
    browser_surface_count: decisions.filter((decision) => decision.selected_surface === 'browser_extension').length,
    native_surface_count: decisions.filter((decision) => decision.selected_surface === 'native_detector').length,
    provider_reconcile_surface_count: decisions.filter((decision) => decision.selected_surface === 'provider_reconcile').length,
    platforms,
    rows: decisions.map((decision) => ({
      platform: decision.platform,
      display_name: decision.display_name,
      accepted: decision.accepted,
      realtime_ready: decision.realtime_ready,
      status: decision.status,
      selected_surface: decision.selected_surface,
      selected_route: decision.selected_route,
      recommended_mode: decision.recommended_mode,
      first_blocked_step: decision.first_blocked_step,
      provider_events_block_realtime: decision.contracts?.provider_events_block_realtime,
      transcript_blocks_realtime: decision.contracts?.transcript_blocks_realtime,
      first_next_action: decision.next_actions?.[0],
    })),
    decisions,
    next_actions: unique(decisions.flatMap((decision) => decision.next_actions ?? [])),
  };
}

export function assertMeetingPlatformAdapterDecision(input = {}, options = {}) {
  const decision = buildMeetingPlatformAdapterDecision(input, options);
  if (decision.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting platform adapter decision is not accepted', {
      platform: decision.platform,
      status: decision.status,
      next_actions: decision.next_actions,
      decision,
    });
  }
  return decision;
}

export function assertMeetingPlatformAdapterDecisionMatrix(input = {}, options = {}) {
  const matrix = buildMeetingPlatformAdapterDecisionMatrix(input, options);
  if (matrix.accepted_count !== matrix.platform_count) {
    throw new MeetingTimelineSdkError('Meeting platform adapter decision matrix is not accepted', {
      platform_count: matrix.platform_count,
      accepted_count: matrix.accepted_count,
      failed_platforms: matrix.rows.filter((row) => row.accepted !== true).map((row) => row.platform),
      next_actions: matrix.next_actions,
      matrix,
    });
  }
  return matrix;
}
