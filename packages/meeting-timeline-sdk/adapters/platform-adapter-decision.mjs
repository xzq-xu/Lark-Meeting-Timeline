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
  buildMeetingPlatformAdapterBlueprint,
} from './platform-adapter-blueprint.mjs';
import {
  MEETING_PLATFORM_KEYS,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_ADAPTER_DECISION_SCHEMA = 'meeting_platform_adapter_decision';
export const MEETING_PLATFORM_ADAPTER_DECISION_MATRIX_SCHEMA = 'meeting_platform_adapter_decision_matrix';
export const MEETING_PLATFORM_ADAPTER_DECISION_SCHEMA_VERSION = 1;
export const MEETING_PLATFORM_HOST_PROFILE_PRESETS = Object.freeze({
  browser_extension: Object.freeze({
    profile: 'browser_extension',
    label: 'Browser extension',
    available_surfaces: Object.freeze(['browser_extension']),
    realtime_axis_surface: 'browser_extension',
    notes: 'Use a content script or equivalent browser DOM observer as the realtime axis surface.',
  }),
  webview_preload: Object.freeze({
    profile: 'webview_preload',
    label: 'WebView preload',
    available_surfaces: Object.freeze(['browser_extension']),
    realtime_axis_surface: 'browser_extension',
    notes: 'Use the browser-extension adapter contract through an Electron/WebView preload bridge.',
  }),
  native_detector: Object.freeze({
    profile: 'native_detector',
    label: 'Native detector',
    available_surfaces: Object.freeze(['native_detector']),
    realtime_axis_surface: 'native_detector',
    notes: 'Use a native window, accessibility, audio, or desktop detector as the realtime axis surface.',
  }),
  hybrid_desktop: Object.freeze({
    profile: 'hybrid_desktop',
    label: 'Hybrid desktop',
    available_surfaces: Object.freeze(['native_detector', 'browser_extension']),
    realtime_axis_surface: 'native_detector_or_browser_extension',
    notes: 'Prefer native detection where it is the platform default, but allow browser DOM observation fallback.',
  }),
  browser_or_native: Object.freeze({
    profile: 'browser_or_native',
    label: 'Browser or native',
    available_surfaces: Object.freeze(['browser_extension', 'native_detector']),
    realtime_axis_surface: 'browser_extension_or_native_detector',
    notes: 'Allow either local observer family; provider events remain nonblocking reconcile only.',
  }),
  full_local: Object.freeze({
    profile: 'full_local',
    label: 'Full local observer',
    available_surfaces: Object.freeze(['browser_extension', 'native_detector', 'host_detector']),
    realtime_axis_surface: 'local_observer',
    notes: 'All local realtime surfaces are available; provider events are still not required for realtime marks.',
  }),
  full: Object.freeze({
    profile: 'full',
    label: 'Full host integration',
    available_surfaces: Object.freeze(['browser_extension', 'native_detector', 'host_detector', 'provider_reconcile']),
    realtime_axis_surface: 'local_observer_with_provider_reconcile',
    notes: 'Local realtime surfaces plus provider reconcile/backfill are available.',
  }),
  provider_reconcile_only: Object.freeze({
    profile: 'provider_reconcile_only',
    label: 'Provider reconcile only',
    available_surfaces: Object.freeze(['provider_reconcile']),
    realtime_axis_surface: 'none',
    notes: 'Provider events can reconcile or backfill, but cannot prove a low-latency realtime annotation axis.',
  }),
  local_detector: Object.freeze({
    profile: 'local_detector',
    label: 'Local detector',
    available_surfaces: Object.freeze(['host_detector']),
    realtime_axis_surface: 'host_detector',
    notes: 'Use an explicit trusted host detector or manual controller for the realtime axis.',
  }),
});

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

function normalizeHostProfile(value) {
  return normalizeSurface(value);
}

function hostProfileInput(input = {}, options = {}) {
  return firstNonEmpty(
    options.hostProfile,
    options.host_profile,
    options.hostRuntimeProfile,
    options.host_runtime_profile,
    input.hostProfile,
    input.host_profile,
    input.hostRuntimeProfile,
    input.host_runtime_profile,
  );
}

function hostProfilePreset(input = {}, options = {}) {
  const raw = hostProfileInput(input, options);
  if (!raw) return undefined;
  return MEETING_PLATFORM_HOST_PROFILE_PRESETS[normalizeHostProfile(raw)];
}

function capabilityEnabled(value) {
  if (value == null) return false;
  if (value === true) return true;
  if (value === false) return false;
  const text = String(value).trim().toLowerCase();
  return ['1', 'true', 'yes', 'on', 'available', 'supported', 'enabled', 'ready'].includes(text);
}

function capabilitySurfaceAliases(key) {
  const normalized = normalizeSurface(key);
  if (!normalized) return [];
  if ([
    'browser',
    'browser_observer',
    'browser_dom_observer',
    'browser_extension',
    'extension',
    'content_script',
  ].includes(normalized)) return ['browser_extension'];
  if ([
    'webview',
    'webview_preload',
    'electron_preload',
  ].includes(normalized)) return ['webview_preload'];
  if ([
    'native',
    'native_detector',
    'native_window_observer',
    'desktop',
    'desktop_observer',
    'desktop_detector',
    'accessibility',
    'accessibility_agent',
  ].includes(normalized)) return ['native_detector'];
  if ([
    'provider',
    'provider_events',
    'provider_event',
    'provider_webhook',
    'provider_reconcile',
    'webhook',
    'long_connection',
  ].includes(normalized)) return ['provider_reconcile'];
  if ([
    'host',
    'host_sdk',
    'host_detector',
    'local_detector',
    'manual',
    'manual_controller',
  ].includes(normalized)) return ['host_detector'];
  return [normalized];
}

function surfaceAvailabilityInput(input = {}, options = {}) {
  return firstNonEmpty(
    options.availableSurfaces,
    options.available_surfaces,
    options.surfaceCapabilities,
    options.surface_capabilities,
    options.hostSurfaces,
    options.host_surfaces,
    input.availableSurfaces,
    input.available_surfaces,
    input.surfaceCapabilities,
    input.surface_capabilities,
    input.hostSurfaces,
    input.host_surfaces,
    hostProfilePreset(input, options)?.available_surfaces,
  );
}

function hostCapabilitiesInput(input = {}, options = {}) {
  return firstNonEmpty(
    options.hostCapabilities,
    options.host_capabilities,
    options.capabilities,
    input.hostCapabilities,
    input.host_capabilities,
    input.capabilities,
  );
}

function configuredAvailableSurfaces(input = {}, options = {}) {
  const declared = surfaceAvailabilityInput(input, options);
  const caps = hostCapabilitiesInput(input, options);
  const profile = hostProfilePreset(input, options);
  const values = [];
  let specified = false;
  if (profile) specified = true;
  if (declared != null) {
    specified = true;
    if (Array.isArray(declared) || typeof declared === 'string' || typeof declared[Symbol.iterator] === 'function') {
      values.push(...asArray(declared));
    } else if (isPlainObject(declared)) {
      for (const [key, value] of Object.entries(declared)) {
        if (capabilityEnabled(value)) values.push(key);
      }
    }
  }
  if (isPlainObject(caps)) {
    specified = true;
    const capabilityPairs = {
      browser_extension: firstNonEmpty(caps.browserExtension, caps.browser_extension, caps.browser, caps.browserObserver, caps.browser_observer),
      webview_preload: firstNonEmpty(caps.webviewPreload, caps.webview_preload, caps.webview, caps.electronPreload, caps.electron_preload),
      native_detector: firstNonEmpty(caps.nativeDetector, caps.native_detector, caps.native, caps.desktopObserver, caps.desktop_observer, caps.accessibilityAgent, caps.accessibility_agent),
      provider_reconcile: firstNonEmpty(caps.providerEvents, caps.provider_events, caps.providerWebhook, caps.provider_webhook, caps.providerReconcile, caps.provider_reconcile, caps.webhook),
      host_detector: firstNonEmpty(caps.hostDetector, caps.host_detector, caps.localDetector, caps.local_detector, caps.hostSdk, caps.host_sdk),
    };
    for (const [surface, enabled] of Object.entries(capabilityPairs)) {
      if (capabilityEnabled(enabled)) values.push(surface);
    }
  }
  if (!specified) return undefined;
  return unique(values.flatMap((value) => capabilitySurfaceAliases(value)));
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

function blueprintSummary(blueprint = {}, selectedSurface) {
  const surface = selectedSurface ? blueprint.surfaces?.[selectedSurface] : undefined;
  return compactObject({
    schema: blueprint.schema,
    ready: blueprint.readiness?.ready === true,
    primary_surface: blueprint.primary_surface,
    surface_order: blueprint.surface_order,
    selected_surface: selectedSurface,
    selected_surface_recommended: surface?.recommended === true,
    selected_surface_priority: surface?.priority,
    first_acceptance_gate: blueprint.acceptance_gates?.realtime_pilot?.[0],
    realtime_pilot_gates: blueprint.acceptance_gates?.realtime_pilot,
    production_gates: blueprint.acceptance_gates?.production,
    provider_blocks_realtime: blueprint.runtime_contract?.provider_events_block_realtime,
    transcript_blocks_realtime: blueprint.runtime_contract?.transcript_blocks_realtime,
    timestamp_field: blueprint.realtime_axis_contract?.timestamp_field,
    sdk_method: 'platformAdapterBlueprint',
  });
}

function startupSurfaceFromBlueprint(blueprint = {}) {
  const order = asArray(blueprint.surface_order)
    .map((surface) => normalizeSurface(surface))
    .map((surface) => (surface === 'desktop_observer' ? 'native_detector' : surface))
    .filter((surface) => surface && surface !== 'provider_reconcile');
  if (order.length > 0) return order[0];
  const primary = normalizeSurface(blueprint.primary_surface);
  if (primary === 'desktop_or_browser_observer') return 'native_detector';
  if (primary === 'browser_extension_or_native_detector') return 'browser_extension';
  if (primary === 'browser_extension_or_desktop_observer') return 'browser_extension';
  if (primary === 'desktop_observer') return 'native_detector';
  return primary;
}

function expandedSurface(value) {
  const surface = normalizeSurface(value);
  if (!surface) return [];
  if (surface === 'desktop_observer') return ['native_detector'];
  if (surface === 'desktop_or_browser_observer') return ['native_detector', 'browser_extension'];
  if (surface === 'browser_extension_or_native_detector') return ['browser_extension', 'native_detector'];
  if (surface === 'browser_extension_or_desktop_observer') return ['browser_extension', 'native_detector'];
  return [surface];
}

function hostSurfaceCompatibility(input = {}, options = {}, adapterBlueprint = {}, rawSurface = {}) {
  const configured = configuredAvailableSurfaces(input, options);
  const profile = hostProfilePreset(input, options);
  const candidateSurfaces = surfaceOrderFromBlueprint(adapterBlueprint);
  const selectedFromRaw = expandedSurface(rawSurface.surface);
  const selectedCandidates = selectedFromRaw.length > 0 ? selectedFromRaw : candidateSurfaces;
  if (!configured) {
    return {
      specified: false,
      host_profile: profile?.profile,
      available_surfaces: candidateSurfaces,
      candidate_surfaces: candidateSurfaces,
      selected_surface: rawSurface.surface,
      selected_surface_available: true,
      source: rawSurface.source,
      missing_surfaces: [],
    };
  }
  const availableSet = new Set(configured);
  const selectable = candidateSurfaces.filter((surface) => availableSet.has(surface));
  const rawAvailable = selectedCandidates.find((surface) => availableSet.has(surface));
  const selected = rawSurface.source === 'explicit'
    ? rawSurface.surface
    : firstNonEmpty(rawAvailable, selectable[0], rawSurface.surface);
  const selectedAvailable = expandedSurface(selected).some((surface) => availableSet.has(surface));
  return {
    specified: true,
    host_profile: profile?.profile,
    host_profile_label: profile?.label,
    host_profile_notes: profile?.notes,
    available_surfaces: configured,
    candidate_surfaces: candidateSurfaces,
    compatible_surfaces: selectable,
    selected_surface: selected,
    selected_surface_available: selectedAvailable,
    source: rawSurface.source === 'explicit'
      ? 'explicit'
      : selected !== rawSurface.surface
        ? 'host_capabilities'
        : rawSurface.source,
    missing_surfaces: candidateSurfaces.filter((surface) => !availableSet.has(surface)),
  };
}

function surfaceOrderFromBlueprint(blueprint = {}) {
  return unique([
    ...asArray(blueprint.surface_order).flatMap((surface) => expandedSurface(surface)),
    ...expandedSurface(blueprint.primary_surface),
    ...Object.keys(blueprint.surfaces ?? {}).flatMap((surface) => expandedSurface(surface)),
  ]).filter((surface) => surface !== 'post_meeting_artifact');
}

function observerModeForSurface(surface) {
  if (surface === 'browser_extension') return 'browser_dom_observer';
  if (surface === 'native_detector') return 'native_window_observer';
  if (surface === 'provider_reconcile') return 'provider_reconcile_only';
  return 'host_surface_observer';
}

function runtimeFactoryForSurface(surface) {
  if (surface === 'browser_extension') return 'createMeetingAppBrowserRuntime';
  if (surface === 'native_detector') return 'createMeetingAppTrackRuntime';
  return undefined;
}

function hostInstallTargetForSurface(surface) {
  if (surface === 'browser_extension') return 'browser_extension_or_webview_preload';
  if (surface === 'native_detector') return 'native_detector_or_desktop_accessibility_agent';
  if (surface === 'provider_reconcile') return 'provider_webhook_or_long_connection';
  return 'host_surface_observer';
}

function evidenceToCollectFirst(surface) {
  if (surface === 'browser_extension') {
    return [
      'live_browser_dom_snapshot_active_meeting',
      'adapter_preflight_current_window_accepts_active_meeting',
      'annotation_insert_current_axis_with_captured_at_ms',
    ];
  }
  if (surface === 'native_detector') {
    return [
      'native_window_or_accessibility_snapshot_active_meeting',
      'adapter_preflight_candidate_window_accepts_active_meeting',
      'annotation_insert_current_axis_with_captured_at_ms',
    ];
  }
  if (surface === 'provider_reconcile') {
    return [
      'provider_start_end_delivery_sample',
      'local_observer_or_manual_axis_fallback_sample',
      'annotation_insert_uses_capture_time_not_provider_delivery_time',
    ];
  }
  return [
    'host_surface_snapshot_active_meeting',
    'annotation_insert_current_axis_with_captured_at_ms',
  ];
}

function hostIntegrationStep(order, id, phase, required, fields = {}) {
  return compactObject({
    order,
    id,
    phase,
    required,
    ...fields,
  });
}

function hostIntegrationChecklist(platform, selectedSurface, adapterBlueprint = {}, route = {}) {
  const localSurface = selectedSurface !== 'provider_reconcile';
  const providerRoute = routeByName(route.routes, 'provider_reconcile');
  const transcriptRoute = routeByName(route.routes, 'post_meeting_artifact_import');
  const steps = [
    hostIntegrationStep(1, 'resolve_adapter_decision', 'startup', true, {
      sdk_method: 'platformAdapterDecision',
      output: 'selected_surface_and_adaptation_strategy',
    }),
    hostIntegrationStep(2, 'run_adapter_preflight', 'startup', true, {
      sdk_method: selectedSurface === 'browser_extension'
        ? 'platformAdapterCurrentWindowPreflight'
        : 'platformAdapterCandidatePreflight',
      required_before: 'first_insert_annotation',
      live_evidence_required: true,
    }),
    localSurface ? hostIntegrationStep(3, 'install_local_observer_runtime', 'runtime', true, {
      observer_mode: observerModeForSurface(selectedSurface),
      runtime_factory: runtimeFactoryForSurface(selectedSurface),
      install_target: hostInstallTargetForSurface(selectedSurface),
    }) : hostIntegrationStep(3, 'add_local_surface_for_realtime_axis', 'runtime', true, {
      status: 'missing_local_surface_for_realtime',
      fallback_surfaces: surfaceOrderFromBlueprint(adapterBlueprint).filter((surface) => surface !== 'provider_reconcile'),
    }),
    hostIntegrationStep(4, 'bind_current_axis', 'axis', true, {
      sdk_method: localSurface ? 'observePlatformCandidates' : 'observePlatformCandidates_from_fallback_local_surface',
      timestamp_field: 'captured_at_ms',
      must_precede: 'insert_realtime_annotation',
    }),
    hostIntegrationStep(5, 'insert_realtime_annotation', 'mark', true, {
      sdk_method: 'insertAnnotation',
      timestamp_field: 'captured_at_ms',
      invariant: 'use_capture_time_not_provider_delivery_time',
    }),
    hostIntegrationStep(6, 'emit_speaker_position_markers', 'optional_track', false, {
      sdk_method: 'speakerTrack',
      content_policy: 'speaker_position_only_no_transcript_text_required',
    }),
    providerRoute ? hostIntegrationStep(7, 'ingest_provider_reconcile', 'post_realtime_reconcile', false, {
      sdk_method: 'ingestProvider',
      transport: providerRoute.transport,
      required_for_production: providerRoute.required_for_production === true,
      blocks_realtime_annotation: false,
    }) : undefined,
    transcriptRoute ? hostIntegrationStep(8, 'import_post_meeting_artifacts', 'post_meeting', false, {
      import_endpoint: transcriptRoute.import_endpoint,
      normalizer: transcriptRoute.normalizer,
      blocks_realtime_annotation: false,
    }) : undefined,
  ].filter(Boolean);
  return {
    platform,
    selected_surface: selectedSurface,
    ready_for_realtime_host_wiring: localSurface,
    required_step_count: steps.filter((step) => step.required).length,
    optional_step_count: steps.filter((step) => !step.required).length,
    steps,
  };
}

function adaptationRiskTags(platform, selectedSurface, fallbackSurfaces = []) {
  return unique([
    selectedSurface === 'browser_extension' ? 'dom_selectors_can_drift_collect_live_snapshot_corpus' : undefined,
    selectedSurface === 'native_detector' ? 'native_detection_needs_window_or_accessibility_permission' : undefined,
    selectedSurface === 'provider_reconcile' ? 'provider_only_surface_cannot_prove_low_latency_axis_without_local_observer' : undefined,
    fallbackSurfaces.length > 0 ? 'multi_surface_fallback_required' : undefined,
    platform === 'microsoft_teams' ? 'teams_users_may_switch_between_desktop_and_browser_surfaces' : undefined,
    'provider_events_may_lag_and_must_not_block_realtime_marks',
    'transcript_or_recording_artifacts_are_post_meeting_only',
  ]);
}

function adaptationStrategy(platform, selectedSurface, adapterBlueprint = {}, route = {}) {
  const surfaceOrder = surfaceOrderFromBlueprint(adapterBlueprint);
  const recommendedFirstSurface = startupSurfaceFromBlueprint(adapterBlueprint) ?? surfaceOrder[0] ?? selectedSurface;
  const fallbackSurfaces = surfaceOrder.filter((surface) => surface !== selectedSurface);
  const providerRoute = routeByName(route.routes, 'provider_reconcile');
  const transcriptRoute = routeByName(route.routes, 'post_meeting_artifact_import');
  const selectedSurfaceConfig = adapterBlueprint.surfaces?.[selectedSurface] ?? {};
  return compactObject({
    platform,
    selected_surface: selectedSurface,
    recommended_first_surface: recommendedFirstSurface,
    fallback_surfaces: fallbackSurfaces,
    selected_observer_mode: observerModeForSurface(selectedSurface),
    selected_surface_role: selectedSurfaceConfig.role,
    selected_surface_priority: selectedSurfaceConfig.priority,
    selected_surface_evidence_kind: selectedSurfaceConfig.evidence?.evidence_kind,
    runtime_factory: runtimeFactoryForSurface(selectedSurface),
    host_install_target: hostInstallTargetForSurface(selectedSurface),
    primary_axis_source: adapterBlueprint.realtime_axis_contract?.primary_source ?? 'local_observer',
    provider_reconcile_role: providerRoute
      ? 'nonblocking_production_reconcile'
      : undefined,
    provider_required_for_realtime: providerRoute?.required_for_realtime === true,
    provider_blocks_realtime_if_missing: providerRoute?.blocks_realtime_if_missing === true,
    transcript_role: transcriptRoute
      ? 'post_meeting_nonblocking_backfill'
      : undefined,
    transcript_required_for_realtime: transcriptRoute?.required_for_realtime === true,
    realtime_dependencies: {
      local_observer_required: true,
      selected_surface_provides_local_observer: selectedSurface !== 'provider_reconcile',
      adapter_preflight_required_before_first_annotation: true,
      provider_event_required: false,
      transcript_required: false,
      timestamp_field: 'captured_at_ms',
    },
    evidence_to_collect_first: evidenceToCollectFirst(selectedSurface),
    host_integration_checklist: hostIntegrationChecklist(platform, selectedSurface, adapterBlueprint, route),
    production_evidence_gates: adapterBlueprint.acceptance_gates?.production,
    risk_tags: adaptationRiskTags(platform, selectedSurface, fallbackSurfaces),
  });
}

function decisionIssues(surface = {}, routeReadiness = {}, adapterBlueprintReady = false) {
  return [
    ...(routeReadiness.issues ?? []),
    surface.selected_surface_available === false ? {
      code: 'selected_surface_unavailable',
      message: 'Selected adapter surface is not available in the host capability set.',
      selected_surface: surface.selected_surface,
      available_surfaces: surface.available_surfaces,
    } : undefined,
    surface.selected_surface === 'provider_reconcile' ? {
      code: 'provider_reconcile_not_realtime_surface',
      message: 'Provider reconcile can backfill or verify the axis, but it cannot be the primary low-latency annotation surface.',
      selected_surface: surface.selected_surface,
    } : undefined,
    adapterBlueprintReady ? undefined : {
      code: 'adapter_blueprint_not_ready',
      message: 'Adapter blueprint must be ready before host wiring.',
    },
  ].filter(Boolean);
}

function decisionStatus(accepted, surface = {}) {
  if (accepted) return 'ready_for_realtime_annotation';
  if (surface.selected_surface_available === false) return 'host_surface_unavailable';
  if (surface.selected_surface === 'provider_reconcile') return 'needs_local_surface_for_realtime_axis';
  return 'needs_adapter_setup';
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
  const platformOptions = {
    ...options,
    platform,
  };
  const route = buildMeetingPlatformAdapterRoute(platform, platformOptions);
  const routeReadiness = verifyMeetingPlatformAdapterRouteReadiness(route);
  const adapterBlueprint = buildMeetingPlatformAdapterBlueprint(platform, platformOptions);
  const adapterBlueprintReady = adapterBlueprint.readiness?.ready === true;
  const rawSurface = surfaceFromInput(objectInput, platformOptions);
  const defaultSurface = rawSurface.source === 'default'
    ? {
      surface: startupSurfaceFromBlueprint(adapterBlueprint) ?? rawSurface.surface,
      source: 'adapter_blueprint',
    }
    : rawSurface;
  const hostCompatibility = hostSurfaceCompatibility(objectInput, platformOptions, adapterBlueprint, defaultSurface);
  const surface = {
    surface: hostCompatibility.selected_surface,
    source: hostCompatibility.source,
    host_compatibility: hostCompatibility,
  };
  const capability = buildMeetingAppAdapterCapabilityReport(platform, {
    ...platformOptions,
    input: inputForCapability(objectInput, platformOptions),
  });
  const executionPlan = buildMeetingAppAdapterExecutionPlan(capability, platformOptions);
  const firstRoute = route.routes?.[0];
  const providerRoute = routeByName(route.routes, 'provider_reconcile');
  const transcriptRoute = routeByName(route.routes, 'post_meeting_artifact_import');
  const hostDetectorReady = platform === 'local_detector'
    && routeReadiness.ready === true
    && adapterBlueprintReady
    && route.realtime_invariants?.provider_events_block_realtime === false;
  const localSurfaceSelected = surface.surface !== 'provider_reconcile';
  const hostSurfaceAvailable = hostCompatibility.selected_surface_available !== false;
  const realtimeReady = routeReadiness.ready === true
    && localSurfaceSelected
    && hostSurfaceAvailable
    && (executionPlan.realtime_ready === true || hostDetectorReady);
  const decision = {
    realtime_ready: realtimeReady,
  };
  const accepted = realtimeReady
    && adapterBlueprintReady
    && route.realtime_invariants?.provider_events_block_realtime === false;
  const issues = decisionIssues(hostCompatibility, routeReadiness, adapterBlueprintReady);
  return compactObject({
    type: 'meeting_platform_adapter_decision',
    schema: MEETING_PLATFORM_ADAPTER_DECISION_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_DECISION_SCHEMA_VERSION,
    accepted,
    realtime_ready: realtimeReady,
    status: decisionStatus(accepted, hostCompatibility),
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
    host_surface_compatibility: hostCompatibility,
    selected_route: firstRoute?.route,
    recommended_mode: route.recommended_mode,
    adaptation_strategy: adaptationStrategy(platform, surface.surface, adapterBlueprint, route),
    adapter_blueprint: blueprintSummary(adapterBlueprint, surface.surface),
    first_blocked_step: executionPlan.first_blocked_step,
    contracts: {
      timestamp_field: 'captured_at_ms',
      adapter_blueprint_required_before_host_wiring: true,
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
      adapter_blueprint: adapterBlueprint,
      capability,
      execution_plan: executionPlan,
    } : undefined,
    issues,
    next_actions: unique([
      ...issues.map((item) => item.code ? `adapter_decision:${item.code}` : undefined),
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
    platform,
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
    adapter_blueprint_ready_count: decisions.filter((decision) => decision.adapter_blueprint?.ready).length,
    host_surface_constrained_count: decisions.filter((decision) => decision.host_surface_compatibility?.specified === true).length,
    host_surface_compatible_count: decisions.filter((decision) => decision.host_surface_compatibility?.selected_surface_available !== false).length,
    local_surface_selected_count: decisions.filter((decision) => decision.selected_surface !== 'provider_reconcile').length,
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
      surface_source: decision.surface_source,
      host_profile: decision.host_surface_compatibility?.host_profile,
      host_surface_available: decision.host_surface_compatibility?.selected_surface_available !== false,
      host_available_surfaces: decision.host_surface_compatibility?.available_surfaces,
      host_compatible_surfaces: decision.host_surface_compatibility?.compatible_surfaces,
      selected_route: decision.selected_route,
      recommended_mode: decision.recommended_mode,
      adapter_blueprint_ready: decision.adapter_blueprint?.ready === true,
      adapter_blueprint_primary_surface: decision.adapter_blueprint?.primary_surface,
      adapter_blueprint_first_acceptance_gate: decision.adapter_blueprint?.first_acceptance_gate,
      recommended_first_surface: decision.adaptation_strategy?.recommended_first_surface,
      selected_observer_mode: decision.adaptation_strategy?.selected_observer_mode,
      fallback_surfaces: decision.adaptation_strategy?.fallback_surfaces,
      first_evidence_to_collect: decision.adaptation_strategy?.evidence_to_collect_first?.[0],
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
