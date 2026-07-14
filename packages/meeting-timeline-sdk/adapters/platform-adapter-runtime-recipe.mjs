import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import { detectMeetingFromUrl } from './meeting-url.mjs';
import {
  buildMeetingPlatformAdapterStartupPlan,
} from './platform-adapter-startup.mjs';
import {
  buildMeetingPlatformRawSignalExampleBatch,
} from './platform-raw-signal.mjs';
import {
  normalizeMeetingPlatform,
} from './platform-setup.mjs';
import {
  MEETING_PLATFORM_RUNTIME_EVENT_ENDPOINT,
} from './platform-runtime-event.mjs';

export const MEETING_PLATFORM_ADAPTER_RUNTIME_RECIPE_SCHEMA = 'meeting_platform_adapter_runtime_recipe';
export const MEETING_PLATFORM_ADAPTER_RUNTIME_RECIPE_MATRIX_SCHEMA = 'meeting_platform_adapter_runtime_recipe_matrix';
export const MEETING_PLATFORM_ADAPTER_RUNTIME_MANIFEST_SCHEMA = 'meeting_platform_adapter_runtime_manifest';
export const MEETING_PLATFORM_ADAPTER_RUNTIME_TARGET_SCHEMA = 'meeting_platform_adapter_runtime_target';
export const MEETING_PLATFORM_ADAPTER_RUNTIME_RECIPE_SCHEMA_VERSION = 1;

const SURFACE_ALIASES = Object.freeze({
  browser: 'browser_extension',
  browser_extension: 'browser_extension',
  'browser-extension': 'browser_extension',
  extension: 'browser_extension',
  webview: 'webview_preload',
  webview_preload: 'webview_preload',
  'webview-preload': 'webview_preload',
  electron: 'webview_preload',
  native: 'native_detector',
  native_detector: 'native_detector',
  'native-detector': 'native_detector',
  native_host: 'native_detector',
  'native-host': 'native_detector',
  host: 'native_detector',
  provider: 'provider_reconcile',
  provider_reconcile: 'provider_reconcile',
  'provider-reconcile': 'provider_reconcile',
});

const DEFAULT_RUNTIME_RECIPE_PLATFORMS = Object.freeze([
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

function issue(severity, code, message, details = {}) {
  return compactObject({ severity, code, message, ...details });
}

function normalizeSurface(value) {
  if (value == null || value === '') return undefined;
  const raw = String(value).trim();
  return SURFACE_ALIASES[raw] ?? SURFACE_ALIASES[raw.toLowerCase()] ?? raw;
}

function getPath(raw, path) {
  const parts = path.split('.');
  let node = raw;
  for (const part of parts) node = node?.[part];
  return node;
}

function firstPath(raw, paths = []) {
  return firstNonEmpty(...paths.map((path) => getPath(raw, path)));
}

function maybeNormalizePlatform(value) {
  if (value == null || value === '') return undefined;
  try {
    return normalizeMeetingPlatform(value);
  } catch {
    return undefined;
  }
}

function selectedPlatforms(input = {}, options = {}) {
  return unique(asArray(firstNonEmpty(
    options.platforms,
    options.platform_keys,
    input.platforms,
    input.platform_keys,
    DEFAULT_RUNTIME_RECIPE_PLATFORMS,
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

function objectInput(input = {}) {
  if (typeof input === 'string' || input instanceof URL) return { url: String(input) };
  return input ?? {};
}

function runtimeEventActionCounts(events = []) {
  const counts = {};
  for (const event of events) {
    if (!event?.action) continue;
    counts[event.action] = (counts[event.action] ?? 0) + 1;
  }
  return counts;
}

function endpointUrl(path, options = {}) {
  const baseUrl = firstNonEmpty(options.baseUrl, options.base_url);
  return baseUrl ? new URL(path, baseUrl).toString() : path;
}

function adapterModuleFor(platform) {
  const slug = platform === 'microsoft_teams' ? 'microsoft-teams' : platform.replaceAll('_', '-');
  return `@ai-annotation/meeting-timeline-sdk/adapters/${slug}`;
}

function bridgeKind(surface, plan = {}) {
  if (surface === 'browser_extension') return 'browser_content_script';
  if (surface === 'webview_preload') return 'webview_preload_bridge';
  if (surface === 'native_detector') return 'native_detector_runtime_event_client';
  if (surface === 'provider_reconcile') return 'provider_reconcile_only_not_realtime_axis';
  return plan.bridge?.preferred_bridge ?? 'host_surface_adapter';
}

function runtimeSequence(plan = {}, examples = {}) {
  const localSurface = plan.selected_surface !== 'provider_reconcile';
  return [
    {
      order: 1,
      action: 'resolve_adapter_decision',
      sdk_method: 'platformAdapterDecision',
      required: true,
      output: 'selected_surface_and_host_compatibility',
    },
    {
      order: 2,
      action: 'install_surface_runtime',
      sdk_method: plan.bridge?.create_function ?? plan.bridge?.first_method ?? 'createMeetingAppTimelineSdk',
      required: localSurface,
      surface: plan.selected_surface,
      bridge: plan.bridge?.preferred_bridge,
      install_target: plan.install_target,
    },
    localSurface ? {
      order: 3,
      action: 'observe_platform_candidates',
      sdk_method: 'observePlatformCandidates',
      message_type: plan.message_contract?.observe_candidates,
      required: true,
      timestamp_field: 'captured_at_ms',
    } : {
      order: 3,
      action: 'add_local_surface_before_realtime_marks',
      sdk_method: 'platformHostProfileCompatibilityMatrix',
      required: true,
      timestamp_field: 'captured_at_ms',
    },
    localSurface ? {
      order: 4,
      action: 'insert_realtime_annotation',
      sdk_method: 'insertAnnotation',
      message_type: plan.message_contract?.insert_annotation,
      required: true,
      timestamp_field: 'captured_at_ms',
    } : undefined,
    {
      order: 5,
      action: 'speaker_position_markers',
      sdk_method: 'speakerTrack',
      required: false,
      text_required: false,
      filter_policy: 'debounce_active_speaker_samples_before_drawing_position_markers',
      sample_runtime_event_count: examples.runtime_events?.filter((event) => event.action === 'speaker_track').length ?? 0,
    },
    {
      order: 6,
      action: 'participant_position_markers',
      sdk_method: 'participantTrack',
      required: false,
      text_required: false,
      sample_runtime_event_count: examples.runtime_events?.filter((event) => event.action === 'participant_track').length ?? 0,
    },
    {
      order: 7,
      action: 'provider_reconcile',
      sdk_method: 'ingestProvider',
      required: false,
      blocks_realtime_annotation: false,
      role: 'reconcile_or_backfill_only',
    },
  ].filter(Boolean);
}

function hostWiring(plan = {}, examples = {}) {
  const surface = plan.selected_surface;
  const localSurface = surface !== 'provider_reconcile';
  return compactObject({
    surface,
    host_profile: plan.decision?.host_surface_compatibility?.host_profile,
    bridge_kind: bridgeKind(surface, plan),
    install_target: plan.install_target,
    bridge: plan.bridge,
    first_required_method: localSurface ? 'observePlatformCandidates' : 'add_local_surface_for_realtime_annotations',
    insert_method: localSurface ? 'insertAnnotation' : undefined,
    optional_track_methods: ['speakerTrack', 'participantTrack'],
    provider_method: 'ingestProvider',
    runtime_event_actions: unique(examples.runtime_events?.map((event) => event.action) ?? []),
    runtime_event_action_counts: runtimeEventActionCounts(examples.runtime_events ?? []),
    message_contract: plan.message_contract,
    code_refs: plan.code_refs,
  });
}

function readinessFromPlan(plan = {}, examples = {}) {
  const localSurface = plan.selected_surface !== 'provider_reconcile';
  return {
    accepted: plan.accepted === true,
    realtime_startup_ready: plan.realtime_startup_ready === true,
    local_surface_selected: localSurface,
    provider_only: plan.selected_surface === 'provider_reconcile',
    raw_signal_examples_ready: (examples.runtime_event_count ?? 0) > 0,
    sample_insert_annotation_ready: (examples.runtime_events ?? []).some((event) => event.action === 'insert_annotation'),
    sample_speaker_track_ready: (examples.runtime_events ?? []).some((event) => event.action === 'speaker_track'),
    provider_events_block_realtime: plan.runtime_contract?.provider_events_block_realtime === true,
    transcript_blocks_realtime: plan.runtime_contract?.transcript_blocks_realtime === true,
  };
}

export function buildMeetingPlatformAdapterRuntimeRecipe(input = {}, options = {}) {
  const plan = buildMeetingPlatformAdapterStartupPlan(input, options);
  if (!plan.platform) {
    return {
      type: 'meeting_platform_adapter_runtime_recipe',
      schema: MEETING_PLATFORM_ADAPTER_RUNTIME_RECIPE_SCHEMA,
      schema_version: MEETING_PLATFORM_ADAPTER_RUNTIME_RECIPE_SCHEMA_VERSION,
      accepted: false,
      runtime_ready: false,
      status: 'missing_platform',
      plan,
      issues: plan.issues ?? [],
      next_actions: plan.next_actions ?? ['provide_platform_or_supported_meeting_url'],
    };
  }
  const examples = buildMeetingPlatformRawSignalExampleBatch({
    ...options,
    platforms: [plan.platform],
    platform_keys: undefined,
    filterActiveSpeakerSamples: firstNonEmpty(
      options.filterActiveSpeakerSamples,
      options.filter_active_speaker_samples,
      true,
    ),
  });
  const readiness = readinessFromPlan(plan, examples);
  const includeExamples = options.includeExamples !== false && options.include_examples !== false;
  return compactObject({
    type: 'meeting_platform_adapter_runtime_recipe',
    schema: MEETING_PLATFORM_ADAPTER_RUNTIME_RECIPE_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_RUNTIME_RECIPE_SCHEMA_VERSION,
    accepted: readiness.accepted && readiness.local_surface_selected,
    runtime_ready: readiness.realtime_startup_ready && readiness.raw_signal_examples_ready && readiness.local_surface_selected,
    status: readiness.local_surface_selected
      ? (readiness.realtime_startup_ready ? 'ready_to_wire_runtime' : 'needs_runtime_attention')
      : 'needs_local_surface_for_realtime_axis',
    platform: plan.platform,
    display_name: plan.display_name,
    selected_surface: plan.selected_surface,
    host_profile: plan.decision?.host_surface_compatibility?.host_profile,
    install_target: plan.install_target,
    runtime_contract: plan.runtime_contract,
    host_wiring: hostWiring(plan, examples),
    sequence: runtimeSequence(plan, examples),
    raw_signal_examples: includeExamples ? examples : {
      schema: examples.schema,
      signal_count: examples.signal_count,
      runtime_event_count: examples.runtime_event_count,
      filtered_speaker_event_count: examples.filtered_speaker_event_count,
      runtime_actions: unique(examples.runtime_events?.map((event) => event.action) ?? []),
    },
    startup_plan: includeExamples ? plan : undefined,
    readiness,
    issues: plan.issues ?? [],
    next_actions: unique([
      ...(plan.next_actions ?? []),
      readiness.local_surface_selected ? 'wire_recipe_sequence_into_host_runtime' : 'choose_browser_or_native_host_profile_before_realtime_marks',
      'replace_raw_signal_examples_with_real_host_samples',
      'keep_speaker_track_text_free_and_filtered',
    ]),
  });
}

export function buildMeetingPlatformAdapterRuntimeRecipeMatrix(input = {}, options = {}) {
  const platforms = selectedPlatforms(input, options);
  const recipes = platforms.map((platform) => buildMeetingPlatformAdapterRuntimeRecipe({
    ...inputForPlatform(platform, objectInput(input), options),
    platform,
  }, {
    ...options,
    platform,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_adapter_runtime_recipe_matrix',
    schema: MEETING_PLATFORM_ADAPTER_RUNTIME_RECIPE_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_RUNTIME_RECIPE_SCHEMA_VERSION,
    platform_count: recipes.length,
    accepted_count: recipes.filter((recipe) => recipe.accepted).length,
    runtime_ready_count: recipes.filter((recipe) => recipe.runtime_ready).length,
    browser_surface_count: recipes.filter((recipe) => recipe.selected_surface === 'browser_extension').length,
    native_surface_count: recipes.filter((recipe) => recipe.selected_surface === 'native_detector').length,
    provider_reconcile_surface_count: recipes.filter((recipe) => recipe.selected_surface === 'provider_reconcile').length,
    raw_signal_runtime_event_count: recipes.reduce((total, recipe) => total + (recipe.raw_signal_examples?.runtime_event_count ?? 0), 0),
    platforms: recipes.map((recipe) => recipe.platform).filter(Boolean),
    rows: recipes.map((recipe) => ({
      platform: recipe.platform,
      display_name: recipe.display_name,
      accepted: recipe.accepted,
      runtime_ready: recipe.runtime_ready,
      status: recipe.status,
      host_profile: recipe.host_profile,
      selected_surface: recipe.selected_surface,
      install_target: recipe.install_target,
      bridge_kind: recipe.host_wiring?.bridge_kind,
      first_required_method: recipe.host_wiring?.first_required_method,
      insert_method: recipe.host_wiring?.insert_method,
      provider_events_block_realtime: recipe.runtime_contract?.provider_events_block_realtime,
      transcript_blocks_realtime: recipe.runtime_contract?.transcript_blocks_realtime,
      raw_signal_runtime_event_count: recipe.raw_signal_examples?.runtime_event_count,
      speaker_track_sample_ready: recipe.readiness?.sample_speaker_track_ready === true,
      first_next_action: recipe.next_actions?.[0],
    })),
    recipes,
    next_actions: unique(recipes.flatMap((recipe) => recipe.next_actions ?? [])),
  };
}

function manifestHostKind(recipe = {}) {
  const bridge = recipe.host_wiring?.bridge_kind;
  if (bridge === 'browser_content_script') return 'browser_extension_content_script';
  if (bridge === 'webview_preload_bridge') return 'embedded_webview_preload';
  if (bridge === 'native_detector_runtime_event_client') return 'native_desktop_detector';
  if (bridge === 'provider_reconcile_only_not_realtime_axis') return 'provider_reconcile_only';
  return 'custom_host_surface_adapter';
}

function bridgeGroups(rows = []) {
  const groups = new Map();
  for (const row of rows) {
    const bridge = row.bridge_kind ?? 'unknown_bridge';
    const current = groups.get(bridge) ?? {
      bridge_kind: bridge,
      host_kind: manifestHostKind({ host_wiring: { bridge_kind: bridge } }),
      platform_count: 0,
      runtime_ready_count: 0,
      platforms: [],
    };
    current.platform_count += 1;
    if (row.runtime_ready === true) current.runtime_ready_count += 1;
    current.platforms.push(row.platform);
    groups.set(bridge, current);
  }
  return [...groups.values()];
}

function registryRow(recipe = {}) {
  const runtimeActions = recipe.host_wiring?.runtime_event_actions ?? recipe.raw_signal_examples?.runtime_actions ?? [];
  return compactObject({
    platform: recipe.platform,
    display_name: recipe.display_name,
    accepted: recipe.accepted,
    runtime_ready: recipe.runtime_ready,
    status: recipe.status,
    selected_surface: recipe.selected_surface,
    host_profile: recipe.host_profile,
    host_kind: manifestHostKind(recipe),
    bridge_kind: recipe.host_wiring?.bridge_kind,
    install_target: recipe.install_target,
    adapter_module: recipe.platform ? adapterModuleFor(recipe.platform) : undefined,
    first_required_method: recipe.host_wiring?.first_required_method,
    insert_method: recipe.host_wiring?.insert_method,
    optional_track_methods: recipe.host_wiring?.optional_track_methods,
    provider_method: recipe.host_wiring?.provider_method,
    timestamp_field: 'captured_at_ms',
    runtime_event_actions: runtimeActions,
    runtime_event_action_counts: recipe.host_wiring?.runtime_event_action_counts,
    sample_runtime_event_count: recipe.raw_signal_examples?.runtime_event_count,
    speaker_position_markers: {
      enabled: recipe.readiness?.sample_speaker_track_ready === true,
      text_required: false,
      filter_policy: 'debounce_active_speaker_samples_before_drawing_position_markers',
    },
    participant_position_markers: {
      enabled: runtimeActions.includes('participant_track'),
      text_required: false,
    },
    provider_reconcile: {
      method: recipe.host_wiring?.provider_method ?? 'ingestProvider',
      blocks_realtime_annotation: recipe.runtime_contract?.provider_events_block_realtime === true,
      role: 'reconcile_or_backfill_only',
    },
    transcript: {
      blocks_realtime_annotation: recipe.runtime_contract?.transcript_blocks_realtime === true,
    },
    first_next_action: recipe.next_actions?.[0],
  });
}

function compactRecipeForManifest(recipe = {}) {
  return compactObject({
    schema: recipe.schema,
    schema_version: recipe.schema_version,
    accepted: recipe.accepted,
    runtime_ready: recipe.runtime_ready,
    status: recipe.status,
    platform: recipe.platform,
    display_name: recipe.display_name,
    selected_surface: recipe.selected_surface,
    host_profile: recipe.host_profile,
    install_target: recipe.install_target,
    host_wiring: recipe.host_wiring,
    sequence: recipe.sequence,
    raw_signal_examples: {
      schema: recipe.raw_signal_examples?.schema,
      signal_count: recipe.raw_signal_examples?.signal_count,
      runtime_event_count: recipe.raw_signal_examples?.runtime_event_count,
      filtered_speaker_event_count: recipe.raw_signal_examples?.filtered_speaker_event_count,
      runtime_actions: recipe.raw_signal_examples?.runtime_actions
        ?? unique(recipe.raw_signal_examples?.runtime_events?.map((event) => event.action) ?? []),
    },
    readiness: recipe.readiness,
    issues: recipe.issues,
    next_actions: recipe.next_actions,
  });
}

export function buildMeetingPlatformAdapterRuntimeManifest(input = {}, options = {}) {
  const matrix = buildMeetingPlatformAdapterRuntimeRecipeMatrix(input, options);
  const rows = matrix.recipes.map((recipe) => registryRow(recipe));
  const includeRecipes = options.includeRecipes === true || options.include_recipes === true;
  const providerBlockingCount = rows.filter((row) => row.provider_reconcile?.blocks_realtime_annotation === true).length;
  const transcriptBlockingCount = rows.filter((row) => row.transcript?.blocks_realtime_annotation === true).length;
  return {
    type: 'meeting_platform_adapter_runtime_manifest',
    schema: MEETING_PLATFORM_ADAPTER_RUNTIME_MANIFEST_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_RUNTIME_RECIPE_SCHEMA_VERSION,
    accepted: matrix.platform_count > 0 && matrix.accepted_count === matrix.platform_count,
    runtime_ready: matrix.platform_count > 0 && matrix.runtime_ready_count === matrix.platform_count,
    platform_count: matrix.platform_count,
    accepted_count: matrix.accepted_count,
    runtime_ready_count: matrix.runtime_ready_count,
    local_surface_count: matrix.browser_surface_count + matrix.native_surface_count,
    browser_surface_count: matrix.browser_surface_count,
    native_surface_count: matrix.native_surface_count,
    provider_reconcile_surface_count: matrix.provider_reconcile_surface_count,
    raw_signal_runtime_event_count: matrix.raw_signal_runtime_event_count,
    base_url: firstNonEmpty(options.baseUrl, options.base_url),
    host_endpoints: {
      runtime_events: endpointUrl(MEETING_PLATFORM_RUNTIME_EVENT_ENDPOINT, options),
      insert_annotation: endpointUrl('/api/annotations', options),
      insert_annotations: endpointUrl('/api/annotations/batch', options),
    },
    runtime_contract: {
      timestamp_field: 'captured_at_ms',
      observe_before_insert: true,
      provider_events_block_realtime: providerBlockingCount > 0,
      transcript_blocks_realtime: transcriptBlockingCount > 0,
      speaker_text_required: false,
      participant_text_required: false,
      required_methods: [
        'observePlatformCandidates',
        'insertAnnotation',
      ],
      optional_methods: [
        'speakerTrack',
        'participantTrack',
        'ingestProvider',
      ],
    },
    dispatch_policy: {
      browser_content_script: 'use_for_google_meet_and_browser_meeting_surfaces',
      native_detector_runtime_event_client: 'use_for_teams_zoom_or_desktop_surfaces_when_browser_dom_is_not_the_primary_surface',
      webview_preload_bridge: 'use_for_embedded_meeting_webviews',
      provider_reconcile_only_not_realtime_axis: 'do_not_use_as_primary_realtime_axis',
    },
    bridge_groups: bridgeGroups(rows),
    platform_registry: {
      row_count: rows.length,
      rows,
    },
    recipes: includeRecipes ? matrix.recipes.map((recipe) => compactRecipeForManifest(recipe)) : undefined,
    matrix_summary: {
      schema: matrix.schema,
      platforms: matrix.platforms,
      next_actions: matrix.next_actions,
    },
    next_actions: unique([
      'install_host_runtime_for_each_bridge_kind',
      'route_google_meet_to_browser_content_script_when_available',
      'route_desktop_platforms_to_native_detector_when_selected_surface_is_native_detector',
      'send_captured_at_ms_with_every_observation_and_annotation',
      'filter_speaker_position_markers_before_drawing_timeline_positions',
      ...matrix.next_actions,
    ]),
  };
}

function runtimeManifestFrom(manifestOrInput = {}, input = {}, options = {}) {
  if (manifestOrInput?.schema === MEETING_PLATFORM_ADAPTER_RUNTIME_MANIFEST_SCHEMA) return manifestOrInput;
  return firstNonEmpty(
    options.runtimeManifest,
    options.runtime_manifest,
    input.runtimeManifest,
    input.runtime_manifest,
    manifestOrInput.runtimeManifest,
    manifestOrInput.runtime_manifest,
    manifestOrInput.manifest,
  ) ?? buildMeetingPlatformAdapterRuntimeManifest({}, {
    ...input,
    ...options,
    platforms: firstNonEmpty(options.platforms, options.platform_keys, input.platforms, input.platform_keys, manifestOrInput.platforms, manifestOrInput.platform_keys),
  });
}

function runtimeTargetInput(manifestOrInput = {}, input = {}) {
  return manifestOrInput?.schema === MEETING_PLATFORM_ADAPTER_RUNTIME_MANIFEST_SCHEMA ? input : manifestOrInput;
}

function rawTargetUrl(input = {}, options = {}) {
  if (typeof input === 'string' || input instanceof URL) return String(input);
  return firstNonEmpty(
    options.url,
    options.href,
    options.meetingUrl,
    options.meeting_url,
    firstPath(input, [
      'meeting.meeting_url',
      'meeting.meetingUrl',
      'meeting.url',
      'meeting.join_url',
      'meeting.joinUrl',
      'meeting_url',
      'meetingUrl',
      'join_url',
      'joinUrl',
      'url',
      'href',
      'window.url',
      'browser.url',
      'tab.url',
    ]),
  );
}

function explicitTargetPlatform(input = {}, options = {}) {
  if (typeof input === 'string' || input instanceof URL) return maybeNormalizePlatform(options.platform);
  return maybeNormalizePlatform(firstNonEmpty(
    options.platform,
    options.platform_key,
    input.platform,
    input.platform_key,
    input.provider,
    input.adapter,
    input.meeting?.platform,
    input.current_meeting?.platform,
    input.currentMeeting?.platform,
    input.detected_meeting?.platform,
    input.detectedMeeting?.platform,
    input.tab?.platform,
    input.window?.platform,
  ));
}

function resolveRuntimeTargetPlatform(manifest = {}, input = {}, options = {}) {
  const url = rawTargetUrl(input, options);
  const explicit = explicitTargetPlatform(input, options);
  const detected = detectMeetingFromUrl({
    url,
    title: typeof input === 'object' ? firstPath(input, ['title', 'tab.title', 'window.title', 'meeting.title']) : undefined,
  });
  const platform = firstNonEmpty(explicit, detected?.platform);
  return {
    platform,
    detected_meeting: detected,
    detection_reason: explicit ? 'explicit_platform' : detected ? 'meeting_url' : 'none',
    url,
    manifest_schema: manifest.schema,
  };
}

function runtimeRegistryRows(manifest = {}) {
  return asArray(manifest.platform_registry?.rows);
}

function runtimeRegistryRow(manifest = {}, platform) {
  return runtimeRegistryRows(manifest).find((row) => row.platform === platform);
}

function requiredSurface(input = {}, options = {}) {
  if (typeof input === 'string' || input instanceof URL) return normalizeSurface(options.surface ?? options.preferredSurface ?? options.preferred_surface);
  return normalizeSurface(firstNonEmpty(
    options.surface,
    options.preferredSurface,
    options.preferred_surface,
    input.surface,
    input.preferredSurface,
    input.preferred_surface,
  ));
}

function targetRuntimeActions(row = {}) {
  const localSurface = row.selected_surface !== 'provider_reconcile';
  return [
    localSurface ? {
      id: 'observe_platform_candidates',
      phase: 'local_axis',
      required: true,
      sdk_method: row.first_required_method ?? 'observePlatformCandidates',
      bridge_kind: row.bridge_kind,
      timestamp_field: 'captured_at_ms',
    } : undefined,
    localSurface ? {
      id: 'insert_realtime_annotation',
      phase: 'realtime_mark',
      required: true,
      sdk_method: row.insert_method ?? 'insertAnnotation',
      timestamp_field: 'captured_at_ms',
    } : undefined,
    row.speaker_position_markers?.enabled === true ? {
      id: 'emit_speaker_position',
      phase: 'optional_track',
      required: false,
      sdk_method: 'speakerTrack',
      text_required: false,
      filter_policy: row.speaker_position_markers.filter_policy,
      timestamp_field: 'captured_at_ms',
    } : undefined,
    row.participant_position_markers?.enabled === true ? {
      id: 'emit_participant_position',
      phase: 'optional_track',
      required: false,
      sdk_method: 'participantTrack',
      text_required: false,
      timestamp_field: 'captured_at_ms',
    } : undefined,
    {
      id: 'provider_reconcile_backfill',
      phase: 'post_axis_reconcile',
      required: false,
      sdk_method: row.provider_reconcile?.method ?? 'ingestProvider',
      realtime_blocking: false,
    },
  ].filter(Boolean);
}

function runtimeTargetReadiness(manifest = {}, row = null, resolved = {}, surface = undefined) {
  const required = surface;
  const issues = [
    manifest.schema === MEETING_PLATFORM_ADAPTER_RUNTIME_MANIFEST_SCHEMA ? undefined : issue('error', 'invalid_runtime_manifest', 'Runtime target requires a runtime manifest.', {
      actual: manifest.schema,
    }),
    manifest.runtime_ready === true ? undefined : issue('error', 'runtime_manifest_not_ready', 'Runtime manifest must be ready before selecting a runtime target.'),
    resolved.platform ? undefined : issue('error', 'platform_not_detected', 'No supported meeting platform was detected from the current input.'),
    row?.platform ? undefined : issue('error', 'platform_not_registered', 'Detected meeting platform is not present in the runtime manifest.', {
      platform: resolved.platform,
    }),
    row?.runtime_ready === true ? undefined : issue('error', 'runtime_target_not_ready', 'Runtime manifest row is not ready.', {
      platform: resolved.platform,
    }),
    required && row?.selected_surface && required !== row.selected_surface ? issue('error', 'surface_mismatch', 'Requested surface does not match the runtime manifest target surface.', {
      platform: resolved.platform,
      requested_surface: required,
      selected_surface: row.selected_surface,
    }) : undefined,
    row?.selected_surface === 'provider_reconcile' ? issue('error', 'provider_reconcile_not_realtime_target', 'Provider reconcile cannot be the primary realtime runtime target.', {
      platform: resolved.platform,
    }) : undefined,
  ].filter(Boolean);
  return {
    accepted: issues.filter((item) => item.severity === 'error').length === 0,
    issue_count: issues.length,
    issues,
  };
}

function runtimeTargetStatus(ready = {}, row = null, resolved = {}) {
  if (ready.accepted === true) return 'ready_to_start_runtime_target';
  if (!resolved.platform) return 'platform_not_detected';
  if (!row?.platform) return 'platform_not_registered';
  if (row?.selected_surface === 'provider_reconcile') return 'needs_local_surface_for_realtime_axis';
  return 'runtime_target_not_ready';
}

function runtimeTargetNextActions(target = {}, ready = {}) {
  if (ready.accepted !== true) return unique([
    ...(ready.issues ?? []).map((item) => item.code),
    'fix_runtime_target_blockers',
  ]);
  return unique([
    `start_${target.host_kind}`,
    'observe_platform_candidates_before_first_mark',
    'insert_realtime_marks_with_captured_at_ms',
    'filter_speaker_position_markers_before_drawing_timeline_positions',
    'reconcile_provider_events_after_local_axis',
  ]);
}

export function buildMeetingPlatformAdapterRuntimeTarget(manifestOrInput = {}, input = {}, options = {}) {
  const manifest = runtimeManifestFrom(manifestOrInput, input, options);
  const targetInput = runtimeTargetInput(manifestOrInput, input);
  const resolved = resolveRuntimeTargetPlatform(manifest, targetInput, options);
  const row = runtimeRegistryRow(manifest, resolved.platform);
  const surface = requiredSurface(targetInput, options);
  const ready = runtimeTargetReadiness(manifest, row, resolved, surface);
  const target = compactObject({
    type: 'meeting_platform_adapter_runtime_target',
    schema: MEETING_PLATFORM_ADAPTER_RUNTIME_TARGET_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_RUNTIME_RECIPE_SCHEMA_VERSION,
    accepted: ready.accepted,
    status: runtimeTargetStatus(ready, row, resolved),
    platform: resolved.platform,
    detection_reason: resolved.detection_reason,
    detected_meeting: resolved.detected_meeting,
    current_url: resolved.url,
    selected_surface: row?.selected_surface,
    host_kind: row?.host_kind,
    bridge_kind: row?.bridge_kind,
    adapter_module: row?.adapter_module,
    first_required_method: row?.first_required_method,
    insert_method: row?.insert_method,
    timestamp_field: 'captured_at_ms',
    runtime_event_actions: row?.runtime_event_actions,
    host_endpoints: manifest.host_endpoints,
    runtime_contract: manifest.runtime_contract,
    dispatch_policy: row?.bridge_kind ? manifest.dispatch_policy?.[row.bridge_kind] : undefined,
    platform_registry_row: row,
    runtime_actions: row ? targetRuntimeActions(row) : [],
    mark_template: row ? {
      platform: row.platform,
      captured_at_ms: firstNonEmpty(
        options.capturedAtMs,
        options.captured_at_ms,
        targetInput.capturedAtMs,
        targetInput.captured_at_ms,
        0,
      ),
      source: 'meeting_platform_adapter_runtime_target',
    } : undefined,
    readiness: ready,
  });
  return {
    ...target,
    next_actions: runtimeTargetNextActions(target, ready),
  };
}

export function assertMeetingPlatformAdapterRuntimeRecipe(input = {}, options = {}) {
  const recipe = buildMeetingPlatformAdapterRuntimeRecipe(input, options);
  if (recipe.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting platform adapter runtime recipe is not accepted', {
      platform: recipe.platform,
      status: recipe.status,
      issues: recipe.issues,
      next_actions: recipe.next_actions,
      recipe,
    });
  }
  return recipe;
}

export function assertMeetingPlatformAdapterRuntimeRecipeMatrix(input = {}, options = {}) {
  const matrix = buildMeetingPlatformAdapterRuntimeRecipeMatrix(input, options);
  if (matrix.accepted_count !== matrix.platform_count) {
    throw new MeetingTimelineSdkError('Meeting platform adapter runtime recipe matrix is not accepted', {
      platform_count: matrix.platform_count,
      accepted_count: matrix.accepted_count,
      failed_platforms: matrix.rows.filter((row) => row.accepted !== true).map((row) => row.platform),
      next_actions: matrix.next_actions,
      matrix,
    });
  }
  return matrix;
}

export function assertMeetingPlatformAdapterRuntimeManifest(input = {}, options = {}) {
  const manifest = buildMeetingPlatformAdapterRuntimeManifest(input, options);
  if (manifest.accepted !== true || manifest.runtime_ready !== true) {
    throw new MeetingTimelineSdkError('Meeting platform adapter runtime manifest is not ready', {
      platform_count: manifest.platform_count,
      accepted_count: manifest.accepted_count,
      runtime_ready_count: manifest.runtime_ready_count,
      next_actions: manifest.next_actions,
      manifest,
    });
  }
  return manifest;
}

export function assertMeetingPlatformAdapterRuntimeTarget(manifestOrInput = {}, input = {}, options = {}) {
  const target = manifestOrInput.schema === MEETING_PLATFORM_ADAPTER_RUNTIME_TARGET_SCHEMA
    ? manifestOrInput
    : buildMeetingPlatformAdapterRuntimeTarget(manifestOrInput, input, options);
  if (target.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting platform adapter runtime target is not ready', {
      code: 'meeting_platform_adapter_runtime_target_not_ready',
      status: target.status,
      issues: target.readiness?.issues ?? [],
      platform: target.platform,
      selected_surface: target.selected_surface,
    });
  }
  return target;
}
