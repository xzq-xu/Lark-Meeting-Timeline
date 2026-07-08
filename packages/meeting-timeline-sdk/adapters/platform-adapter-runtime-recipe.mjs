import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  buildMeetingPlatformAdapterStartupPlan,
} from './platform-adapter-startup.mjs';
import {
  buildMeetingPlatformRawSignalExampleBatch,
} from './platform-raw-signal.mjs';
import {
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_ADAPTER_RUNTIME_RECIPE_SCHEMA = 'meeting_platform_adapter_runtime_recipe';
export const MEETING_PLATFORM_ADAPTER_RUNTIME_RECIPE_MATRIX_SCHEMA = 'meeting_platform_adapter_runtime_recipe_matrix';
export const MEETING_PLATFORM_ADAPTER_RUNTIME_RECIPE_SCHEMA_VERSION = 1;

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
