import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  MEETING_APP_EXTENSION_MESSAGE_TYPES,
  MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS,
} from './meeting-app-extension.mjs';
import {
  MEETING_APP_ADAPTER_SPEC_SCHEMA,
  buildMeetingAppAdapterSpec,
} from './meeting-app-adapter-spec.mjs';

export const MEETING_APP_ADAPTER_RUNTIME_CONFIG_SCHEMA = 'meeting_app_adapter_runtime_config';
export const MEETING_APP_ADAPTER_RUNTIME_CONFIG_MATRIX_SCHEMA = 'meeting_app_adapter_runtime_config_matrix';
export const MEETING_APP_ADAPTER_RUNTIME_CONFIG_SCHEMA_VERSION = 1;

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

function issue(severity, code, message, details = {}) {
  return compactObject({ severity, code, message, ...details });
}

function specFromInput(specOrPlatform = {}, options = {}) {
  return specOrPlatform?.schema === MEETING_APP_ADAPTER_SPEC_SCHEMA
    ? specOrPlatform
    : buildMeetingAppAdapterSpec(specOrPlatform, options);
}

function buildCaptureOptions(spec = {}, options = {}) {
  const capture = spec.capture ?? {};
  return {
    source: firstNonEmpty(options.source, 'meeting_app_adapter_runtime_config'),
    platform: spec.adapter_key,
    captureProfile: false,
    capture_profile: false,
    controlSelectors: capture.control_selectors ?? [],
    control_selectors: capture.control_selectors ?? [],
    participantSelectors: capture.participant_selectors ?? [],
    participant_selectors: capture.participant_selectors ?? [],
    textSelectors: capture.text_selectors ?? [],
    text_selectors: capture.text_selectors ?? [],
    maxControls: firstNonEmpty(options.maxControls, options.max_controls),
    maxParticipants: firstNonEmpty(options.maxParticipants, options.max_participants),
    maxTexts: firstNonEmpty(options.maxTexts, options.max_texts),
    includeShadowDom: firstNonEmpty(options.includeShadowDom, options.include_shadow_dom, false),
    include_shadow_dom: firstNonEmpty(options.include_shadow_dom, options.includeShadowDom, false),
  };
}

function buildBrowserRuntimeOptions(spec = {}, options = {}) {
  const runtime = spec.runtime ?? {};
  const captureOptions = buildCaptureOptions(spec, options);
  return {
    runtimePreset: false,
    runtime_preset: false,
    adapterKey: spec.adapter_key,
    adapter_key: spec.adapter_key,
    source: firstNonEmpty(options.source, 'meeting_app_adapter_runtime_config'),
    observeMutations: runtime.observe_mutations === true,
    observe_mutations: runtime.observe_mutations === true,
    mutationDebounceMs: runtime.mutation_debounce_ms,
    mutation_debounce_ms: runtime.mutation_debounce_ms,
    speakerStableFollowupMs: runtime.speaker_stable_followup_ms,
    speaker_stable_followup_ms: runtime.speaker_stable_followup_ms,
    sampleIntervalMs: runtime.sample_interval_ms,
    sample_interval_ms: runtime.sample_interval_ms,
    unchangedObserveEveryMs: runtime.unchanged_observe_every_ms,
    unchanged_observe_every_ms: runtime.unchanged_observe_every_ms,
    mutationTrackSelectors: runtime.mutation_track_selectors ?? [],
    mutation_track_selectors: runtime.mutation_track_selectors ?? [],
    mutationIgnoreSelectors: runtime.mutation_ignore_selectors ?? [],
    mutation_ignore_selectors: runtime.mutation_ignore_selectors ?? [],
    captureOptions,
    capture_options: captureOptions,
    applyOptions: {
      speakerAsAnnotation: true,
      speaker_as_annotation: true,
      ...(options.applyOptions ?? {}),
      ...(options.apply_options ?? {}),
    },
    speakerOptions: {
      minStableMs: runtime.speaker_stable_followup_ms,
      min_stable_ms: runtime.speaker_stable_followup_ms,
      ...(options.speakerOptions ?? {}),
      ...(options.speaker_options ?? {}),
    },
    participantOptions: {
      ...(options.participantOptions ?? {}),
      ...(options.participant_options ?? {}),
    },
  };
}

function buildContentScriptOptions(spec = {}, options = {}) {
  const browserRuntimeOptions = buildBrowserRuntimeOptions(spec, options);
  return {
    ...browserRuntimeOptions,
    extensionMessaging: firstNonEmpty(options.extensionMessaging, options.extension_messaging, true),
    extension_messaging: firstNonEmpty(options.extension_messaging, options.extensionMessaging, true),
    windowMessaging: firstNonEmpty(options.windowMessaging, options.window_messaging, false),
    window_messaging: firstNonEmpty(options.window_messaging, options.windowMessaging, false),
    allowedOrigins: unique(asArray(options.allowedOrigins ?? options.allowed_origins)),
    allowed_origins: unique(asArray(options.allowed_origins ?? options.allowedOrigins)),
    messagePrefixes: ['meeting_timeline', 'meeting-timeline'],
    message_prefixes: ['meeting_timeline', 'meeting-timeline'],
  };
}

function buildExtensionManifestFragment(spec = {}) {
  return {
    matches: spec.extension?.matches ?? [],
    host_permissions: spec.extension?.host_permissions ?? [],
    permissions: spec.extension?.permissions ?? [],
    content_scripts: [{
      matches: spec.extension?.matches ?? [],
      js: ['content-script.js'],
      run_at: 'document_idle',
    }],
    message_types: spec.extension?.message_types,
  };
}

function runtimeIssues(spec = {}, config = {}) {
  return [
    spec.accepted === true
      ? undefined
      : issue('error', 'adapter_spec_not_accepted', 'Runtime config requires an accepted adapter spec.', {
        adapter_key: spec.adapter_key,
        spec_issue_count: spec.issue_count,
      }),
    config.browser_runtime_options?.runtimePreset === false
      ? undefined
      : issue('error', 'runtime_preset_must_be_disabled_for_spec', 'Custom adapter runtime config must disable built-in runtimePreset.'),
    config.capture_options?.controlSelectors?.length > 0
      ? undefined
      : issue('error', 'missing_control_capture_options', 'Runtime capture options require control selectors.'),
    config.capture_options?.participantSelectors?.length > 0
      ? undefined
      : issue('error', 'missing_participant_capture_options', 'Runtime capture options require participant selectors.'),
    config.browser_runtime_options?.observeMutations === true
      ? undefined
      : issue('error', 'mutation_observation_disabled', 'Runtime config must enable mutation observation.'),
    asArray(config.browser_runtime_options?.mutationTrackSelectors).length > 0
      ? undefined
      : issue('error', 'missing_mutation_track_selectors', 'Runtime config requires mutation track selectors.'),
    config.contracts?.timestamp_field === 'captured_at_ms'
      ? undefined
      : issue('error', 'wrong_timestamp_field', 'Runtime config must preserve captured_at_ms.'),
    config.contracts?.provider_events_block_realtime === false
      ? undefined
      : issue('error', 'provider_blocks_realtime', 'Provider events must not block realtime annotation placement.'),
    config.contracts?.transcript_blocks_realtime === false
      ? undefined
      : issue('error', 'transcript_blocks_realtime', 'Transcript import must not block realtime annotation placement.'),
  ].filter(Boolean);
}

export function buildMeetingAppAdapterRuntimeConfig(specOrPlatform = {}, options = {}) {
  const spec = specFromInput(specOrPlatform, options);
  const captureOptions = buildCaptureOptions(spec, options);
  const browserRuntimeOptions = buildBrowserRuntimeOptions(spec, options);
  const contentScriptOptions = buildContentScriptOptions(spec, options);
  const config = {
    type: 'meeting_app_adapter_runtime_config',
    schema: MEETING_APP_ADAPTER_RUNTIME_CONFIG_SCHEMA,
    schema_version: MEETING_APP_ADAPTER_RUNTIME_CONFIG_SCHEMA_VERSION,
    accepted: false,
    adapter_key: spec.adapter_key,
    display_name: spec.display_name,
    source_spec: {
      schema: spec.schema,
      adapter_key: spec.adapter_key,
      source: spec.source,
      accepted: spec.accepted,
      issue_count: spec.issue_count,
    },
    entrypoints: {
      browser_runtime_factory: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-browser-runtime#createMeetingAppBrowserRuntime',
      content_script_bridge_factory: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-content-script#installMeetingAppContentScriptBridge',
      capture_function: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-capture#captureMeetingAppDomSnapshot',
      track_runtime_factory: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-track-runtime#createMeetingAppTrackRuntime',
    },
    extension: buildExtensionManifestFragment(spec),
    capture_options: captureOptions,
    browser_runtime_options: browserRuntimeOptions,
    content_script_options: contentScriptOptions,
    host_endpoints: spec.host_endpoints ?? MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS,
    contracts: {
      timestamp_field: 'captured_at_ms',
      candidate_observation_message_type: MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      local_observer_may_start_axis: true,
      required_signals: spec.contracts?.required_signals ?? [],
      output_intents: spec.contracts?.output_intents ?? ['realtime_axis', 'speaker_track', 'participant_track', 'annotation_insert'],
    },
    usage: {
      create_browser_runtime: 'createMeetingAppBrowserRuntime(timelineClient, config.browser_runtime_options)',
      install_content_script_bridge: 'installMeetingAppContentScriptBridge(timelineClient, config.content_script_options)',
      capture_snapshot: 'captureMeetingAppDomSnapshot({ document, window, location }, config.capture_options)',
    },
    readiness: {
      adapter_spec_accepted: spec.accepted === true,
      capture_options_ready: asArray(captureOptions.controlSelectors).length > 0 && asArray(captureOptions.participantSelectors).length > 0,
      mutation_observer_ready: browserRuntimeOptions.observeMutations === true && asArray(browserRuntimeOptions.mutationTrackSelectors).length > 0,
      content_script_ready: asArray(spec.extension?.matches).length > 0,
      timestamp_ready: true,
      provider_nonblocking: true,
      transcript_nonblocking: true,
      requires_live_snapshot_before_production: true,
    },
  };
  const issues = runtimeIssues(spec, config);
  const blocking = issues.filter((item) => item.severity === 'error');
  return {
    ...config,
    accepted: blocking.length === 0,
    issue_count: issues.length,
    blocking_count: blocking.length,
    warning_count: issues.filter((item) => item.severity === 'warning').length,
    issues,
    next_actions: unique([
      ...(blocking.map((item) => item.code)),
      'capture_live_dom_snapshots_before_production_rollout',
    ]),
  };
}

function configInputs(options = {}) {
  const specs = asArray(firstNonEmpty(options.configs, options.adapters, options.adapter_specs, options.adapterSpecs, options.specs));
  const platforms = asArray(firstNonEmpty(options.platforms, options.platform_keys, options.platformKeys, []));
  return [...platforms, ...specs];
}

export function buildMeetingAppAdapterRuntimeConfigMatrix(options = {}) {
  const configs = configInputs(options).map((input) => buildMeetingAppAdapterRuntimeConfig(input, {
    ...options,
    configs: undefined,
    adapters: undefined,
    adapter_specs: undefined,
    adapterSpecs: undefined,
    specs: undefined,
    platforms: undefined,
    platform_keys: undefined,
    platformKeys: undefined,
  }));
  return {
    type: 'meeting_app_adapter_runtime_config_matrix',
    schema: MEETING_APP_ADAPTER_RUNTIME_CONFIG_MATRIX_SCHEMA,
    schema_version: MEETING_APP_ADAPTER_RUNTIME_CONFIG_SCHEMA_VERSION,
    accepted: configs.every((config) => config.accepted === true),
    config_count: configs.length,
    accepted_count: configs.filter((config) => config.accepted === true).length,
    custom_count: configs.filter((config) => config.source_spec.source !== 'built_in_manifest').length,
    built_in_count: configs.filter((config) => config.source_spec.source === 'built_in_manifest').length,
    capture_ready_count: configs.filter((config) => config.readiness.capture_options_ready).length,
    mutation_ready_count: configs.filter((config) => config.readiness.mutation_observer_ready).length,
    content_script_ready_count: configs.filter((config) => config.readiness.content_script_ready).length,
    rows: configs.map((config) => ({
      adapter_key: config.adapter_key,
      display_name: config.display_name,
      source: config.source_spec.source,
      accepted: config.accepted,
      capture_options_ready: config.readiness.capture_options_ready,
      mutation_observer_ready: config.readiness.mutation_observer_ready,
      content_script_ready: config.readiness.content_script_ready,
      control_selector_count: asArray(config.capture_options.controlSelectors).length,
      participant_selector_count: asArray(config.capture_options.participantSelectors).length,
      mutation_track_selector_count: asArray(config.browser_runtime_options.mutationTrackSelectors).length,
      blocking_count: config.blocking_count,
      warning_count: config.warning_count,
      first_next_action: config.next_actions?.[0],
    })),
    configs,
    next_actions: unique(configs.flatMap((config) => config.next_actions ?? [])),
  };
}

export function assertMeetingAppAdapterRuntimeConfig(configOrSpec = {}, options = {}) {
  const config = configOrSpec?.schema === MEETING_APP_ADAPTER_RUNTIME_CONFIG_SCHEMA
    ? configOrSpec
    : buildMeetingAppAdapterRuntimeConfig(configOrSpec, options);
  if (config.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting app adapter runtime config acceptance failed', {
      code: 'meeting_app_adapter_runtime_config_rejected',
      adapter_key: config.adapter_key,
      issues: config.issues,
      next_actions: config.next_actions,
    });
  }
  return config;
}

export function assertMeetingAppAdapterRuntimeConfigMatrix(matrixOrOptions = {}, options = {}) {
  const matrix = matrixOrOptions?.schema === MEETING_APP_ADAPTER_RUNTIME_CONFIG_MATRIX_SCHEMA
    ? matrixOrOptions
    : buildMeetingAppAdapterRuntimeConfigMatrix({ ...matrixOrOptions, ...options });
  if (matrix.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting app adapter runtime config matrix acceptance failed', {
      code: 'meeting_app_adapter_runtime_config_matrix_rejected',
      rows: matrix.rows,
      next_actions: matrix.next_actions,
    });
  }
  return matrix;
}
