import { compactObject } from '../index.mjs';
import {
  MEETING_PLATFORM_KEYS,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';
import {
  buildMeetingPlatformAdaptationPackage,
} from './platform-adaptation-package.mjs';
import {
  MEETING_APP_BROWSER_RUNTIME_PRESETS,
  meetingAppBrowserRuntimePreset,
} from './meeting-app-browser-runtime.mjs';
import {
  MEETING_APP_EXTENSION_MESSAGE_TYPES,
  buildMeetingAppContentScriptManifest,
  buildMeetingAppExtensionAttachedMessage,
  buildMeetingAppExtensionClientCallMessage,
  buildMeetingAppExtensionObserveCandidatesMessage,
  buildMeetingAppExtensionStatusMessage,
} from './meeting-app-extension.mjs';
import {
  buildMeetingAppRuntimeObserverPlan,
} from './meeting-app-profile.mjs';
import {
  MEETING_PLATFORM_RUNTIME_EVENT_ENDPOINT,
  MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA,
  buildMeetingPlatformRuntimeEventPlan,
} from './platform-runtime-event.mjs';

export const MEETING_PLATFORM_RUNTIME_BUNDLE_SCHEMA = 'meeting_platform_runtime_bundle';
export const MEETING_PLATFORM_RUNTIME_BUNDLE_MATRIX_SCHEMA = 'meeting_platform_runtime_bundle_matrix';
export const MEETING_PLATFORM_RUNTIME_BUNDLE_SCHEMA_VERSION = 1;

const DEFAULT_RUNTIME_BUNDLE_PLATFORMS = MEETING_PLATFORM_KEYS.filter((platform) => platform !== 'local_detector');

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
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, DEFAULT_RUNTIME_BUNDLE_PLATFORMS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

function bundleId(platform, options = {}) {
  return String(firstNonEmpty(
    options.bundleId,
    options.bundle_id,
    `${platform}-meeting-runtime-bundle`,
  ));
}

function runtimePresetFor(platform, options = {}) {
  return meetingAppBrowserRuntimePreset({
    ...options,
    platform,
  }) ?? MEETING_APP_BROWSER_RUNTIME_PRESETS[platform] ?? null;
}

function contentScriptJs(options = {}) {
  const js = firstNonEmpty(
    options.contentScriptJs,
    options.content_script_js,
    options.js,
    'meeting-app-content-script.bundle.js',
  );
  return Array.isArray(js) ? js : [js];
}

function messageExamples(platform, options = {}) {
  const capturedAtMs = firstNonEmpty(options.capturedAtMs, options.captured_at_ms, 1_782_614_400_000);
  const url = firstNonEmpty(options.url, options.href, options.meeting_url, options.meetingUrl);
  const extensionPlatform = platform === 'local_detector' ? undefined : platform;
  return {
    attached: buildMeetingAppExtensionAttachedMessage({
      platform: extensionPlatform,
      capturedAtMs,
      url,
    }),
    status: buildMeetingAppExtensionStatusMessage({
      requestId: 'status-001',
      platform: extensionPlatform,
      capturedAtMs,
      url,
    }),
    observe_candidates: buildMeetingAppExtensionObserveCandidatesMessage({
      requestId: 'observe-001',
      capturedAtMs,
      tabs: [{
        url,
        title: `${extensionPlatform ?? 'local'} meeting`,
        active: true,
      }],
    }),
    insert_annotation: buildMeetingAppExtensionClientCallMessage('insertMark', {
      id: 'note-001',
      label: 'why?',
      captured_at_ms: capturedAtMs,
      kind: 'question',
      source: 'meeting_app_runtime_bundle',
    }, {
      platform: extensionPlatform,
      capturedAtMs,
      url,
    }),
    content_script_sample: {
      type: 'meeting_timeline.sample',
      payload: compactObject({
        platform: extensionPlatform,
        captured_at_ms: capturedAtMs,
        url,
      }),
    },
    content_script_insert_annotation: {
      type: 'meeting_timeline.insert_mark',
      payload: {
        mark: {
          id: 'note-001',
          label: 'why?',
          captured_at_ms: capturedAtMs,
          kind: 'question',
          source: 'meeting_app_runtime_bundle',
        },
      },
    },
  };
}

function startOptions(platform, adaptationPackage = {}, preset = {}, options = {}) {
  const observer = adaptationPackage.local_observer ?? {};
  return compactObject({
    platform,
    runtimePreset: platform,
    browserRuntimePreset: platform,
    sampleIntervalMs: firstNonEmpty(options.sampleIntervalMs, options.sample_interval_ms, preset.sampleIntervalMs),
    mutationDebounceMs: firstNonEmpty(options.mutationDebounceMs, options.mutation_debounce_ms, preset.mutationDebounceMs),
    speakerStableFollowupMs: firstNonEmpty(
      options.speakerStableFollowupMs,
      options.speaker_stable_followup_ms,
      preset.speakerStableFollowupMs,
    ),
    observeMutations: firstNonEmpty(options.observeMutations, options.observe_mutations, preset.observeMutations, true),
    mutationTrackSelectors: preset.mutationTrackSelectors ?? [],
    mutationIgnoreSelectors: preset.mutationIgnoreSelectors ?? [],
    captureOptions: {
      platform,
      captureProfile: firstNonEmpty(observer.capture_profile, platform),
      source: 'meeting_app_runtime_bundle',
    },
    observeOptions: {
      source: 'meeting_app_runtime_bundle',
      requireMeetingEnd: firstNonEmpty(options.requireMeetingEnd, options.require_meeting_end, false),
    },
  });
}

function endpointUrl(path, options = {}) {
  const baseUrl = firstNonEmpty(options.baseUrl, options.base_url);
  return baseUrl ? new URL(path, baseUrl).toString() : path;
}

function hostEndpoints(adaptationPackage = {}, options = {}) {
  const endpoints = adaptationPackage.collector?.timeline_ingest?.endpoints ?? {};
  return compactObject({
    startMeeting: endpoints.startMeeting,
    endMeeting: endpoints.endMeeting,
    insertMark: adaptationPackage.annotation_pipeline?.insert_endpoint ?? endpoints.insertMark,
    insertMarks: endpoints.insertMarks,
    runtimeEvents: endpointUrl(MEETING_PLATFORM_RUNTIME_EVENT_ENDPOINT, options),
    importTranscript: adaptationPackage.transcript?.import_endpoint ?? endpoints.importTranscript,
  });
}

function contentScriptBridgeOptions(platform, start = {}, options = {}) {
  return compactObject({
    baseUrl: firstNonEmpty(options.baseUrl, options.base_url),
    platforms: platform === 'local_detector' ? [] : [platform],
    extensionMessaging: firstNonEmpty(options.extensionMessaging, options.extension_messaging, true),
    windowMessaging: firstNonEmpty(options.windowMessaging, options.window_messaging, false),
    startOptions: start,
  });
}

function contentScriptManifest(platform, options = {}, js = []) {
  try {
    return buildMeetingAppContentScriptManifest({
      ...options,
      platforms: [platform],
      js,
    });
  } catch (error) {
    if (platform !== 'local_detector') throw error;
    return {
      manifest_version: firstNonEmpty(options.manifestVersion, options.manifest_version, 3),
      name: firstNonEmpty(options.name, 'Meeting Timeline Local Detector Bridge'),
      version: firstNonEmpty(options.version, '0.1.0'),
      description: 'Local detector runtime does not require a browser content-script match.',
      host_permissions: [],
      content_scripts: [],
    };
  }
}

function runtimeObserverPlan(platform, options = {}) {
  if (platform === 'local_detector') return null;
  return buildMeetingAppRuntimeObserverPlan({ platform }, {
    ...options,
    platform,
    surface: firstNonEmpty(options.surface, options.targetSurface, options.target_surface, 'browser-extension'),
  });
}

export function buildMeetingPlatformRuntimeBundle(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const adaptationPackage = buildMeetingPlatformAdaptationPackage(key, options);
  const preset = runtimePresetFor(key, options);
  const js = contentScriptJs(options);
  const extension = adaptationPackage.extension ?? {};
  const manifest = contentScriptManifest(key, options, js);
  const start = startOptions(key, adaptationPackage, preset ?? {}, options);
  const endpoints = hostEndpoints(adaptationPackage, options);
  const runtimeEventPlan = buildMeetingPlatformRuntimeEventPlan(key, options);
  const observerPlan = runtimeObserverPlan(key, options);
  const messages = messageExamples(key, {
    ...options,
    url: firstNonEmpty(options.url, options.href, extension.matches?.[0]?.replace('*', '')),
  });

  return compactObject({
    type: 'meeting_platform_runtime_bundle',
    schema: MEETING_PLATFORM_RUNTIME_BUNDLE_SCHEMA,
    schema_version: MEETING_PLATFORM_RUNTIME_BUNDLE_SCHEMA_VERSION,
    id: bundleId(key, options),
    platform: key,
    display_name: adaptationPackage.display_name,
    objective: 'browser_or_native_host_runtime_configuration_for_realtime_meeting_timeline_annotations',
    runtime_contract: adaptationPackage.runtime_contract,
    modules: {
      platform_integration_runtime: '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime',
      content_script_bridge: '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime',
      runtime_event: '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-event',
      content_script: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-content-script',
      browser_runtime: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-browser-runtime',
      timeline_runtime: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-runtime',
      live_adapter: '@ai-annotation/meeting-timeline-sdk/adapters/platform-live-adapter',
      adaptation_package: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adaptation-package',
      observer_plan: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-profile',
    },
    browser: {
      matches: extension.matches ?? [],
      host_permissions: extension.host_permissions ?? [],
      permissions: extension.permissions ?? adaptationPackage.runtime_profile?.extension?.recommended_permissions ?? [],
      content_scripts: extension.content_scripts ?? [],
      manifest,
      candidate_observation: adaptationPackage.candidate_observation,
    },
    runtime: {
      preset: key,
      start_options: start,
      mutation_observer: {
        enabled: start.observeMutations === true,
        debounce_ms: start.mutationDebounceMs,
        track_selectors: start.mutationTrackSelectors ?? [],
        ignore_selectors: start.mutationIgnoreSelectors ?? [],
      },
      capture_profile: start.captureOptions?.captureProfile,
      required_snapshots: adaptationPackage.local_observer?.required_snapshots ?? [],
      speaker_filter: adaptationPackage.speaker_markers?.filter,
      content_script_bridge: {
        module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime',
        create_function: 'createMeetingPlatformIntegrationContentScriptBridge',
        install_function: 'installMeetingPlatformIntegrationContentScriptBridge',
        options: contentScriptBridgeOptions(key, start, options),
      },
      observer_plan: observerPlan,
      observation_loop: observerPlan ? {
        factory: observerPlan.observer_runtime?.factory,
        surface: observerPlan.surface,
        timestamp_field: observerPlan.signal_contract?.timestamp_field,
        cadence: observerPlan.cadence,
        trigger_policy: observerPlan.trigger_policy,
      } : null,
    },
    messaging: {
      message_types: MEETING_APP_EXTENSION_MESSAGE_TYPES,
      bridge_message_types: [
        'meeting_timeline.sample',
        'meeting_timeline.insert_mark',
        'meeting_timeline.insert_marks',
        'meeting_timeline.provider_event',
      ],
      background_message_types: [
        MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
      ],
      candidate_observation: adaptationPackage.candidate_observation,
      accepted_methods: Object.keys(endpoints).filter((method) => endpoints[method]),
      runtime_event: {
        schema: MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA,
        endpoint: endpoints.runtimeEvents,
        build_function: 'buildMeetingPlatformRuntimeEvent',
        client_factory: 'createMeetingPlatformRuntimeEventClient',
        plan_schema: runtimeEventPlan.schema,
        plan: runtimeEventPlan,
      },
      examples: messages,
    },
    host: {
      endpoints,
      annotation_timestamp_field: 'captured_at_ms',
      insert_policy: adaptationPackage.annotation_pipeline?.realtime_policy,
    },
    provider_reconcile: adaptationPackage.provider_observer,
    transcript: adaptationPackage.transcript,
    readiness: {
      sdk_wiring_ready: adaptationPackage.readiness?.sdk_wiring_ready === true,
      runtime_ready: (extension.matches ?? []).length > 0 || key === 'local_detector',
      observer_plan_ready: observerPlan?.sdk_ready === true,
      observer_preflight_status: observerPlan?.preflight_status,
      provider_required_for_realtime: adaptationPackage.provider_observer?.required_for_realtime === true,
      transcript_blocks_realtime: adaptationPackage.transcript?.blocks_realtime_annotation === true,
      missing_items: adaptationPackage.readiness?.missing_items ?? [],
    },
    observer_plan: observerPlan,
    adaptation_package: adaptationPackage,
    next_actions: unique([
      ...(observerPlan?.next_actions ?? []),
      'install_content_script_or_native_preload',
      'install_meeting_platform_integration_content_script_bridge',
      'start_meeting_app_content_script_bridge',
      'send_captured_at_ms_with_every_annotation',
      ...(adaptationPackage.next_actions ?? []),
    ]),
  });
}

export function buildMeetingPlatformRuntimeBundleMatrix(options = {}) {
  const bundles = selectedPlatforms(options).map((platform) => buildMeetingPlatformRuntimeBundle(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_runtime_bundle_matrix',
    schema: MEETING_PLATFORM_RUNTIME_BUNDLE_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_RUNTIME_BUNDLE_SCHEMA_VERSION,
    platform_count: bundles.length,
    runtime_ready_count: bundles.filter((bundle) => bundle.readiness.runtime_ready).length,
    sdk_wiring_ready_count: bundles.filter((bundle) => bundle.readiness.sdk_wiring_ready).length,
    observer_plan_ready_count: bundles.filter((bundle) => bundle.readiness.observer_plan_ready).length,
    candidate_observer_count: bundles.filter((bundle) => bundle.messaging?.candidate_observation?.runtime_event_action === 'observe_platform_candidates').length,
    provider_required_for_realtime_count: bundles.filter((bundle) => bundle.readiness.provider_required_for_realtime).length,
    transcript_blocking_count: bundles.filter((bundle) => bundle.readiness.transcript_blocks_realtime).length,
    platforms: bundles.map((bundle) => bundle.platform),
    rows: bundles.map((bundle) => ({
      platform: bundle.platform,
      display_name: bundle.display_name,
      runtime_ready: bundle.readiness.runtime_ready,
      sdk_wiring_ready: bundle.readiness.sdk_wiring_ready,
      browser_match_count: bundle.browser?.matches?.length ?? 0,
      browser_permission_count: bundle.browser?.permissions?.length ?? 0,
      candidate_observation_ready: bundle.messaging?.candidate_observation?.runtime_event_action === 'observe_platform_candidates',
      candidate_observer_message_type: bundle.messaging?.candidate_observation?.message_type,
      candidate_observer_permission: bundle.messaging?.candidate_observation?.required_permission,
      observer_plan_ready: bundle.readiness.observer_plan_ready === true,
      observer_preflight_status: bundle.readiness.observer_preflight_status,
      observer_factory: bundle.runtime?.observation_loop?.factory,
      observer_fallback_poll_interval_ms: bundle.runtime?.observation_loop?.cadence?.fallback_poll_interval_ms,
      observer_changed_debounce_ms: bundle.runtime?.observation_loop?.cadence?.changed_observe_every_ms,
      observer_meeting_end_grace_ms: bundle.runtime?.observation_loop?.cadence?.meeting_missing_end_grace_ms,
      sample_interval_ms: bundle.runtime?.start_options?.sampleIntervalMs,
      mutation_debounce_ms: bundle.runtime?.mutation_observer?.debounce_ms,
      speaker_min_stable_ms: bundle.runtime?.speaker_filter?.min_stable_ms,
      provider_transport: bundle.provider_reconcile?.transport,
      provider_required_for_realtime: bundle.readiness.provider_required_for_realtime,
      transcript_blocks_realtime: bundle.readiness.transcript_blocks_realtime,
      missing_item_count: bundle.readiness.missing_items?.length ?? 0,
    })),
    bundles,
    next_actions: unique(bundles.flatMap((bundle) => bundle.next_actions ?? [])),
  };
}
