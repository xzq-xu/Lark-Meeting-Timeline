import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  createMeetingAppBrowserRuntime,
  meetingAppBrowserInput,
} from './meeting-app-browser-runtime.mjs';
import {
  createMeetingAppObserverScheduler,
  buildMeetingAppObserverSchedulerConfig,
  buildMeetingAppObserverSchedulerConfigMatrix,
  normalizeMeetingAppObserverTrigger,
} from './meeting-app-observer-scheduler.mjs';
import {
  buildMeetingPlatformRuntimeBundle,
  buildMeetingPlatformRuntimeBundleMatrix,
} from './platform-runtime-bundle.mjs';
import {
  MEETING_PLATFORM_KEYS,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_RUNTIME_HOST_CONFIG_SCHEMA = 'meeting_platform_runtime_host_config';
export const MEETING_PLATFORM_RUNTIME_HOST_CONFIG_MATRIX_SCHEMA = 'meeting_platform_runtime_host_config_matrix';
export const MEETING_PLATFORM_RUNTIME_HOST_HANDOFF_SCHEMA = 'meeting_platform_runtime_host_handoff';
export const MEETING_PLATFORM_RUNTIME_HOST_HANDOFF_MATRIX_SCHEMA = 'meeting_platform_runtime_host_handoff_matrix';
export const MEETING_PLATFORM_RUNTIME_HOST_STATE_SCHEMA = 'meeting_platform_runtime_host_state';
export const MEETING_PLATFORM_RUNTIME_HOST_SCHEMA_VERSION = 1;

const DEFAULT_RUNTIME_HOST_PLATFORMS = MEETING_PLATFORM_KEYS.filter((platform) => platform !== 'local_detector');

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
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, DEFAULT_RUNTIME_HOST_PLATFORMS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

function normalizeBundle(bundleOrPlatform = {}, options = {}) {
  if (bundleOrPlatform?.schema === 'meeting_platform_runtime_bundle') return bundleOrPlatform;
  const input = typeof bundleOrPlatform === 'string' ? { platform: bundleOrPlatform } : (bundleOrPlatform ?? {});
  const platform = firstNonEmpty(
    options.platform,
    options.provider,
    input.platform,
    input.provider,
    input.key,
    input.name,
  );
  if (!platform) {
    throw new MeetingTimelineSdkError('meeting platform runtime host requires a platform or runtime bundle', {
      supported_platforms: DEFAULT_RUNTIME_HOST_PLATFORMS,
    });
  }
  return buildMeetingPlatformRuntimeBundle(normalizeMeetingPlatform(platform), {
    ...input,
    ...options,
  });
}

function numeric(value, fallback = 0) {
  const number = Number(value);
  return Number.isFinite(number) ? Math.max(0, number) : fallback;
}

function hasRuntimeShape(value) {
  return Boolean(value)
    && typeof value === 'object'
    && (typeof value.sample === 'function' || typeof value.tick === 'function');
}

function browserWindow(options = {}) {
  return options.window ?? options.win ?? globalThis.window;
}

function browserDocument(options = {}) {
  const win = browserWindow(options);
  return options.document ?? options.doc ?? win?.document ?? globalThis.document;
}

function eventTarget(options = {}) {
  return options.eventTarget ?? options.event_target ?? browserWindow(options);
}

function mutationObserverCtor(options = {}) {
  const win = browserWindow(options);
  return options.MutationObserver
    ?? options.mutationObserver
    ?? options.mutation_observer
    ?? options.mutationObserverCtor
    ?? options.mutation_observer_ctor
    ?? win?.MutationObserver
    ?? globalThis.MutationObserver;
}

function mutationRoot(options = {}) {
  const doc = browserDocument(options);
  return options.mutationRoot
    ?? options.mutation_root
    ?? doc?.body
    ?? doc?.documentElement
    ?? doc;
}

function mutationObserveOptions(options = {}) {
  return {
    childList: true,
    subtree: true,
    attributes: true,
    characterData: true,
    ...(options.mutationObserverOptions ?? {}),
    ...(options.mutation_observer_options ?? {}),
  };
}

function stopEvents(options = {}) {
  const events = firstNonEmpty(options.stopEvents, options.stop_events, ['pagehide', 'beforeunload']);
  return asArray(events).filter(Boolean).map(String);
}

function inputProviderFor(runtime = {}, options = {}) {
  const explicit = firstNonEmpty(options.inputProvider, options.input_provider);
  if (typeof explicit === 'function') return explicit;
  if (runtime.inputProvider && typeof runtime.inputProvider === 'function') return runtime.inputProvider.bind(runtime);
  const input = firstNonEmpty(options.input, options.snapshot);
  if (input != null) return () => input;
  return (extra = {}) => meetingAppBrowserInput({ ...options, ...extra });
}

function runtimeOptionsFrom(config = {}, options = {}) {
  return compactObject({
    ...(config.runtime_bundle?.runtime?.start_options ?? {}),
    ...(options.runtimeOptions ?? {}),
    ...(options.runtime_options ?? {}),
    ...(options.browserRuntimeOptions ?? {}),
    ...(options.browser_runtime_options ?? {}),
    platform: config.platform,
    provider: config.platform,
  });
}

function createRuntime(clientOrRuntime, config = {}, options = {}) {
  if (hasRuntimeShape(options.runtime)) return options.runtime;
  if (hasRuntimeShape(clientOrRuntime)) return clientOrRuntime;
  return createMeetingAppBrowserRuntime(clientOrRuntime, runtimeOptionsFrom(config, options));
}

function normalizeState(state = {}) {
  return {
    type: 'meeting_platform_runtime_host_state',
    schema: MEETING_PLATFORM_RUNTIME_HOST_STATE_SCHEMA,
    version: MEETING_PLATFORM_RUNTIME_HOST_SCHEMA_VERSION,
    running: state.running === true,
    platform: state.platform ?? null,
    change_observer_installed: state.change_observer_installed === true,
    lifecycle_installed: state.lifecycle_installed === true,
    scheduled_trigger_count: numeric(state.scheduled_trigger_count ?? state.scheduledTriggerCount, 0),
    last_trigger: state.last_trigger ?? state.lastTrigger ?? null,
    last_error: state.last_error ?? state.lastError ?? null,
  };
}

function mutationAllowed(records = [], options = {}) {
  const rows = asArray(records);
  const filter = options.mutationFilter ?? options.mutation_filter;
  if (typeof filter === 'function') {
    return rows.some((record) => filter(record, options) !== false);
  }
  return rows.length > 0;
}

function schedulerConfigFromBundle(bundle = {}, options = {}) {
  const existing = bundle.runtime?.observer_scheduler?.config;
  if (existing?.schema === 'meeting_app_observer_scheduler_config') return existing;
  if (!bundle.runtime?.observer_plan) return null;
  return buildMeetingAppObserverSchedulerConfig(bundle.runtime.observer_plan, options);
}

export function buildMeetingPlatformRuntimeHostConfig(bundleOrPlatform = {}, options = {}) {
  const bundle = normalizeBundle(bundleOrPlatform, options);
  const schedulerConfig = schedulerConfigFromBundle(bundle, options);
  return compactObject({
    type: 'meeting_platform_runtime_host_config',
    schema: MEETING_PLATFORM_RUNTIME_HOST_CONFIG_SCHEMA,
    version: MEETING_PLATFORM_RUNTIME_HOST_SCHEMA_VERSION,
    platform: bundle.platform,
    display_name: bundle.display_name,
    runtime_bundle_id: bundle.id,
    runtime_bundle_schema: bundle.schema,
    runtime_bundle: bundle,
    runtime_factory: schedulerConfig?.runtime_factory ?? 'createMeetingAppBrowserRuntime',
    runtime_options: runtimeOptionsFrom({
      platform: bundle.platform,
      runtime_bundle: bundle,
    }, options),
    observer_plan: bundle.runtime?.observer_plan ?? null,
    observer_scheduler_config: schedulerConfig,
    driver: {
      change_observer: {
        enabled: bundle.runtime?.mutation_observer?.enabled === true && schedulerConfig != null,
        trigger: bundle.platform === 'local_detector' ? 'host_snapshot_change' : 'dom_mutation',
        schedule_by_default: true,
        observe_options: mutationObserveOptions(options),
      },
      keep_alive: {
        enabled: schedulerConfig != null,
        interval_ms: schedulerConfig?.cadence?.unchanged_observe_every_ms
          ?? schedulerConfig?.cadence?.fallback_poll_interval_ms
          ?? 10_000,
      },
      lifecycle: {
        enabled: true,
        stop_events: stopEvents(options),
      },
      runtime_start_default: false,
    },
    host_api: [
      'start',
      'stop',
      'changed',
      'keepAlive',
      'speakerCandidate',
      'candidateMissing',
      'sample',
      'getState',
    ],
    readiness: {
      runtime_ready: bundle.readiness?.runtime_ready === true,
      observer_plan_ready: bundle.readiness?.observer_plan_ready === true,
      observer_scheduler_ready: schedulerConfig?.sdk_ready === true,
      host_ready: bundle.readiness?.runtime_ready === true && schedulerConfig?.sdk_ready === true,
      provider_required_for_realtime: bundle.readiness?.provider_required_for_realtime === true,
      transcript_blocks_realtime: bundle.readiness?.transcript_blocks_realtime === true,
    },
    next_actions: unique([
      'createMeetingPlatformRuntimeHost(client, hostConfig)',
      'call_host.start_when_meeting_surface_is_open',
      'feed_host.changed_from_dom_or_native_change_events',
      'feed_host.candidateMissing_when_meeting_surface_disappears',
      ...(bundle.next_actions ?? []),
    ]),
  });
}

export function buildMeetingPlatformRuntimeHostConfigMatrix(options = {}) {
  const platforms = selectedPlatforms(options);
  const configs = platforms.map((platform) => buildMeetingPlatformRuntimeHostConfig(platform, {
    ...options,
    platform,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_runtime_host_config_matrix',
    schema: MEETING_PLATFORM_RUNTIME_HOST_CONFIG_MATRIX_SCHEMA,
    version: MEETING_PLATFORM_RUNTIME_HOST_SCHEMA_VERSION,
    platform_count: configs.length,
    host_ready_count: configs.filter((config) => config.readiness.host_ready).length,
    change_observer_count: configs.filter((config) => config.driver.change_observer.enabled).length,
    keep_alive_count: configs.filter((config) => config.driver.keep_alive.enabled).length,
    platforms,
    rows: configs.map((config) => ({
      platform: config.platform,
      display_name: config.display_name,
      runtime_factory: config.runtime_factory,
      host_ready: config.readiness.host_ready,
      observer_scheduler_ready: config.readiness.observer_scheduler_ready,
      change_observer_enabled: config.driver.change_observer.enabled,
      keep_alive_interval_ms: config.driver.keep_alive.interval_ms,
      provider_required_for_realtime: config.readiness.provider_required_for_realtime,
      transcript_blocks_realtime: config.readiness.transcript_blocks_realtime,
    })),
    configs,
    runtime_bundle_matrix: buildMeetingPlatformRuntimeBundleMatrix({
      ...options,
      platforms,
    }),
    observer_scheduler_config_matrix: buildMeetingAppObserverSchedulerConfigMatrix({
      ...options,
      platforms,
    }),
    next_actions: unique(configs.flatMap((config) => config.next_actions ?? [])),
  };
}

function packageExample(platform, options = {}) {
  const variableName = String(firstNonEmpty(options.clientVariable, options.client_variable, 'timeline'));
  return [
    "import { createMeetingPlatformRuntimeHost } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-runtime-host';",
    '',
    `const host = createMeetingPlatformRuntimeHost(${variableName}, '${platform}', { window, document });`,
    'host.start();',
    '',
    '// Optional host hooks:',
    'host.changed();',
    'host.speakerCandidate();',
    'host.candidateMissing();',
    'host.stop();',
  ].join('\n');
}

function handoffAcceptance(config = {}) {
  const checks = [
    {
      id: 'runtime_bundle_ready',
      ok: config.readiness?.runtime_ready === true,
      evidence: 'runtime_bundle.readiness.runtime_ready',
    },
    {
      id: 'observer_scheduler_ready',
      ok: config.readiness?.observer_scheduler_ready === true,
      evidence: 'observer_scheduler_config.sdk_ready',
    },
    {
      id: 'change_observer_enabled',
      ok: config.driver?.change_observer?.enabled === true,
      evidence: 'driver.change_observer.enabled',
    },
    {
      id: 'keep_alive_enabled',
      ok: config.driver?.keep_alive?.enabled === true,
      evidence: 'driver.keep_alive.enabled',
    },
    {
      id: 'provider_not_required_for_realtime',
      ok: config.readiness?.provider_required_for_realtime !== true,
      evidence: 'readiness.provider_required_for_realtime',
    },
    {
      id: 'transcript_not_required_for_realtime',
      ok: config.readiness?.transcript_blocks_realtime !== true,
      evidence: 'readiness.transcript_blocks_realtime',
    },
  ];
  return {
    accepted: checks.every((check) => check.ok === true),
    check_count: checks.length,
    passed_count: checks.filter((check) => check.ok).length,
    checks,
  };
}

export function buildMeetingPlatformRuntimeHostHandoff(bundleOrPlatform = {}, options = {}) {
  const config = bundleOrPlatform?.schema === MEETING_PLATFORM_RUNTIME_HOST_CONFIG_SCHEMA
    ? bundleOrPlatform
    : buildMeetingPlatformRuntimeHostConfig(bundleOrPlatform, options);
  const bundle = config.runtime_bundle ?? {};
  const acceptance = handoffAcceptance(config);
  return compactObject({
    type: 'meeting_platform_runtime_host_handoff',
    schema: MEETING_PLATFORM_RUNTIME_HOST_HANDOFF_SCHEMA,
    version: MEETING_PLATFORM_RUNTIME_HOST_SCHEMA_VERSION,
    platform: config.platform,
    display_name: config.display_name,
    objective: 'handoff_package_for_embedding_realtime_meeting_timeline_runtime_host',
    package_entry: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-runtime-host',
    imports: {
      create_host: 'createMeetingPlatformRuntimeHost',
      build_config: 'buildMeetingPlatformRuntimeHostConfig',
      build_matrix: 'buildMeetingPlatformRuntimeHostConfigMatrix',
    },
    host_config: config,
    browser: {
      matches: bundle.browser?.matches ?? [],
      host_permissions: bundle.browser?.host_permissions ?? [],
      permissions: bundle.browser?.permissions ?? [],
      content_scripts: bundle.browser?.content_scripts ?? [],
      manifest: bundle.browser?.manifest,
    },
    runtime: {
      factory: config.runtime_factory,
      driver: config.driver,
      timestamp_field: bundle.host?.annotation_timestamp_field ?? config.observer_scheduler_config?.timestamp_field ?? 'captured_at_ms',
      provider_events_role: config.observer_plan?.signal_contract?.provider_events_role ?? 'reconcile_and_backfill_only',
      transcript_role: config.observer_plan?.signal_contract?.transcript_role ?? 'post_meeting_backfill_only',
    },
    host_hooks: [
      {
        hook: 'surface_open_or_content_script_attached',
        call: 'host.start()',
        purpose: 'start_keep_alive_and_change_observer',
      },
      {
        hook: 'dom_mutation_or_native_snapshot_change',
        call: 'host.changed()',
        purpose: 'debounced_realtime_axis_sample',
      },
      {
        hook: 'active_speaker_candidate_detected',
        call: 'host.speakerCandidate()',
        purpose: 'schedule_stable_speaker_followup',
      },
      {
        hook: 'meeting_surface_missing_or_tab_closed',
        call: 'host.candidateMissing()',
        purpose: 'emit_meeting_ended_after_grace_window',
      },
      {
        hook: 'surface_unload_or_host_dispose',
        call: 'host.stop()',
        purpose: 'clear_observers_and_timers',
      },
    ],
    example: packageExample(config.platform, options),
    acceptance,
    readiness: {
      accepted: acceptance.accepted,
      host_ready: config.readiness?.host_ready === true,
      runtime_ready: config.readiness?.runtime_ready === true,
      provider_required_for_realtime: config.readiness?.provider_required_for_realtime === true,
      transcript_blocks_realtime: config.readiness?.transcript_blocks_realtime === true,
    },
    next_actions: unique([
      'embed_example_bootstrap_in_browser_extension_webview_or_native_host',
      'forward_surface_change_events_to_host.changed',
      'forward_surface_missing_or_close_to_host.candidateMissing',
      'validate_with_live_meeting_snapshot_before_pilot',
      ...(config.next_actions ?? []),
    ]),
  });
}

export function buildMeetingPlatformRuntimeHostHandoffMatrix(options = {}) {
  const platforms = selectedPlatforms(options);
  const handoffs = platforms.map((platform) => buildMeetingPlatformRuntimeHostHandoff(platform, {
    ...options,
    platform,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_runtime_host_handoff_matrix',
    schema: MEETING_PLATFORM_RUNTIME_HOST_HANDOFF_MATRIX_SCHEMA,
    version: MEETING_PLATFORM_RUNTIME_HOST_SCHEMA_VERSION,
    platform_count: handoffs.length,
    accepted_count: handoffs.filter((handoff) => handoff.acceptance?.accepted === true).length,
    host_ready_count: handoffs.filter((handoff) => handoff.readiness?.host_ready === true).length,
    provider_blocking_count: handoffs.filter((handoff) => handoff.readiness?.provider_required_for_realtime === true).length,
    transcript_blocking_count: handoffs.filter((handoff) => handoff.readiness?.transcript_blocks_realtime === true).length,
    platforms,
    rows: handoffs.map((handoff) => ({
      platform: handoff.platform,
      display_name: handoff.display_name,
      accepted: handoff.acceptance?.accepted === true,
      host_ready: handoff.readiness?.host_ready === true,
      browser_match_count: handoff.browser?.matches?.length ?? 0,
      host_hook_count: handoff.host_hooks?.length ?? 0,
      timestamp_field: handoff.runtime?.timestamp_field,
      provider_events_role: handoff.runtime?.provider_events_role,
      transcript_role: handoff.runtime?.transcript_role,
    })),
    handoffs,
    next_actions: unique(handoffs.flatMap((handoff) => handoff.next_actions ?? [])),
  };
}

export function createMeetingPlatformRuntimeHost(clientOrRuntime, bundleOrPlatform = {}, options = {}) {
  const config = bundleOrPlatform?.schema === MEETING_PLATFORM_RUNTIME_HOST_CONFIG_SCHEMA
    ? bundleOrPlatform
    : buildMeetingPlatformRuntimeHostConfig(bundleOrPlatform, options);
  if (!config.observer_scheduler_config) {
    throw new MeetingTimelineSdkError('meeting platform runtime host requires an observer scheduler config', {
      platform: config.platform,
      reason: 'missing_observer_scheduler_config',
    });
  }
  const runtime = createRuntime(clientOrRuntime, config, options);
  const inputProvider = inputProviderFor(runtime, {
    ...runtimeOptionsFrom(config, options),
    ...options,
  });
  const scheduler = options.scheduler ?? createMeetingAppObserverScheduler(runtime, config.observer_plan ?? config.platform, {
    ...options,
    config: config.observer_scheduler_config,
    inputProvider,
    runtimeInputMode: firstNonEmpty(options.runtimeInputMode, options.runtime_input_mode, runtime.inputProvider ? 'provider' : 'explicit'),
  });
  const timeout = options.setTimeout ?? globalThis.setTimeout;
  const target = eventTarget(options);
  const listeners = [];
  let state = normalizeState({ platform: config.platform });
  let observer = null;
  const scheduledByTrigger = new Map();

  function getState() {
    return {
      ...state,
      running: state.running === true,
      platform: config.platform,
      change_observer_installed: observer != null,
      lifecycle_installed: listeners.length > 0,
      scheduler: scheduler.getState?.(),
      runtime: runtime.getState?.(),
    };
  }

  function setError(error) {
    state = {
      ...state,
      last_error: error?.message ?? String(error),
    };
  }

  function scheduleTrigger(trigger, input = undefined, triggerOptions = {}) {
    const normalized = normalizeMeetingAppObserverTrigger(trigger);
    const previous = scheduledByTrigger.get(normalized);
    if (previous?.id && scheduler.clearScheduled) scheduler.clearScheduled(previous.id);
    const result = scheduler.handleTrigger(normalized, input, {
      ...triggerOptions,
      schedule: triggerOptions.schedule ?? true,
    });
    if (result?.scheduled && result.id) {
      scheduledByTrigger.set(normalized, result);
      state = {
        ...state,
        scheduled_trigger_count: state.scheduled_trigger_count + 1,
        last_trigger: normalized,
      };
    }
    return result;
  }

  function installChangeObserver(startOptions = {}) {
    if (observer) return { installed: false, reason: 'already_installed', state: getState() };
    if (firstNonEmpty(startOptions.observeChanges, startOptions.observe_changes, config.driver.change_observer.enabled) !== true) {
      return { installed: false, reason: 'disabled', state: getState() };
    }
    const Observer = mutationObserverCtor({ ...options, ...startOptions });
    const root = mutationRoot({ ...options, ...startOptions });
    if (!Observer) return { installed: false, reason: 'missing_mutation_observer', state: getState() };
    if (!root) return { installed: false, reason: 'missing_mutation_root', state: getState() };
    const observerOptions = mutationObserveOptions({ ...options, ...startOptions });
    observer = new Observer((records = []) => {
      if (!mutationAllowed(records, { ...options, ...startOptions })) return;
      try {
        scheduleTrigger(config.driver.change_observer.trigger, undefined, {
          ...(startOptions.changeTriggerOptions ?? {}),
          ...(startOptions.change_trigger_options ?? {}),
        });
      } catch (error) {
        setError(error);
      }
    });
    observer.observe(root, observerOptions);
    state = {
      ...state,
      change_observer_installed: true,
    };
    return { installed: true, root, options: observerOptions, state: getState() };
  }

  function removeChangeObserver() {
    if (observer?.disconnect) observer.disconnect();
    observer = null;
    state = {
      ...state,
      change_observer_installed: false,
    };
    return getState();
  }

  function installLifecycleHandlers(startOptions = {}) {
    if (listeners.length > 0) return { installed: false, reason: 'already_installed', events: listeners.map(([name]) => name) };
    if (startOptions.lifecycle === false || startOptions.installLifecycleHandlers === false || startOptions.install_lifecycle_handlers === false) {
      return { installed: false, reason: 'disabled' };
    }
    if (!target?.addEventListener) return { installed: false, reason: 'missing_event_target' };
    const handler = () => stop();
    for (const eventName of stopEvents({ ...options, ...startOptions })) {
      target.addEventListener(eventName, handler);
      listeners.push([eventName, handler]);
    }
    state = {
      ...state,
      lifecycle_installed: true,
    };
    return { installed: true, events: listeners.map(([name]) => name) };
  }

  function removeLifecycleHandlers() {
    if (target?.removeEventListener) {
      for (const [eventName, handler] of listeners.splice(0)) target.removeEventListener(eventName, handler);
    } else {
      listeners.splice(0);
    }
    state = {
      ...state,
      lifecycle_installed: false,
    };
    return getState();
  }

  function start(startOptions = {}) {
    installLifecycleHandlers(startOptions);
    installChangeObserver(startOptions);
    if (firstNonEmpty(startOptions.startScheduler, startOptions.start_scheduler, true) === true) {
      scheduler.start(inputProvider, {
        immediate: startOptions.immediate === true,
        ...(startOptions.schedulerStartOptions ?? {}),
        ...(startOptions.scheduler_start_options ?? {}),
      });
    }
    if (firstNonEmpty(startOptions.startRuntime, startOptions.start_runtime, config.driver.runtime_start_default) === true
      && typeof runtime.start === 'function') {
      runtime.start(startOptions);
    }
    state = {
      ...state,
      running: true,
      last_error: null,
    };
    return getState();
  }

  function stop() {
    removeChangeObserver();
    removeLifecycleHandlers();
    scheduler.stop?.();
    if (typeof runtime.stop === 'function') runtime.stop();
    for (const result of scheduledByTrigger.values()) {
      if (result?.id && scheduler.clearScheduled) scheduler.clearScheduled(result.id);
    }
    scheduledByTrigger.clear();
    state = {
      ...state,
      running: false,
    };
    return getState();
  }

  function changed(input = undefined, triggerOptions = {}) {
    return triggerOptions.immediate === true
      ? scheduler.triggerChanged(input, triggerOptions)
      : scheduleTrigger(config.driver.change_observer.trigger, input, triggerOptions);
  }

  function sample(trigger = 'manual', input = undefined, sampleOptions = {}) {
    return scheduler.sample(trigger, input, sampleOptions);
  }

  function keepAlive(input = undefined, triggerOptions = {}) {
    return scheduler.triggerKeepAlive(input, triggerOptions);
  }

  function speakerCandidate(input = undefined, triggerOptions = {}) {
    return triggerOptions.immediate === true
      ? scheduler.triggerSpeakerCandidate(input, triggerOptions)
      : scheduleTrigger('active_speaker_candidate', input, triggerOptions);
  }

  function candidateMissing(input = undefined, triggerOptions = {}) {
    return triggerOptions.immediate === true
      ? scheduler.triggerCandidateMissing(input, triggerOptions)
      : scheduleTrigger('meeting_candidate_missing', input, triggerOptions);
  }

  return {
    config,
    runtime,
    scheduler,
    inputProvider,
    start,
    stop,
    dispose: stop,
    installChangeObserver,
    removeChangeObserver,
    installLifecycleHandlers,
    removeLifecycleHandlers,
    changed,
    keepAlive,
    speakerCandidate,
    candidateMissing,
    sample,
    getState,
    reset(nextState = {}) {
      state = normalizeState({ ...nextState, platform: config.platform });
      scheduler.reset?.(nextState.scheduler ?? nextState.scheduler_state ?? {});
      return getState();
    },
    scheduleTrigger,
    setTimeout: timeout,
  };
}
