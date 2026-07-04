import { MeetingTimelineSdkError, compactObject, normalizeAbsoluteMs } from '../index.mjs';
import {
  MEETING_APP_RUNTIME_OBSERVER_PLAN_SCHEMA,
  buildMeetingAppRuntimeObserverPlan,
  buildMeetingAppRuntimeObserverPlanMatrix,
} from './meeting-app-profile.mjs';

export const MEETING_APP_OBSERVER_SCHEDULER_CONFIG_SCHEMA = 'meeting_app_observer_scheduler_config';
export const MEETING_APP_OBSERVER_SCHEDULER_CONFIG_MATRIX_SCHEMA = 'meeting_app_observer_scheduler_config_matrix';
export const MEETING_APP_OBSERVER_SCHEDULER_STATE_SCHEMA = 'meeting_app_observer_scheduler_state';
export const MEETING_APP_OBSERVER_SCHEDULER_SCHEMA_VERSION = 1;

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

function numeric(value, fallback = 0) {
  const number = Number(value);
  return Number.isFinite(number) ? Math.max(0, number) : fallback;
}

function nowMs(options = {}) {
  const raw = typeof options.now === 'function'
    ? options.now()
    : firstNonEmpty(options.now, options.nowMs, options.now_ms, Date.now());
  return normalizeAbsoluteMs(raw, 'meeting_app_observer_scheduler_time');
}

function normalizePlan(planOrPlatform = {}, options = {}) {
  if (planOrPlatform?.schema === MEETING_APP_RUNTIME_OBSERVER_PLAN_SCHEMA) return planOrPlatform;
  return buildMeetingAppRuntimeObserverPlan(planOrPlatform, options);
}

function triggerNames(plan = {}) {
  return asArray(plan.trigger_policy).map((item) => item.trigger).filter(Boolean);
}

function triggerPolicyMap(plan = {}) {
  return Object.fromEntries(asArray(plan.trigger_policy)
    .filter((item) => item?.trigger)
    .map((item) => [item.trigger, item]));
}

export function normalizeMeetingAppObserverTrigger(trigger = 'manual') {
  const text = String(trigger ?? 'manual').trim().toLowerCase().replace(/[-\s]+/g, '_');
  const aliases = {
    changed: 'dom_mutation',
    mutation: 'dom_mutation',
    mutations: 'dom_mutation',
    dom_changed: 'dom_mutation',
    host_changed: 'host_snapshot_change',
    host_change: 'host_snapshot_change',
    keepalive: 'unchanged_keep_alive',
    keep_alive: 'unchanged_keep_alive',
    interval: 'unchanged_keep_alive',
    speaker: 'active_speaker_candidate',
    speaker_candidate: 'active_speaker_candidate',
    active_speaker: 'active_speaker_candidate',
    missing: 'meeting_candidate_missing',
    candidate_missing: 'meeting_candidate_missing',
    meeting_missing: 'meeting_candidate_missing',
    end_grace: 'meeting_candidate_missing',
  };
  return aliases[text] ?? text;
}

function triggerDelayMs(trigger, config = {}, options = {}) {
  const normalized = normalizeMeetingAppObserverTrigger(trigger);
  const policy = config.trigger_policy_map?.[normalized] ?? {};
  const cadence = config.cadence ?? {};
  const explicit = firstNonEmpty(options.delayMs, options.delay_ms, policy.debounce_ms, policy.interval_ms, policy.delay_ms, policy.grace_ms);
  if (explicit != null) return numeric(explicit, 0);
  if (normalized === 'unchanged_keep_alive') return cadence.unchanged_observe_every_ms ?? cadence.fallback_poll_interval_ms ?? 1_000;
  if (normalized === 'active_speaker_candidate') return cadence.speaker_stable_followup_ms ?? 300;
  if (normalized === 'meeting_candidate_missing') return cadence.meeting_missing_end_grace_ms ?? 4_000;
  return cadence.changed_observe_every_ms ?? cadence.min_observe_interval_ms ?? 250;
}

function schedulerTrackTriggers(options = {}) {
  const raw = firstNonEmpty(options.trackTriggers, options.track_triggers, [
    'dom_mutation',
    'host_snapshot_change',
    'active_speaker_candidate',
  ]);
  return unique(asArray(raw).map((trigger) => normalizeMeetingAppObserverTrigger(trigger)));
}

function triggerSampleOptions(trigger, config = {}, options = {}) {
  const normalized = normalizeMeetingAppObserverTrigger(trigger);
  const force = firstNonEmpty(
    options.force,
    normalized === 'active_speaker_candidate' || normalized === 'meeting_candidate_missing' ? true : undefined,
  );
  return compactObject({
    ...(config.sample_options ?? {}),
    ...(options.sampleOptions ?? {}),
    ...(options.sample_options ?? {}),
    ...options,
    trigger: normalized,
    trigger_reason: normalized,
    force,
    stabilityFollowup: normalized === 'active_speaker_candidate' ? true : undefined,
    stability_followup: normalized === 'active_speaker_candidate' ? true : undefined,
    missingCandidate: normalized === 'meeting_candidate_missing' ? true : undefined,
    missing_candidate: normalized === 'meeting_candidate_missing' ? true : undefined,
  });
}

function trackTriggerEnabled(trigger, config = {}, options = {}) {
  const policy = config.track_sample_policy ?? {};
  const enabled = firstNonEmpty(options.observeTracks, options.observe_tracks, policy.enabled, false);
  if (enabled !== true) return false;
  return asArray(policy.triggers).includes(normalizeMeetingAppObserverTrigger(trigger));
}

function inputFromProvider(provider, fallbackInput) {
  if (typeof provider === 'function') return provider();
  return provider ?? fallbackInput ?? {};
}

function runtimeInputMode(runtime = {}, config = {}, options = {}) {
  const explicit = firstNonEmpty(options.runtimeInputMode, options.runtime_input_mode);
  if (explicit) return String(explicit).trim();
  if (runtime.inputProvider && config.runtime_factory === 'createMeetingAppBrowserRuntime') return 'provider';
  return 'explicit';
}

function normalizeSchedulerState(state = {}) {
  return {
    type: 'meeting_app_observer_scheduler_state',
    schema: MEETING_APP_OBSERVER_SCHEDULER_STATE_SCHEMA,
    version: MEETING_APP_OBSERVER_SCHEDULER_SCHEMA_VERSION,
    running: state.running === true,
    trigger_count: numeric(state.trigger_count ?? state.triggerCount, 0),
    sample_count: numeric(state.sample_count ?? state.sampleCount, 0),
    track_sample_count: numeric(state.track_sample_count ?? state.trackSampleCount, 0),
    scheduled_count: numeric(state.scheduled_count ?? state.scheduledCount, 0),
    pending_timer_count: numeric(state.pending_timer_count ?? state.pendingTimerCount, 0),
    last_trigger: state.last_trigger ?? state.lastTrigger ?? null,
    last_sample_at_ms: state.last_sample_at_ms ?? state.lastSampleAtMs ?? null,
    last_result: state.last_result ?? state.lastResult ?? null,
    last_track_result: state.last_track_result ?? state.lastTrackResult ?? null,
    last_error: state.last_error ?? state.lastError ?? null,
  };
}

export function buildMeetingAppObserverSchedulerConfig(planOrPlatform = {}, options = {}) {
  const plan = normalizePlan(planOrPlatform, options);
  const cadence = {
    min_observe_interval_ms: numeric(plan.cadence?.min_observe_interval_ms, 250),
    changed_observe_every_ms: numeric(plan.cadence?.changed_observe_every_ms, 250),
    unchanged_observe_every_ms: numeric(plan.cadence?.unchanged_observe_every_ms, 10_000),
    fallback_poll_interval_ms: numeric(plan.cadence?.fallback_poll_interval_ms, 10_000),
    speaker_stable_followup_ms: numeric(plan.cadence?.speaker_stable_followup_ms, 300),
    speaker_min_stable_ms: numeric(plan.cadence?.speaker_min_stable_ms, 650),
    speaker_end_idle_ms: numeric(plan.cadence?.speaker_end_idle_ms, 1_500),
    meeting_missing_end_grace_ms: numeric(plan.cadence?.meeting_missing_end_grace_ms, 4_000),
  };
  const triggers = triggerNames(plan);
  return compactObject({
    type: 'meeting_app_observer_scheduler_config',
    schema: MEETING_APP_OBSERVER_SCHEDULER_CONFIG_SCHEMA,
    version: MEETING_APP_OBSERVER_SCHEDULER_SCHEMA_VERSION,
    platform: plan.platform,
    surface: plan.surface,
    display_name: plan.display_name,
    plan_schema: plan.schema,
    sdk_ready: plan.sdk_ready,
    preflight_status: plan.preflight_status,
    runtime_factory: plan.observer_runtime?.factory,
    observe_method: plan.observer_runtime?.observe_method ?? 'observeMeetingApp',
    track_observe_method: plan.observer_runtime?.track_observe_method ?? 'observeMeetingAppTracks',
    timestamp_field: plan.signal_contract?.timestamp_field ?? 'captured_at_ms',
    input_contract: plan.input_contract,
    cadence,
    trigger_names: triggers,
    trigger_policy: plan.trigger_policy ?? [],
    trigger_policy_map: triggerPolicyMap(plan),
    sample_options: {
      platform: plan.platform,
      provider: plan.platform,
      captureOptions: plan.observer_runtime?.capture_options,
      capture_options: plan.observer_runtime?.capture_options,
      observeOptions: {
        source: firstNonEmpty(options.source, options.observeSource, options.observe_source, 'meeting_app_observer_scheduler'),
      },
      observe_options: {
        source: firstNonEmpty(options.source, options.observeSource, options.observe_source, 'meeting_app_observer_scheduler'),
      },
      minObserveIntervalMs: cadence.min_observe_interval_ms,
      min_observe_interval_ms: cadence.min_observe_interval_ms,
      changedObserveEveryMs: cadence.changed_observe_every_ms,
      changed_observe_every_ms: cadence.changed_observe_every_ms,
      unchangedObserveEveryMs: cadence.unchanged_observe_every_ms,
      unchanged_observe_every_ms: cadence.unchanged_observe_every_ms,
      sampleIntervalMs: cadence.fallback_poll_interval_ms,
      sample_interval_ms: cadence.fallback_poll_interval_ms,
      speakerStableFollowupMs: cadence.speaker_stable_followup_ms,
      speaker_stable_followup_ms: cadence.speaker_stable_followup_ms,
    },
    track_sample_policy: {
      enabled: firstNonEmpty(options.observeTracks, options.observe_tracks, false) === true,
      triggers: schedulerTrackTriggers(options),
      output_intents: plan.track_runtime?.output_intents ?? [],
      content_policy: plan.track_runtime?.content_policy,
    },
    plan,
    next_actions: unique([
      'createMeetingAppObserverScheduler(runtime, observerPlan)',
      'call_scheduler.handleTrigger_from_dom_or_native_change_events',
      'call_scheduler.start_for_keep_alive_sampling',
      ...(plan.next_actions ?? []),
    ]),
  });
}

export function buildMeetingAppObserverSchedulerConfigMatrix(options = {}) {
  const planMatrix = buildMeetingAppRuntimeObserverPlanMatrix(options);
  const configs = planMatrix.plans.map((plan) => buildMeetingAppObserverSchedulerConfig(plan, options));
  return {
    type: 'meeting_app_observer_scheduler_config_matrix',
    schema: MEETING_APP_OBSERVER_SCHEDULER_CONFIG_MATRIX_SCHEMA,
    version: MEETING_APP_OBSERVER_SCHEDULER_SCHEMA_VERSION,
    platform_count: configs.length,
    sdk_ready_count: configs.filter((config) => config.sdk_ready === true).length,
    track_enabled_count: configs.filter((config) => config.track_sample_policy?.enabled === true).length,
    platforms: configs.map((config) => config.platform),
    rows: configs.map((config) => ({
      platform: config.platform,
      surface: config.surface,
      sdk_ready: config.sdk_ready,
      preflight_status: config.preflight_status,
      runtime_factory: config.runtime_factory,
      trigger_count: config.trigger_names.length,
      fallback_poll_interval_ms: config.cadence.fallback_poll_interval_ms,
      changed_observe_every_ms: config.cadence.changed_observe_every_ms,
      meeting_missing_end_grace_ms: config.cadence.meeting_missing_end_grace_ms,
      track_enabled: config.track_sample_policy?.enabled === true,
    })),
    configs,
    observer_plan_matrix: planMatrix,
    next_actions: unique(configs.flatMap((config) => config.next_actions ?? [])),
  };
}

function normalizeRuntime(runtime) {
  if (!runtime || typeof runtime !== 'object') {
    throw new MeetingTimelineSdkError('meeting app observer scheduler requires a runtime object', {
      reason: 'missing_runtime',
    });
  }
  const sample = runtime.sample ?? runtime.tick;
  if (typeof sample !== 'function') {
    throw new MeetingTimelineSdkError('meeting app observer scheduler runtime must expose sample() or tick()', {
      reason: 'invalid_runtime',
    });
  }
  return runtime;
}

function callRuntimeSample(fn, runtime, input, sampleOptions, config, schedulerOptions) {
  if (runtimeInputMode(runtime, config, { ...schedulerOptions, ...sampleOptions }) === 'provider') {
    return fn(sampleOptions);
  }
  return fn(input, sampleOptions);
}

export function createMeetingAppObserverScheduler(runtimeInput, planOrPlatform = {}, options = {}) {
  const runtime = normalizeRuntime(runtimeInput);
  const config = options.config?.schema === MEETING_APP_OBSERVER_SCHEDULER_CONFIG_SCHEMA
    ? options.config
    : buildMeetingAppObserverSchedulerConfig(planOrPlatform, options);
  const timeout = options.setTimeout ?? globalThis.setTimeout;
  const clearTimeoutFn = options.clearTimeout ?? globalThis.clearTimeout;
  const interval = options.setInterval ?? globalThis.setInterval;
  const clearIntervalFn = options.clearInterval ?? globalThis.clearInterval;
  let state = normalizeSchedulerState(options.initialState ?? options.initial_state);
  let keepAliveTimer = null;
  const timers = new Map();
  let timerSeq = 0;
  const defaultInputProvider = firstNonEmpty(options.inputProvider, options.input_provider, options.input);

  function getState() {
    return {
      ...state,
      running: keepAliveTimer != null,
      pending_timer_count: timers.size,
    };
  }

  function reset(nextState = {}) {
    state = normalizeSchedulerState(nextState);
    return getState();
  }

  function sampleMethod() {
    return typeof runtime.sample === 'function' ? runtime.sample.bind(runtime) : runtime.tick.bind(runtime);
  }

  function trackSampleMethod() {
    if (typeof runtime.sampleTracks === 'function') return runtime.sampleTracks.bind(runtime);
    if (typeof runtime.tickTracks === 'function') return runtime.tickTracks.bind(runtime);
    return null;
  }

  async function sample(trigger = 'manual', input = undefined, sampleOptions = {}) {
    const normalized = normalizeMeetingAppObserverTrigger(trigger);
    const atMs = nowMs({ ...options, ...sampleOptions });
    const provider = firstNonEmpty(sampleOptions.inputProvider, sampleOptions.input_provider, defaultInputProvider);
    const sampleInput = input === undefined ? await inputFromProvider(provider, {}) : input;
    const mergedOptions = triggerSampleOptions(normalized, config, {
      ...sampleOptions,
      observedAtMs: atMs,
      observed_at_ms: atMs,
      capturedAtMs: atMs,
      captured_at_ms: atMs,
    });
    state = {
      ...state,
      trigger_count: state.trigger_count + 1,
      last_trigger: normalized,
      last_sample_at_ms: atMs,
      last_error: null,
    };
    try {
      const result = await callRuntimeSample(sampleMethod(), runtime, sampleInput, mergedOptions, config, options);
      let trackResult = null;
      if (trackTriggerEnabled(normalized, config, sampleOptions)) {
        const trackSample = trackSampleMethod();
        if (trackSample) {
          trackResult = await callRuntimeSample(trackSample, runtime, sampleInput, {
            ...mergedOptions,
            ...(sampleOptions.trackSampleOptions ?? {}),
            ...(sampleOptions.track_sample_options ?? {}),
          }, config, options);
        }
      }
      state = {
        ...state,
        sample_count: state.sample_count + 1,
        track_sample_count: state.track_sample_count + (trackResult ? 1 : 0),
        last_result: result,
        last_track_result: trackResult,
      };
      return compactObject({
        trigger: normalized,
        sampled: true,
        result,
        track_result: trackResult,
        options: mergedOptions,
        state: getState(),
      });
    } catch (error) {
      state = {
        ...state,
        last_error: error?.message ?? String(error),
      };
      throw error;
    }
  }

  function schedule(trigger = 'manual', input = undefined, scheduleOptions = {}) {
    if (typeof timeout !== 'function') {
      throw new MeetingTimelineSdkError('meeting app observer scheduler cannot schedule without setTimeout()', {
        reason: 'missing_set_timeout',
      });
    }
    const normalized = normalizeMeetingAppObserverTrigger(trigger);
    const id = `timer-${++timerSeq}`;
    const delayMs = triggerDelayMs(normalized, config, scheduleOptions);
    const handle = timeout(() => {
      timers.delete(id);
      return Promise.resolve(sample(normalized, input, scheduleOptions)).catch((error) => {
        state = {
          ...state,
          last_error: error?.message ?? String(error),
        };
      });
    }, delayMs);
    timers.set(id, { id, trigger: normalized, handle, delay_ms: delayMs });
    state = {
      ...state,
      scheduled_count: state.scheduled_count + 1,
      last_trigger: normalized,
    };
    return {
      scheduled: true,
      id,
      trigger: normalized,
      delay_ms: delayMs,
      state: getState(),
    };
  }

  function clearScheduled(id = null) {
    const ids = id == null ? [...timers.keys()] : [id];
    for (const timerId of ids) {
      const row = timers.get(timerId);
      if (!row) continue;
      if (typeof clearTimeoutFn === 'function') clearTimeoutFn(row.handle);
      timers.delete(timerId);
    }
    return getState();
  }

  function handleTrigger(trigger = 'manual', input = undefined, triggerOptions = {}) {
    if (triggerOptions.schedule === true || triggerOptions.deferred === true) {
      return schedule(trigger, input, triggerOptions);
    }
    return sample(trigger, input, triggerOptions);
  }

  function start(inputProvider = null, startOptions = {}) {
    if (keepAliveTimer != null) return getState();
    if (typeof interval !== 'function') {
      throw new MeetingTimelineSdkError('meeting app observer scheduler cannot start without setInterval()', {
        reason: 'missing_set_interval',
      });
    }
    const provider = inputProvider ?? defaultInputProvider;
    const intervalMs = numeric(firstNonEmpty(
      startOptions.intervalMs,
      startOptions.interval_ms,
      config.cadence.unchanged_observe_every_ms,
      config.cadence.fallback_poll_interval_ms,
      1_000,
    ), 1_000);
    keepAliveTimer = interval(() => {
      Promise.resolve(inputFromProvider(provider, {}))
        .then((input) => sample('unchanged_keep_alive', input, startOptions))
        .catch((error) => {
          state = {
            ...state,
            last_error: error?.message ?? String(error),
          };
        });
    }, intervalMs);
    state = {
      ...state,
      running: true,
    };
    if (startOptions.immediate === true) {
      Promise.resolve(inputFromProvider(provider, {}))
        .then((input) => sample('unchanged_keep_alive', input, startOptions))
        .catch((error) => {
          state = {
            ...state,
            last_error: error?.message ?? String(error),
          };
        });
    }
    return getState();
  }

  function stop() {
    if (keepAliveTimer != null && typeof clearIntervalFn === 'function') {
      clearIntervalFn(keepAliveTimer);
    }
    keepAliveTimer = null;
    clearScheduled();
    state = {
      ...state,
      running: false,
    };
    return getState();
  }

  return {
    config,
    sample,
    schedule,
    clearScheduled,
    handleTrigger,
    triggerChanged(input = undefined, triggerOptions = {}) {
      return handleTrigger(config.surface === 'native_detector' ? 'host_snapshot_change' : 'dom_mutation', input, triggerOptions);
    },
    triggerKeepAlive(input = undefined, triggerOptions = {}) {
      return handleTrigger('unchanged_keep_alive', input, triggerOptions);
    },
    triggerSpeakerCandidate(input = undefined, triggerOptions = {}) {
      return handleTrigger('active_speaker_candidate', input, triggerOptions);
    },
    triggerCandidateMissing(input = undefined, triggerOptions = {}) {
      return handleTrigger('meeting_candidate_missing', input, triggerOptions);
    },
    start,
    stop,
    getState,
    reset,
  };
}
