import { createMeetingAppTimelineRuntime } from './meeting-app-runtime.mjs';

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function globalValue(name) {
  try {
    return globalThis?.[name];
  } catch {
    return undefined;
  }
}

function browserWindow(options = {}) {
  return options.window ?? options.win ?? globalValue('window');
}

function browserDocument(options = {}) {
  const win = browserWindow(options);
  return options.document ?? options.doc ?? win?.document ?? globalValue('document');
}

function browserLocation(options = {}) {
  const win = browserWindow(options);
  const doc = browserDocument(options);
  return options.location ?? win?.location ?? doc?.location ?? globalValue('location');
}

function browserNavigator(options = {}) {
  const win = browserWindow(options);
  return options.navigator ?? win?.navigator ?? globalValue('navigator');
}

function browserName(options = {}) {
  const nav = browserNavigator(options);
  return firstNonEmpty(
    options.browserName,
    options.browser_name,
    nav?.userAgentData?.brands?.[0]?.brand,
    nav?.userAgent,
  );
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
    ?? globalValue('MutationObserver');
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

function mutationEnabled(options = {}) {
  return [
    options.mutationObserver,
    options.mutation_observer,
    options.observeMutations,
    options.observe_mutations,
  ].some((value) => value === true);
}

function debounceMs(options = {}, fallback = 150) {
  const value = firstNonEmpty(options.debounceMs, options.debounce_ms, options.mutationDebounceMs, options.mutation_debounce_ms, fallback);
  const numeric = Number(value);
  return Number.isFinite(numeric) ? Math.max(0, numeric) : fallback;
}

function normalizeStopEvents(options = {}) {
  const value = options.stopEvents ?? options.stop_events ?? ['pagehide', 'beforeunload'];
  return Array.isArray(value) ? value : [value].filter(Boolean);
}

function messageType(message = {}) {
  return String(message?.type ?? message?.action ?? message?.kind ?? '').trim();
}

function messagePayload(message = {}) {
  return message?.payload ?? message?.data ?? message;
}

export function meetingAppBrowserInput(options = {}) {
  const win = browserWindow(options);
  const doc = browserDocument(options);
  const location = browserLocation(options);
  return Object.fromEntries(Object.entries({
    window: win,
    document: doc,
    location,
    url: options.url ?? location?.href,
    title: options.title ?? doc?.title,
  }).filter(([, value]) => value !== undefined));
}

export function createMeetingAppBrowserRuntime(clientOrOptions, options = {}) {
  const inputOptions = {
    window: options.window,
    win: options.win,
    document: options.document,
    doc: options.doc,
    location: options.location,
    navigator: options.navigator,
  };
  const runtime = createMeetingAppTimelineRuntime(clientOrOptions, {
    source: options.source ?? 'browser_content_script',
    ...options,
    captureOptions: {
      browserName: browserName(options),
      ...(options.captureOptions ?? {}),
      ...(options.capture_options ?? {}),
    },
  });
  const target = eventTarget(options);
  const stopEvents = normalizeStopEvents(options);
  const listeners = [];
  let lifecycleInstalled = false;
  let mutationObserver = null;
  let mutationTimer = null;
  let mutationPendingCount = 0;
  let mutationSampleCount = 0;
  let mutationLastResult = null;
  let mutationLastError = null;
  let mutationLastObservedAtMs = null;
  let mutationLastOptions = null;

  function inputProvider(extra = {}) {
    return meetingAppBrowserInput({
      ...inputOptions,
      ...extra,
    });
  }

  function installLifecycleHandlers(lifecycleOptions = {}) {
    if (lifecycleInstalled || !target?.addEventListener) {
      lifecycleInstalled = true;
      return { installed: false, reason: target?.addEventListener ? 'already_installed' : 'missing_event_target' };
    }
    const handler = () => runtime.stop();
    for (const eventName of normalizeStopEvents({
      ...options,
      ...lifecycleOptions,
      stopEvents: lifecycleOptions.stopEvents ?? lifecycleOptions.stop_events ?? stopEvents,
    })) {
      target.addEventListener(eventName, handler);
      listeners.push([eventName, handler]);
    }
    lifecycleInstalled = true;
    return { installed: true, events: listeners.map(([eventName]) => eventName) };
  }

  function removeLifecycleHandlers() {
    if (target?.removeEventListener) {
      for (const [eventName, handler] of listeners.splice(0)) {
        target.removeEventListener(eventName, handler);
      }
    } else {
      listeners.splice(0);
    }
    lifecycleInstalled = false;
    return { installed: false };
  }

  function mutationState() {
    return {
      installed: mutationObserver != null,
      pending_count: mutationPendingCount,
      sample_count: mutationSampleCount,
      last_observed_at_ms: mutationLastObservedAtMs,
      last_result: mutationLastResult,
      last_error: mutationLastError,
      debounce_ms: mutationLastOptions ? debounceMs(mutationLastOptions) : null,
    };
  }

  function clearMutationTimer() {
    if (mutationTimer != null) {
      clearTimeout(mutationTimer);
      mutationTimer = null;
    }
  }

  async function flushMutationObserver(flushOptions = {}) {
    clearMutationTimer();
    if (mutationPendingCount <= 0 && flushOptions.force !== true) {
      return {
        flushed: false,
        reason: 'no_pending_mutations',
        state: mutationState(),
      };
    }
    mutationPendingCount = 0;
    const sampleOptions = {
      ...(mutationLastOptions?.sampleOptions ?? mutationLastOptions?.sample_options ?? {}),
      ...(flushOptions.sampleOptions ?? flushOptions.sample_options ?? {}),
      ...flushOptions,
    };
    try {
      mutationLastObservedAtMs = Date.now();
      const result = await runtime.sample(inputProvider(), sampleOptions);
      mutationSampleCount += 1;
      mutationLastResult = result;
      mutationLastError = null;
      return {
        flushed: true,
        result,
        state: mutationState(),
      };
    } catch (error) {
      mutationLastError = error?.message ?? String(error);
      mutationLastResult = null;
      return {
        flushed: false,
        reason: 'sample_error',
        error: mutationLastError,
        state: mutationState(),
      };
    }
  }

  function scheduleMutationSample(scheduleOptions = {}) {
    clearMutationTimer();
    const waitMs = debounceMs(scheduleOptions);
    mutationTimer = setTimeout(() => {
      flushMutationObserver().catch((error) => {
        mutationLastError = error?.message ?? String(error);
      });
    }, waitMs);
    return mutationState();
  }

  function installMutationObserver(mutationOptions = {}) {
    if (mutationObserver) return { installed: false, reason: 'already_installed', state: mutationState() };
    const merged = {
      ...options,
      ...mutationOptions,
    };
    const Observer = mutationObserverCtor(merged);
    const root = mutationRoot(merged);
    if (!Observer) return { installed: false, reason: 'missing_mutation_observer', state: mutationState() };
    if (!root) return { installed: false, reason: 'missing_mutation_root', state: mutationState() };
    mutationLastOptions = merged;
    mutationObserver = new Observer((records = []) => {
      mutationPendingCount += Math.max(1, Array.isArray(records) ? records.length : 1);
      scheduleMutationSample(mutationLastOptions);
    });
    mutationObserver.observe(root, mutationObserveOptions(merged));
    return {
      installed: true,
      root,
      options: mutationObserveOptions(merged),
      state: mutationState(),
    };
  }

  function removeMutationObserver() {
    clearMutationTimer();
    if (mutationObserver?.disconnect) mutationObserver.disconnect();
    mutationObserver = null;
    mutationPendingCount = 0;
    mutationLastOptions = null;
    return mutationState();
  }

  async function handleMessage(message = {}, messageOptions = {}) {
    const type = messageType(message);
    const payload = messagePayload(message);
    if ([
      'meeting_timeline.sample',
      'meeting_timeline_sample',
      'sample',
    ].includes(type)) {
      return { handled: true, action: 'sample', result: await runtime.sample(inputProvider(payload), messageOptions) };
    }
    if ([
      'meeting_timeline.insert_mark',
      'meeting_timeline_insert_mark',
      'insert_mark',
      'insertAnnotation',
      'annotation',
    ].includes(type)) {
      const mark = payload.mark ?? payload.annotation ?? payload;
      return { handled: true, action: 'insertMark', result: await runtime.insertMark(mark, messageOptions) };
    }
    if ([
      'meeting_timeline.insert_marks',
      'meeting_timeline_insert_marks',
      'insert_marks',
      'insertMarks',
    ].includes(type)) {
      const marks = payload.marks ?? payload.annotations ?? payload.items ?? payload;
      return { handled: true, action: 'insertMarks', result: await runtime.insertMarks(marks, messageOptions) };
    }
    if ([
      'meeting_timeline.provider_event',
      'meeting_timeline_provider_event',
      'provider_event',
      'ingest_provider',
    ].includes(type)) {
      return {
        handled: true,
        action: 'ingestProvider',
        result: await runtime.ingestProvider(payload.platform ?? payload.provider, payload.body ?? payload.event ?? payload.payload, messageOptions),
      };
    }
    if ([
      'meeting_timeline.stop',
      'meeting_timeline_stop',
      'stop',
    ].includes(type)) {
      return { handled: true, action: 'stop', result: stop() };
    }
    if ([
      'meeting_timeline.flush_mutations',
      'meeting_timeline_flush_mutations',
      'flush_mutations',
    ].includes(type)) {
      return { handled: true, action: 'flushMutationObserver', result: await flushMutationObserver(messageOptions) };
    }
    return { handled: false, reason: 'unsupported_message_type', type };
  }

  function start(startOptions = {}) {
    if (startOptions.lifecycle !== false && startOptions.installLifecycleHandlers !== false && startOptions.install_lifecycle_handlers !== false) {
      installLifecycleHandlers(startOptions);
    }
    if (mutationEnabled({
      ...options,
      ...startOptions,
    })) {
      installMutationObserver(startOptions);
    }
    return runtime.start(() => inputProvider(), startOptions);
  }

  function stop() {
    if (options.keepMutationObserverOnStop !== true && options.keep_mutation_observer_on_stop !== true) {
      removeMutationObserver();
    }
    return runtime.stop();
  }

  function dispose() {
    const stopped = runtime.stop();
    removeMutationObserver();
    removeLifecycleHandlers();
    return stopped;
  }

  return {
    ...runtime,
    inputProvider,
    start,
    stop,
    sample(sampleOptions = {}) {
      return runtime.sample(inputProvider(), sampleOptions);
    },
    tick(sampleOptions = {}) {
      return runtime.tick(inputProvider(), sampleOptions);
    },
    handleMessage,
    installLifecycleHandlers,
    removeLifecycleHandlers,
    installMutationObserver,
    removeMutationObserver,
    flushMutationObserver,
    dispose,
    getState() {
      return {
        ...runtime.getState(),
        browser_runtime: {
          lifecycle_installed: lifecycleInstalled,
          stop_events: listeners.map(([eventName]) => eventName),
          mutation_observer: mutationState(),
        },
      };
    },
  };
}
