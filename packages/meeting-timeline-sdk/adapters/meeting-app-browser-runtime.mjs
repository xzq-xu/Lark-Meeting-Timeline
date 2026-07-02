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
    return { handled: false, reason: 'unsupported_message_type', type };
  }

  function start(startOptions = {}) {
    if (startOptions.lifecycle !== false && startOptions.installLifecycleHandlers !== false && startOptions.install_lifecycle_handlers !== false) {
      installLifecycleHandlers(startOptions);
    }
    return runtime.start(() => inputProvider(), startOptions);
  }

  function stop() {
    return runtime.stop();
  }

  function dispose() {
    const stopped = runtime.stop();
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
    dispose,
    getState() {
      return {
        ...runtime.getState(),
        browser_runtime: {
          lifecycle_installed: lifecycleInstalled,
          stop_events: listeners.map(([eventName]) => eventName),
        },
      };
    },
  };
}
