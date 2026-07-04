import { createMeetingAppBrowserRuntime } from './meeting-app-browser-runtime.mjs';

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

function extensionRuntime(options = {}) {
  const chromeRuntime = options.chromeRuntime
    ?? options.chrome_runtime
    ?? options.chrome?.runtime
    ?? globalValue('chrome')?.runtime;
  const browserRuntime = options.browserRuntime
    ?? options.browser_runtime
    ?? options.browser?.runtime
    ?? globalValue('browser')?.runtime;
  return options.extensionRuntime
    ?? options.extension_runtime
    ?? chromeRuntime
    ?? browserRuntime;
}

function messageType(message = {}) {
  return String(message?.type ?? message?.action ?? message?.kind ?? '').trim();
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

const DEFAULT_MESSAGE_PREFIXES = Object.freeze([
  'meeting_timeline',
  'meeting-timeline',
]);

const DEFAULT_DIRECT_MESSAGE_TYPES = Object.freeze([
  'sample',
  'insert_mark',
  'insert_marks',
  'insertAnnotation',
  'annotation',
  'provider_event',
  'ingest_provider',
  'stop',
  'flush_mutations',
  'sample_tracks',
  'track_sample',
  'start_tracks',
  'stop_tracks',
  'observe_tracks',
  'preview_tracks',
]);

function messagePrefixes(options = {}) {
  return asArray(firstNonEmpty(
    options.messagePrefixes,
    options.message_prefixes,
    DEFAULT_MESSAGE_PREFIXES,
  )).filter(Boolean).map(String);
}

function directMessageTypes(options = {}) {
  return asArray(firstNonEmpty(
    options.directMessageTypes,
    options.direct_message_types,
    DEFAULT_DIRECT_MESSAGE_TYPES,
  )).filter(Boolean).map(String);
}

function isMeetingTimelineMessage(message = {}, options = {}) {
  const type = messageType(message);
  if (!type) return false;
  if (type === 'meeting_timeline.response' || type.endsWith('.response') || type.endsWith('_response')) return false;
  if (typeof options.messageFilter === 'function') return options.messageFilter(message, options) !== false;
  if (typeof options.message_filter === 'function') return options.message_filter(message, options) !== false;
  if (directMessageTypes(options).includes(type)) return true;
  return messagePrefixes(options).some((prefix) => type === prefix || type.startsWith(`${prefix}.`) || type.startsWith(`${prefix}_`));
}

function normalizeError(error) {
  return {
    name: error?.name ?? 'Error',
    message: error?.message ?? String(error),
  };
}

function windowOrigin(win) {
  const origin = win?.location?.origin;
  if (origin) return origin;
  const href = win?.location?.href;
  if (!href) return undefined;
  try {
    return new URL(href).origin;
  } catch {
    return undefined;
  }
}

function originAllowed(event = {}, options = {}) {
  const allowed = asArray(options.allowedOrigins ?? options.allowed_origins).filter(Boolean).map(String);
  if (allowed.includes('*')) return true;
  const origin = event.origin;
  if (allowed.length > 0) return allowed.includes(String(origin));
  const sameOrigin = windowOrigin(browserWindow(options));
  if (!origin || !sameOrigin) return true;
  return origin === sameOrigin;
}

function extensionOnMessage(runtimeApi) {
  return runtimeApi?.onMessage ?? runtimeApi?.runtime?.onMessage;
}

function sendResponseSafely(sendResponse, payload) {
  if (typeof sendResponse !== 'function') return false;
  try {
    sendResponse(payload);
    return true;
  } catch {
    return false;
  }
}

export function createMeetingAppContentScriptBridge(clientOrOptions, options = {}) {
  const runtime = options.runtime ?? createMeetingAppBrowserRuntime(clientOrOptions, options);
  const listeners = [];
  let extensionInstalled = false;
  let windowInstalled = false;

  async function dispatchMessage(message = {}, messageOptions = {}) {
    if (!isMeetingTimelineMessage(message, options)) {
      return { handled: false, reason: 'unsupported_message_type', type: messageType(message) };
    }
    return runtime.handleMessage(message, {
      ...(options.messageOptions ?? options.message_options ?? {}),
      ...messageOptions,
    });
  }

  function installExtensionMessaging(installOptions = {}) {
    if (extensionInstalled) return { installed: false, reason: 'already_installed' };
    const runtimeApi = extensionRuntime({ ...options, ...installOptions });
    const onMessage = extensionOnMessage(runtimeApi);
    if (!onMessage?.addListener) return { installed: false, reason: 'missing_extension_runtime' };
    const listener = (message, sender, sendResponse) => {
      if (!isMeetingTimelineMessage(message, options)) return false;
      dispatchMessage(message, { sender }).then((result) => {
        sendResponseSafely(sendResponse, result);
      }).catch((error) => {
        sendResponseSafely(sendResponse, {
          handled: false,
          reason: 'message_handler_error',
          error: normalizeError(error),
        });
      });
      return true;
    };
    onMessage.addListener(listener);
    listeners.push({
      kind: 'extension',
      remove() {
        onMessage.removeListener?.(listener);
      },
    });
    extensionInstalled = true;
    return { installed: true, kind: 'extension' };
  }

  function installWindowMessaging(installOptions = {}) {
    if (windowInstalled) return { installed: false, reason: 'already_installed' };
    const win = browserWindow({ ...options, ...installOptions });
    if (!win?.addEventListener) return { installed: false, reason: 'missing_window' };
    const listener = (event = {}) => {
      const message = event.data;
      if (!isMeetingTimelineMessage(message, options)) return;
      if (!originAllowed(event, { ...options, ...installOptions })) return;
      dispatchMessage(message, { event }).then((result) => {
        if (
          installOptions.postWindowResponses === false
          || installOptions.post_window_responses === false
          || options.postWindowResponses === false
          || options.post_window_responses === false
        ) return;
        const target = event.source ?? win;
        const targetOrigin = event.origin || windowOrigin(win) || '*';
        target?.postMessage?.({
          type: 'meeting_timeline.response',
          request_type: messageType(message),
          request_id: message?.id ?? message?.requestId ?? message?.request_id,
          result,
        }, targetOrigin);
      }).catch((error) => {
        const target = event.source ?? win;
        const targetOrigin = event.origin || windowOrigin(win) || '*';
        target?.postMessage?.({
          type: 'meeting_timeline.response',
          request_type: messageType(message),
          request_id: message?.id ?? message?.requestId ?? message?.request_id,
          result: {
            handled: false,
            reason: 'message_handler_error',
            error: normalizeError(error),
          },
        }, targetOrigin);
      });
    };
    win.addEventListener('message', listener);
    listeners.push({
      kind: 'window',
      remove() {
        win.removeEventListener?.('message', listener);
      },
    });
    windowInstalled = true;
    return { installed: true, kind: 'window' };
  }

  function removeMessaging(kind) {
    const keep = [];
    let removed = 0;
    for (const item of listeners.splice(0)) {
      if (!kind || item.kind === kind) {
        item.remove();
        removed += 1;
        if (item.kind === 'extension') extensionInstalled = false;
        if (item.kind === 'window') windowInstalled = false;
      } else {
        keep.push(item);
      }
    }
    listeners.push(...keep);
    return { removed };
  }

  function start(startOptions = {}) {
    const extensionMessaging = firstNonEmpty(
      startOptions.extensionMessaging,
      startOptions.extension_messaging,
      options.extensionMessaging,
      options.extension_messaging,
      true,
    );
    const windowMessaging = firstNonEmpty(
      startOptions.windowMessaging,
      startOptions.window_messaging,
      options.windowMessaging,
      options.window_messaging,
      false,
    );
    const installs = [];
    if (extensionMessaging !== false) installs.push(installExtensionMessaging(startOptions));
    if (windowMessaging === true) installs.push(installWindowMessaging(startOptions));
    const runtimeState = startOptions.startRuntime === false || startOptions.start_runtime === false
      ? runtime.getState?.()
      : runtime.start(startOptions);
    return {
      running: runtimeState?.running ?? runtimeState?.monitor?.running ?? true,
      installs,
      runtime: runtimeState,
    };
  }

  function stop() {
    const state = runtime.stop();
    return {
      runtime: state,
      messaging: removeMessaging(),
    };
  }

  function dispose() {
    const messaging = removeMessaging();
    const runtimeState = typeof runtime.dispose === 'function' ? runtime.dispose() : runtime.stop();
    return {
      runtime: runtimeState,
      messaging,
    };
  }

  function getState() {
    return {
      runtime: runtime.getState?.(),
      messaging: {
        extension_installed: extensionInstalled,
        window_installed: windowInstalled,
        listener_count: listeners.length,
      },
    };
  }

  return {
    runtime,
    dispatchMessage,
    start,
    stop,
    dispose,
    installExtensionMessaging,
    installWindowMessaging,
    removeMessaging,
    getState,
  };
}

export function installMeetingAppContentScriptBridge(clientOrOptions, options = {}) {
  const bridge = createMeetingAppContentScriptBridge(clientOrOptions, options);
  bridge.start(options.startOptions ?? options.start_options ?? {});
  return bridge;
}

export default createMeetingAppContentScriptBridge;
