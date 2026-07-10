import { captureMeetingAppDomSnapshot } from './meeting-app-capture.mjs';
import { createMeetingAppTimelineObserver, normalizeMeetingAppSnapshot } from './meeting-apps.mjs';

export const MEETING_APP_LOCAL_CONTENT_RUNTIME_SCHEMA = 'meeting_app_local_content_runtime';
export const MEETING_APP_LOCAL_CONTENT_RUNTIME_SCHEMA_VERSION = 1;

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function runtimeApi(options = {}) {
  return options.extensionRuntime
    ?? options.extension_runtime
    ?? options.chrome?.runtime
    ?? globalThis.chrome?.runtime
    ?? options.browser?.runtime
    ?? globalThis.browser?.runtime;
}

function messageType(message = {}) {
  return String(message.type ?? message.action ?? message.kind ?? '').trim();
}

function messagePayload(message = {}) {
  return message.payload ?? message.data ?? message.input ?? message;
}

function supportedMessage(type) {
  return type.startsWith('meeting_timeline.')
    || type.startsWith('meeting_timeline_')
    || ['sample', 'insert_mark', 'insert_marks', 'stop', 'preflight_current_window'].includes(type);
}

function sendResponse(sendResponse, value) {
  try {
    sendResponse?.(value);
  } catch {
    // The sender may have closed while a sample was running.
  }
}

function captureInput(options = {}, captureOptions = {}) {
  const win = options.window ?? globalThis;
  const document = options.document ?? win.document ?? globalThis.document;
  const location = options.location ?? win.location ?? globalThis.location;
  const platform = firstNonEmpty(captureOptions.platform, options.platform, options.browser_runtime_preset, options.runtimePreset);
  const observedAtMs = firstNonEmpty(captureOptions.observedAtMs, captureOptions.observed_at_ms, Date.now());
  return captureMeetingAppDomSnapshot({
    window: win,
    document,
    location,
    platform,
    url: location?.href,
    title: document?.title,
  }, {
    ...(options.captureOptions ?? options.capture_options ?? {}),
    ...captureOptions,
    platform,
    captureProfile: firstNonEmpty(captureOptions.captureProfile, captureOptions.capture_profile, platform),
    source: firstNonEmpty(captureOptions.source, options.source, 'meeting_app_local_content_runtime'),
    observedAtMs,
  });
}

function debounceValue(options = {}) {
  const value = firstNonEmpty(options.mutationDebounceMs, options.mutation_debounce_ms, 180);
  const number = Number(value);
  return Number.isFinite(number) ? Math.max(0, number) : 180;
}

function speakerFollowupValue(options = {}) {
  const value = firstNonEmpty(
    options.speakerStableFollowupMs,
    options.speaker_stable_followup_ms,
    options.speakerOptions?.minStableMs,
    options.speaker_options?.min_stable_ms,
    700,
  );
  const number = Number(value);
  return Number.isFinite(number) ? Math.max(0, number) : 700;
}

function fallbackIntervalValue(options = {}) {
  const value = firstNonEmpty(options.sampleIntervalMs, options.sample_interval_ms, 10_000);
  const number = Number(value);
  return Number.isFinite(number) ? Math.max(250, number) : 10_000;
}

export function createMeetingAppLocalContentRuntime(client, options = {}) {
  const observer = createMeetingAppTimelineObserver(client, {
    source: options.source ?? 'meeting_app_extension',
    applyOptions: {
      speakerAsAnnotation: true,
      participantAsAnnotation: false,
      ...(options.applyOptions ?? options.apply_options ?? {}),
    },
    speakerOptions: {
      minStableMs: 650,
      switchStableMs: 700,
      endIdleMs: 1_500,
      ...(options.speakerOptions ?? options.speaker_options ?? {}),
    },
  });
  let running = false;
  let mutationObserver = null;
  let mutationTimer = null;
  let speakerFollowupTimer = null;
  let fallbackTimer = null;
  let sampleCount = 0;
  let lastSample = null;
  let lastError = null;
  const listeners = [];

  function clearTimers() {
    if (mutationTimer != null) clearTimeout(mutationTimer);
    if (speakerFollowupTimer != null) clearTimeout(speakerFollowupTimer);
    if (fallbackTimer != null) clearInterval(fallbackTimer);
    mutationTimer = null;
    speakerFollowupTimer = null;
    fallbackTimer = null;
  }

  async function sample(sampleOptions = {}) {
    try {
      const snapshot = captureInput(options, sampleOptions);
      const result = await observer.observe(snapshot, sampleOptions);
      sampleCount += 1;
      lastSample = { captured_at_ms: snapshot.observedAtMs ?? snapshot.observed_at_ms, snapshot, result };
      lastError = null;
      return lastSample;
    } catch (error) {
      lastError = String(error?.message ?? error);
      return { error: lastError, captured_at_ms: Date.now() };
    }
  }

  function scheduleMutationSample() {
    if (mutationTimer != null) clearTimeout(mutationTimer);
    mutationTimer = setTimeout(() => {
      mutationTimer = null;
      sample({ trigger: 'mutation' }).then(() => {
        if (speakerFollowupTimer != null) clearTimeout(speakerFollowupTimer);
        speakerFollowupTimer = setTimeout(() => {
          speakerFollowupTimer = null;
          sample({ trigger: 'speaker_stability_followup' });
        }, speakerFollowupValue(options));
      });
    }, debounceValue(options));
  }

  function installMutationObserver() {
    const win = options.window ?? globalThis;
    const document = options.document ?? win.document ?? globalThis.document;
    const Observer = options.MutationObserver ?? win.MutationObserver ?? globalThis.MutationObserver;
    const root = document?.body ?? document?.documentElement;
    if (!Observer || !root) return { installed: false, reason: 'missing_mutation_observer_or_root' };
    mutationObserver = new Observer(scheduleMutationSample);
    mutationObserver.observe(root, { childList: true, subtree: true, attributes: true, characterData: true });
    return { installed: true };
  }

  function preflight() {
    const snapshot = captureInput(options, { trigger: 'preflight' });
    const normalized = normalizeMeetingAppSnapshot(snapshot, {
      platform: firstNonEmpty(options.platform, options.browser_runtime_preset, options.runtimePreset),
      source: options.source ?? 'meeting_app_local_content_runtime',
    });
    const meetingId = normalized.meeting_id ?? normalized.meeting?.meeting_id;
    return {
      accepted: Boolean(meetingId && snapshot.capture?.control_count > 0),
      platform: normalized.platform ?? snapshot.platform,
      meeting_id: meetingId,
      captured_at_ms: snapshot.observedAtMs ?? snapshot.observed_at_ms,
      control_count: snapshot.capture?.control_count ?? 0,
      participant_count: snapshot.capture?.participant_count ?? 0,
      active_speaker_count: snapshot.capture?.active_speaker_count ?? 0,
      normalized,
      snapshot,
    };
  }

  async function handleMessage(message = {}) {
    const type = messageType(message);
    const payload = messagePayload(message);
    if (!supportedMessage(type)) return { handled: false, reason: 'unsupported_message_type', type };
    if (['meeting_timeline.sample', 'meeting_timeline_sample', 'sample'].includes(type)) {
      return { handled: true, action: 'sample', result: await sample(payload) };
    }
    if (['meeting_timeline.insert_mark', 'meeting_timeline_insert_mark', 'insert_mark'].includes(type)) {
      return { handled: true, action: 'insertMark', result: await client.insertMark(payload.mark ?? payload.annotation ?? payload) };
    }
    if (['meeting_timeline.insert_marks', 'meeting_timeline_insert_marks', 'insert_marks'].includes(type)) {
      return { handled: true, action: 'insertMarks', result: await client.insertMarks(payload.marks ?? payload.annotations ?? payload.items ?? payload) };
    }
    if (['meeting_timeline.preflight_current_window', 'meeting_timeline_preflight_current_window', 'preflight_current_window'].includes(type)) {
      return { handled: true, action: 'preflightCurrentWindow', result: preflight() };
    }
    if (['meeting_timeline.stop', 'meeting_timeline_stop', 'stop'].includes(type)) {
      return { handled: true, action: 'stop', result: stop() };
    }
    return { handled: false, reason: 'unsupported_message_type', type };
  }

  function installMessaging() {
    const extension = runtimeApi(options);
    if (extension?.onMessage?.addListener) {
      const listener = (message, sender, respond) => {
        if (!supportedMessage(messageType(message))) return false;
        handleMessage(message, { sender }).then((result) => sendResponse(respond, result));
        return true;
      };
      extension.onMessage.addListener(listener);
      listeners.push(() => extension.onMessage.removeListener?.(listener));
    }
    const win = options.window ?? globalThis;
    if (options.windowMessaging !== false && win?.addEventListener) {
      const listener = (event = {}) => {
        if (!supportedMessage(messageType(event.data))) return;
        handleMessage(event.data).then((result) => {
          event.source?.postMessage?.({
            type: 'meeting_timeline.response',
            request_type: messageType(event.data),
            request_id: event.data?.request_id ?? event.data?.id,
            result,
          }, event.origin || '*');
        });
      };
      win.addEventListener('message', listener);
      listeners.push(() => win.removeEventListener?.('message', listener));
    }
  }

  function start() {
    if (running) return getState();
    running = true;
    installMutationObserver();
    installMessaging();
    sample({ trigger: 'runtime_start' });
    fallbackTimer = setInterval(() => sample({ trigger: 'fallback_interval' }), fallbackIntervalValue(options));
    return getState();
  }

  function stop() {
    running = false;
    clearTimers();
    mutationObserver?.disconnect?.();
    mutationObserver = null;
    for (const remove of listeners.splice(0)) remove();
    return getState();
  }

  function getState() {
    return {
      type: 'meeting_app_local_content_runtime_state',
      schema: MEETING_APP_LOCAL_CONTENT_RUNTIME_SCHEMA,
      schema_version: MEETING_APP_LOCAL_CONTENT_RUNTIME_SCHEMA_VERSION,
      running,
      sample_count: sampleCount,
      last_captured_at_ms: lastSample?.captured_at_ms,
      last_signal_types: lastSample?.result?.signals?.map((signal) => signal.type) ?? [],
      last_error: lastError,
      mutation_observer_installed: mutationObserver != null,
    };
  }

  return {
    type: MEETING_APP_LOCAL_CONTENT_RUNTIME_SCHEMA,
    schema: MEETING_APP_LOCAL_CONTENT_RUNTIME_SCHEMA,
    schema_version: MEETING_APP_LOCAL_CONTENT_RUNTIME_SCHEMA_VERSION,
    client,
    options,
    observer,
    sample,
    preflight,
    handleMessage,
    start,
    stop,
    getState,
  };
}

export function installMeetingAppLocalContentRuntime(client, options = {}) {
  const runtime = createMeetingAppLocalContentRuntime(client, options);
  runtime.start();
  return runtime;
}

export default createMeetingAppLocalContentRuntime;
