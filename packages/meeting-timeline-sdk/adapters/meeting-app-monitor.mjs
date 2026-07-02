import { MeetingTimelineSdkError, compactObject, normalizeAbsoluteMs } from '../index.mjs';
import { captureMeetingAppDomSnapshot } from './meeting-app-capture.mjs';

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function nowMs(options = {}) {
  const value = typeof options.now === 'function'
    ? options.now()
    : firstNonEmpty(options.now, options.nowMs, options.now_ms, Date.now());
  return normalizeAbsoluteMs(value, 'meeting_app_monitor_time');
}

function asArray(value) {
  return Array.isArray(value) ? value : value == null ? [] : [value];
}

function normalizeText(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function pickText(value) {
  const text = normalizeText(value);
  return text || undefined;
}

function audioBucket(value) {
  if (value == null || value === '') return undefined;
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return undefined;
  if (numeric >= 0.2) return 'active';
  if (numeric > 0) return 'low';
  return 'silent';
}

function controlSignature(control = {}) {
  return compactObject({
    role: pickText(control.role),
    label: pickText(control.label ?? control.ariaLabel ?? control.title ?? control.text),
    disabled: control.disabled === true ? true : undefined,
  });
}

function participantSignature(participant = {}) {
  return compactObject({
    id: pickText(participant.id),
    name: pickText(participant.name ?? participant.label ?? participant.ariaLabel ?? participant.text),
    speaking: participant.speaking === true || participant.isSpeaking === true ? true : undefined,
    audio: audioBucket(participant.audioLevel ?? participant.audio_level),
  });
}

function textSignature(text = {}) {
  return pickText(text.label ?? text.text ?? text.ariaLive);
}

export function meetingAppDomSnapshotSignature(snapshot = {}) {
  const page = snapshot.page ?? {};
  const dom = snapshot.dom ?? {};
  return JSON.stringify(compactObject({
    url: pickText(snapshot.url ?? page.url ?? dom.url),
    title: pickText(snapshot.title ?? page.title ?? dom.title),
    visible: page.documentVisible,
    controls: asArray(page.controls ?? page.buttons ?? dom.controls).slice(0, 40).map(controlSignature),
    participants: asArray(page.participants ?? page.tiles ?? dom.participants).slice(0, 40).map(participantSignature),
    texts: asArray(page.texts ?? dom.texts).slice(0, 20).map(textSignature).filter(Boolean),
  }));
}

function normalizeTarget(target) {
  if (!target || typeof target !== 'object') {
    throw new MeetingTimelineSdkError('meeting app monitor target is required', {
      reason: 'missing_monitor_target',
    });
  }
  if (typeof target.observeMeetingApp === 'function') {
    return (snapshot, options) => target.observeMeetingApp(snapshot, options);
  }
  if (typeof target.observeApp === 'function') {
    return (snapshot, options) => target.observeApp(snapshot, options);
  }
  if (typeof target.observe === 'function') {
    return (snapshot, options) => target.observe(snapshot, options);
  }
  throw new MeetingTimelineSdkError('meeting app monitor target must expose observeMeetingApp(), observeApp(), or observe()', {
    reason: 'invalid_monitor_target',
  });
}

function normalizeMonitorState(state = null) {
  return {
    running: false,
    sampleCount: Number(state?.sampleCount ?? state?.sample_count ?? 0),
    emitCount: Number(state?.emitCount ?? state?.emit_count ?? 0),
    skippedCount: Number(state?.skippedCount ?? state?.skipped_count ?? 0),
    lastObservedAtMs: state?.lastObservedAtMs ?? state?.last_observed_at_ms ?? null,
    lastEmittedAtMs: state?.lastEmittedAtMs ?? state?.last_emitted_at_ms ?? null,
    lastSignature: state?.lastSignature ?? state?.last_signature ?? null,
    lastSnapshot: state?.lastSnapshot ?? state?.last_snapshot ?? null,
    lastResult: state?.lastResult ?? state?.last_result ?? null,
    lastSkipReason: state?.lastSkipReason ?? state?.last_skip_reason ?? null,
  };
}

function intervalMs(options = {}, key, fallback) {
  const snake = key.replace(/[A-Z]/g, (char) => `_${char.toLowerCase()}`);
  const value = firstNonEmpty(options[key], options[snake], fallback);
  const numeric = Number(value);
  return Number.isFinite(numeric) ? Math.max(0, numeric) : fallback;
}

function shouldEmitSample(state, signature, atMs, options = {}) {
  if (options.force === true) return { emit: true, reason: 'forced' };
  if (!state.lastSignature) return { emit: true, reason: 'first_sample' };
  const changed = signature !== state.lastSignature;
  const sinceLastEmit = state.lastEmittedAtMs == null ? Infinity : atMs - state.lastEmittedAtMs;
  const minObserveMs = intervalMs(options, 'minObserveIntervalMs', 250);
  const unchangedEveryMs = intervalMs(options, 'unchangedObserveEveryMs', 1_000);
  const changedEveryMs = intervalMs(options, 'changedObserveEveryMs', minObserveMs);
  if (changed && sinceLastEmit >= changedEveryMs) return { emit: true, reason: 'changed' };
  if (!changed && sinceLastEmit >= unchangedEveryMs) return { emit: true, reason: 'keep_alive' };
  return {
    emit: false,
    reason: changed ? 'changed_throttled' : 'unchanged_throttled',
  };
}

function inputFromProvider(provider, fallbackInput) {
  if (typeof provider === 'function') return provider();
  return provider ?? fallbackInput ?? {};
}

export function createMeetingAppDomMonitor(target, options = {}) {
  const observeTarget = normalizeTarget(target);
  let state = normalizeMonitorState(options.initialState ?? options.initial_state);
  let timer = null;

  async function sample(input = {}, sampleOptions = {}) {
    const atMs = nowMs({
      ...options,
      ...sampleOptions,
    });
    const captureOptions = {
      ...(options.captureOptions ?? options.capture_options ?? {}),
      ...(sampleOptions.captureOptions ?? sampleOptions.capture_options ?? {}),
      platform: firstNonEmpty(
        sampleOptions.platform,
        sampleOptions.provider,
        options.platform,
        options.provider,
        sampleOptions.captureProfile,
        sampleOptions.capture_profile,
        options.captureProfile,
        options.capture_profile,
        sampleOptions.captureOptions?.platform,
        sampleOptions.capture_options?.platform,
        options.captureOptions?.platform,
        options.capture_options?.platform,
      ),
      captureProfile: firstNonEmpty(
        sampleOptions.captureProfile,
        sampleOptions.capture_profile,
        options.captureProfile,
        options.capture_profile,
        sampleOptions.captureOptions?.captureProfile,
        sampleOptions.capture_options?.captureProfile,
        sampleOptions.captureOptions?.capture_profile,
        sampleOptions.capture_options?.capture_profile,
        options.captureOptions?.captureProfile,
        options.capture_options?.captureProfile,
        options.captureOptions?.capture_profile,
        options.capture_options?.capture_profile,
      ),
      observedAtMs: atMs,
    };
    const observeOptions = {
      ...(options.observeOptions ?? options.observe_options ?? {}),
      ...(sampleOptions.observeOptions ?? sampleOptions.observe_options ?? {}),
      observedAtMs: atMs,
    };
    const snapshot = captureMeetingAppDomSnapshot(input, captureOptions);
    const signature = meetingAppDomSnapshotSignature(snapshot);
    const decision = shouldEmitSample(state, signature, atMs, {
      ...options,
      ...sampleOptions,
    });

    state = {
      ...state,
      sampleCount: state.sampleCount + 1,
      lastObservedAtMs: atMs,
      lastSnapshot: snapshot,
    };

    if (!decision.emit) {
      state = {
        ...state,
        skippedCount: state.skippedCount + 1,
        lastSkipReason: decision.reason,
      };
      return {
        emitted: false,
        skipped: true,
        reason: decision.reason,
        signature,
        snapshot,
        state: getState(),
      };
    }

    const result = await observeTarget(snapshot, observeOptions);
    state = {
      ...state,
      emitCount: state.emitCount + 1,
      lastEmittedAtMs: atMs,
      lastSignature: signature,
      lastResult: result,
      lastSkipReason: null,
    };
    return {
      emitted: true,
      skipped: false,
      reason: decision.reason,
      signature,
      snapshot,
      result,
      state: getState(),
    };
  }

  function getState() {
    return {
      ...state,
      running: timer != null,
    };
  }

  function reset(nextState = null) {
    state = normalizeMonitorState(nextState);
    return getState();
  }

  function stop() {
    if (timer != null) {
      clearInterval(timer);
      timer = null;
    }
    return getState();
  }

  function start(inputProvider = null, startOptions = {}) {
    if (timer != null) return getState();
    const scheduleMs = intervalMs({
      ...options,
      ...startOptions,
    }, 'sampleIntervalMs', 1_000);
    const provider = inputProvider ?? options.inputProvider ?? options.input_provider ?? options.input;
    timer = setInterval(() => {
      Promise.resolve(inputFromProvider(provider, {}))
        .then((input) => sample(input, startOptions))
        .catch((error) => {
          state = {
            ...state,
            lastResult: {
              ok: false,
              error: error?.message ?? String(error),
            },
          };
        });
    }, scheduleMs);
    if (startOptions.immediate !== false && options.immediate !== false) {
      Promise.resolve(inputFromProvider(provider, {}))
        .then((input) => sample(input, startOptions))
        .catch((error) => {
          state = {
            ...state,
            lastResult: {
              ok: false,
              error: error?.message ?? String(error),
            },
          };
        });
    }
    return getState();
  }

  return {
    sample,
    tick: sample,
    start,
    stop,
    getState,
    reset,
  };
}
