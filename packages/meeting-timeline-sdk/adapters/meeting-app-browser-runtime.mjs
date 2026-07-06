import { createMeetingAppTimelineRuntime } from './meeting-app-runtime.mjs';
import {
  meetingAppDomCaptureProfile,
} from './meeting-app-capture.mjs';

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

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

function compactStrings(values = []) {
  const seen = new Set();
  const output = [];
  for (const value of asArray(values).flatMap(asArray)) {
    const text = String(value ?? '').trim();
    if (!text || seen.has(text)) continue;
    seen.add(text);
    output.push(text);
  }
  return output;
}

const COMMON_MUTATION_IGNORE_SELECTORS = Object.freeze([
  '.caption-line',
  '.captions-line',
  '[data-caption-line]',
  '[data-transcript-line]',
  '[data-chat-message]',
  '[data-message-id]',
  '[class*="caption" i]',
  '[class*="transcript" i]',
  '[class*="chat" i]',
  '[aria-label*="caption" i]',
  '[aria-label*="subtitles" i]',
  '[aria-label*="字幕"]',
  '[aria-label*="聊天"]',
]);

function runtimePresetForPlatform(platform, timing = {}) {
  return Object.freeze({
    platform,
    captureOptions: Object.freeze({ platform }),
    observeMutations: true,
    observe_mutations: true,
    mutationDebounceMs: timing.mutationDebounceMs ?? 150,
    mutation_debounce_ms: timing.mutationDebounceMs ?? 150,
    speakerStableFollowupMs: timing.speakerStableFollowupMs ?? 300,
    speaker_stable_followup_ms: timing.speakerStableFollowupMs ?? 300,
    sampleIntervalMs: timing.sampleIntervalMs ?? 10_000,
    sample_interval_ms: timing.sampleIntervalMs ?? 10_000,
    unchangedObserveEveryMs: timing.unchangedObserveEveryMs ?? 10_000,
    unchanged_observe_every_ms: timing.unchangedObserveEveryMs ?? 10_000,
    mutationTrackSelectors: Object.freeze([]),
    mutation_track_selectors: Object.freeze([]),
    mutationIgnoreSelectors: COMMON_MUTATION_IGNORE_SELECTORS,
    mutation_ignore_selectors: COMMON_MUTATION_IGNORE_SELECTORS,
  });
}

export const MEETING_APP_BROWSER_RUNTIME_PRESETS = Object.freeze({
  google_meet: runtimePresetForPlatform('google_meet', {
    mutationDebounceMs: 150,
    speakerStableFollowupMs: 300,
  }),
  microsoft_teams: runtimePresetForPlatform('microsoft_teams', {
    mutationDebounceMs: 180,
    speakerStableFollowupMs: 350,
  }),
  zoom: runtimePresetForPlatform('zoom', {
    mutationDebounceMs: 180,
    speakerStableFollowupMs: 350,
  }),
  lark: runtimePresetForPlatform('lark', {
    mutationDebounceMs: 180,
    speakerStableFollowupMs: 350,
  }),
  webex: runtimePresetForPlatform('webex', {
    mutationDebounceMs: 180,
    speakerStableFollowupMs: 350,
  }),
});

function runtimePresetProfileSelectors(platform) {
  const profile = meetingAppDomCaptureProfile(platform);
  return compactStrings([
    profile?.controlSelectors,
    profile?.participantSelectors,
    profile?.textSelectors,
  ]);
}

function copyRuntimePreset(preset) {
  if (!preset) return null;
  const mutationTrackSelectors = compactStrings([
    preset.mutationTrackSelectors,
    preset.mutation_track_selectors,
    runtimePresetProfileSelectors(preset.platform),
  ]);
  return {
    ...preset,
    captureOptions: { ...(preset.captureOptions ?? {}) },
    capture_options: { ...(preset.captureOptions ?? {}) },
    mutationTrackSelectors,
    mutation_track_selectors: mutationTrackSelectors,
    mutationIgnoreSelectors: [...(preset.mutationIgnoreSelectors ?? [])],
    mutation_ignore_selectors: [...(preset.mutationIgnoreSelectors ?? [])],
  };
}

function explicitPresetInput(options = {}) {
  return firstNonEmpty(
    options.runtimePreset,
    options.runtime_preset,
    options.browserRuntimePreset,
    options.browser_runtime_preset,
    options.captureProfile,
    options.capture_profile,
    options.platform,
    options.provider,
    options.captureOptions?.platform,
    options.capture_options?.platform,
    options.captureOptions?.captureProfile,
    options.captureOptions?.capture_profile,
    options.capture_options?.captureProfile,
    options.capture_options?.capture_profile,
  );
}

function runtimePresetDisabled(options = {}) {
  return [
    options.runtimePreset,
    options.runtime_preset,
    options.browserRuntimePreset,
    options.browser_runtime_preset,
  ].some((value) => {
    if (value === false) return true;
    if (value == null) return false;
    const text = String(value).trim().toLowerCase();
    return ['none', 'off', 'false', 'disabled'].includes(text);
  });
}

function mergeSelectorOptions(presetSelectors = [], options = {}, camel, snake) {
  return compactStrings([
    presetSelectors,
    firstNonEmpty(options[camel], options[snake]),
  ]);
}

export function meetingAppBrowserRuntimePreset(platformOrInput = {}, options = {}) {
  const input = typeof platformOrInput === 'string'
    ? { platform: platformOrInput }
    : (platformOrInput ?? {});
  if (runtimePresetDisabled({ ...input, ...options })) return null;
  const inferredInput = {
    ...meetingAppBrowserInput(input),
    ...input,
  };
  const captureOptions = {
    ...(inferredInput.captureOptions ?? {}),
    ...(inferredInput.capture_options ?? {}),
    ...(options.captureOptions ?? {}),
    ...(options.capture_options ?? {}),
  };
  const explicit = explicitPresetInput({
    ...inferredInput,
    ...options,
    captureOptions,
    capture_options: captureOptions,
  });
  const profile = meetingAppDomCaptureProfile(
    explicit ? { ...inferredInput, platform: explicit } : inferredInput,
    { ...options, ...captureOptions },
  );
  return copyRuntimePreset(MEETING_APP_BROWSER_RUNTIME_PRESETS[profile?.platform]);
}

function selectorList(options = {}, camel, snake) {
  return asArray(firstNonEmpty(options[camel], options[snake])).filter(Boolean).map(String);
}

function nodeParent(node) {
  return node?.parentElement ?? node?.parentNode ?? node?.host ?? null;
}

function nodeMatchesSelector(node, selector) {
  if (!node || !selector) return false;
  if (typeof node.matches === 'function') {
    try {
      if (node.matches(selector)) return true;
    } catch {
      return false;
    }
  }
  const text = String(selector).trim();
  if (!text) return false;
  const tag = String(node.tagName ?? node.nodeName ?? '').toLowerCase();
  if (/^[a-z][a-z0-9-]*$/i.test(text)) return tag === text.toLowerCase();
  if (text.startsWith('.')) {
    return String(node.className ?? node.getAttribute?.('class') ?? '').split(/\s+/).includes(text.slice(1));
  }
  const attr = text.match(/^\[([a-zA-Z0-9_-]+)([*]?=)?(?:"([^"]*)"|'([^']*)'|([^\]\s]+))?(?:\s+i)?\]$/);
  if (!attr) return false;
  const [, name, operator, doubleQuoted, singleQuoted, bare] = attr;
  const actual = node.getAttribute?.(name) ?? node[name];
  if (operator == null) return actual != null && actual !== '';
  if (actual == null) return false;
  const expected = doubleQuoted ?? singleQuoted ?? bare ?? '';
  if (operator === '*=') return String(actual).toLowerCase().includes(String(expected).toLowerCase());
  return String(actual) === String(expected);
}

function nodeOrAncestorMatches(node, selectors = [], root = null, maxDepth = 6) {
  let current = node;
  let depth = 0;
  while (current && depth <= maxDepth) {
    if (selectors.some((selector) => nodeMatchesSelector(current, selector))) return true;
    if (current === root) break;
    current = nodeParent(current);
    depth += 1;
  }
  return false;
}

function mutationRecordNodes(record = {}) {
  return [
    record.target,
    ...asArray(record.addedNodes),
    ...asArray(record.removedNodes),
  ].filter(Boolean);
}

function mutationRecordAllowed(record = {}, filterOptions = {}) {
  const customFilter = filterOptions.mutationFilter ?? filterOptions.mutation_filter;
  if (typeof customFilter === 'function') {
    return customFilter(record, filterOptions) !== false;
  }
  const root = filterOptions.mutationRoot ?? filterOptions.mutation_root ?? null;
  const trackSelectors = selectorList(filterOptions, 'mutationTrackSelectors', 'mutation_track_selectors');
  const ignoreSelectors = selectorList(filterOptions, 'mutationIgnoreSelectors', 'mutation_ignore_selectors');
  const ignoreAttributes = selectorList(filterOptions, 'mutationIgnoreAttributes', 'mutation_ignore_attributes');
  const nodes = mutationRecordNodes(record);
  if (ignoreAttributes.length && record.type === 'attributes' && ignoreAttributes.includes(String(record.attributeName ?? ''))) {
    return false;
  }
  if (ignoreSelectors.length && nodes.some((node) => nodeOrAncestorMatches(node, ignoreSelectors, root))) {
    return false;
  }
  if (trackSelectors.length) {
    return nodes.some((node) => nodeOrAncestorMatches(node, trackSelectors, root));
  }
  return true;
}

function speakerStableFollowupEnabled(options = {}) {
  return options.speakerStableFollowup !== false
    && options.speaker_stable_followup !== false
    && options.mutationSpeakerFollowup !== false
    && options.mutation_speaker_followup !== false;
}

function trackObservationEnabled(options = {}) {
  const explicit = firstNonEmpty(
    options.observeTracks,
    options.observe_tracks,
    options.trackMutations,
    options.track_mutations,
  );
  if (explicit != null) return explicit !== false && explicit !== 'false';
  return Boolean(
    options.trackRuntimeOptions
    || options.track_runtime_options
    || options.speakerTrackOptions
    || options.speaker_track_options
    || options.participantTrackOptions
    || options.participant_track_options
  );
}

function trackSampleOptions(options = {}, sampleOptions = {}) {
  return {
    ...(options.trackSampleOptions ?? {}),
    ...(options.track_sample_options ?? {}),
    ...(sampleOptions.trackSampleOptions ?? {}),
    ...(sampleOptions.track_sample_options ?? {}),
    ...sampleOptions,
  };
}

function speakerStableFollowupMs(options = {}) {
  const speakerOptions = {
    ...(options.speakerOptions ?? {}),
    ...(options.speaker_options ?? {}),
    ...(options.sampleOptions?.speakerOptions ?? {}),
    ...(options.sample_options?.speakerOptions ?? {}),
    ...(options.sampleOptions?.speaker_options ?? {}),
    ...(options.sample_options?.speaker_options ?? {}),
  };
  const value = firstNonEmpty(
    options.speakerStableFollowupMs,
    options.speaker_stable_followup_ms,
    options.mutationStableFollowupMs,
    options.mutation_stable_followup_ms,
    speakerOptions.minStableMs,
    speakerOptions.min_stable_ms,
    300,
  );
  const numeric = Number(value);
  return Number.isFinite(numeric) ? Math.max(0, numeric) : 300;
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
  const runtimePreset = meetingAppBrowserRuntimePreset(options) ?? {};
  const mergedCaptureOptions = {
    ...(runtimePreset.captureOptions ?? {}),
    ...(runtimePreset.capture_options ?? {}),
    ...(options.captureOptions ?? {}),
    ...(options.capture_options ?? {}),
  };
  const mergedOptions = {
    ...runtimePreset,
    ...options,
    observeMutations: firstNonEmpty(options.observeMutations, options.observe_mutations, runtimePreset.observeMutations, runtimePreset.observe_mutations),
    observe_mutations: firstNonEmpty(options.observe_mutations, options.observeMutations, runtimePreset.observe_mutations, runtimePreset.observeMutations),
    mutationTrackSelectors: mergeSelectorOptions(runtimePreset.mutationTrackSelectors, options, 'mutationTrackSelectors', 'mutation_track_selectors'),
    mutation_track_selectors: mergeSelectorOptions(runtimePreset.mutation_track_selectors, options, 'mutation_track_selectors', 'mutationTrackSelectors'),
    mutationIgnoreSelectors: mergeSelectorOptions(runtimePreset.mutationIgnoreSelectors, options, 'mutationIgnoreSelectors', 'mutation_ignore_selectors'),
    mutation_ignore_selectors: mergeSelectorOptions(runtimePreset.mutation_ignore_selectors, options, 'mutation_ignore_selectors', 'mutationIgnoreSelectors'),
    captureOptions: mergedCaptureOptions,
    capture_options: mergedCaptureOptions,
  };
  const inputOptions = {
    window: mergedOptions.window,
    win: mergedOptions.win,
    document: mergedOptions.document,
    doc: mergedOptions.doc,
    location: mergedOptions.location,
    navigator: mergedOptions.navigator,
  };
  const runtime = createMeetingAppTimelineRuntime(clientOrOptions, {
    source: mergedOptions.source ?? 'browser_content_script',
    ...mergedOptions,
    captureOptions: {
      browserName: browserName(mergedOptions),
      ...(mergedOptions.captureOptions ?? {}),
      ...(mergedOptions.capture_options ?? {}),
    },
  });
  const target = eventTarget(mergedOptions);
  const stopEvents = normalizeStopEvents(mergedOptions);
  const listeners = [];
  let lifecycleInstalled = false;
  let mutationObserver = null;
  let mutationTimer = null;
  let mutationPendingCount = 0;
  let mutationIgnoredCount = 0;
  let mutationSampleCount = 0;
  let mutationLastResult = null;
  let mutationLastError = null;
  let mutationLastObservedAtMs = null;
  let mutationLastOptions = null;
  let speakerFollowupTimer = null;

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
      ...mergedOptions,
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
      ignored_count: mutationIgnoredCount,
      sample_count: mutationSampleCount,
      last_observed_at_ms: mutationLastObservedAtMs,
      last_result: mutationLastResult,
      last_error: mutationLastError,
      debounce_ms: mutationLastOptions ? debounceMs(mutationLastOptions) : null,
      speaker_followup_scheduled: speakerFollowupTimer != null,
    };
  }

  function clearMutationTimer() {
    if (mutationTimer != null) {
      clearTimeout(mutationTimer);
      mutationTimer = null;
    }
  }

  function clearSpeakerFollowupTimer() {
    if (speakerFollowupTimer != null) {
      clearTimeout(speakerFollowupTimer);
      speakerFollowupTimer = null;
    }
  }

  function signalTypesFromSampleResult(sampleResult = {}) {
    const result = sampleResult.result ?? sampleResult;
    return [
      ...(result.signals ?? []),
      ...(result.rawSignals ?? result.raw_signals ?? []),
    ].map((signal) => signal?.type).filter(Boolean);
  }

  function hasPendingSpeakerCandidate(sampleResult = {}) {
    const result = sampleResult.result ?? sampleResult;
    const speakerState = result.state?.speakerState ?? result.state?.speaker_state;
    if (speakerState?.activeSpeaker) return false;
    if (speakerState?.candidateSpeaker) return true;
    const snapshot = sampleResult.snapshot ?? {};
    const participants = [
      ...(snapshot.page?.participants ?? snapshot.page?.tiles ?? []),
      ...(snapshot.dom?.participants ?? snapshot.dom?.tiles ?? []),
    ];
    return participants.some((participant) => (
      participant?.speaking === true
      || participant?.isSpeaking === true
      || participant?.active_speaker === true
      || participant?.activeSpeaker === true
    ));
  }

  function scheduleSpeakerStableFollowup(sampleResult = {}, scheduleOptions = {}) {
    if (!speakerStableFollowupEnabled(scheduleOptions)) return mutationState();
    if (!hasPendingSpeakerCandidate(sampleResult)) return mutationState();
    if (signalTypesFromSampleResult(sampleResult).includes('speaker_started')) return mutationState();
    clearSpeakerFollowupTimer();
    speakerFollowupTimer = setTimeout(() => {
      speakerFollowupTimer = null;
      flushMutationObserver({ force: true, stabilityFollowup: true, stability_followup: true }).catch((error) => {
        mutationLastError = error?.message ?? String(error);
      });
    }, speakerStableFollowupMs(scheduleOptions));
    return mutationState();
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
      let trackResult = null;
      if (trackObservationEnabled({
        ...mergedOptions,
        ...mutationLastOptions,
        ...sampleOptions,
      })) {
        trackResult = await runtime.sampleTracks(inputProvider(), trackSampleOptions({
          ...mergedOptions,
          ...mutationLastOptions,
        }, sampleOptions));
      }
      mutationSampleCount += 1;
      mutationLastResult = trackResult ? { axis: result, tracks: trackResult } : result;
      mutationLastError = null;
      if (flushOptions.stabilityFollowup !== true && flushOptions.stability_followup !== true) {
        scheduleSpeakerStableFollowup(result, {
          ...mutationLastOptions,
          ...sampleOptions,
        });
      }
      return {
        flushed: true,
        result,
        track_result: trackResult,
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
      ...mergedOptions,
      ...mutationOptions,
    };
    const Observer = mutationObserverCtor(merged);
    const root = mutationRoot(merged);
    if (!Observer) return { installed: false, reason: 'missing_mutation_observer', state: mutationState() };
    if (!root) return { installed: false, reason: 'missing_mutation_root', state: mutationState() };
    mutationLastOptions = {
      ...merged,
      mutationRoot: root,
      mutation_root: root,
    };
    mutationObserver = new Observer((records = []) => {
      const rows = Array.isArray(records) ? records : [records];
      const allowed = rows.filter((record) => mutationRecordAllowed(record, mutationLastOptions));
      mutationIgnoredCount += Math.max(0, rows.length - allowed.length);
      if (allowed.length === 0) return;
      mutationPendingCount += Math.max(1, allowed.length);
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
    clearSpeakerFollowupTimer();
    if (mutationObserver?.disconnect) mutationObserver.disconnect();
    mutationObserver = null;
    mutationPendingCount = 0;
    mutationIgnoredCount = 0;
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
    if ([
      'meeting_timeline.sample_tracks',
      'meeting_timeline_sample_tracks',
      'sample_tracks',
      'track_sample',
    ].includes(type)) {
      return { handled: true, action: 'sampleTracks', result: await sampleTracks(messageOptions) };
    }
    if ([
      'meeting_timeline.start_tracks',
      'meeting_timeline_start_tracks',
      'start_tracks',
    ].includes(type)) {
      return { handled: true, action: 'startTracks', result: startTracks(messageOptions) };
    }
    if ([
      'meeting_timeline.stop_tracks',
      'meeting_timeline_stop_tracks',
      'stop_tracks',
    ].includes(type)) {
      return { handled: true, action: 'stopTracks', result: runtime.stopTracks() };
    }
    if ([
      'meeting_timeline.observe_tracks',
      'meeting_timeline_observe_tracks',
      'observe_tracks',
    ].includes(type)) {
      return { handled: true, action: 'observeMeetingAppTracks', result: await runtime.observeMeetingAppTracks(payload, messageOptions) };
    }
    if ([
      'meeting_timeline.preview_tracks',
      'meeting_timeline_preview_tracks',
      'preview_tracks',
    ].includes(type)) {
      return { handled: true, action: 'previewMeetingAppTracks', result: runtime.previewMeetingAppTracks(payload, messageOptions) };
    }
    return { handled: false, reason: 'unsupported_message_type', type };
  }

  function start(startOptions = {}) {
    if (startOptions.lifecycle !== false && startOptions.installLifecycleHandlers !== false && startOptions.install_lifecycle_handlers !== false) {
      installLifecycleHandlers(startOptions);
    }
    if (mutationEnabled({
      ...mergedOptions,
      ...startOptions,
    })) {
      installMutationObserver(startOptions);
    }
    const state = runtime.start(() => inputProvider(), startOptions);
    if (trackObservationEnabled({ ...mergedOptions, ...startOptions })
      && (startOptions.startTracksInterval === true || startOptions.start_tracks_interval === true)) {
      runtime.startTracks(() => inputProvider(), startOptions);
    }
    return state;
  }

  function stop() {
    if (mergedOptions.keepMutationObserverOnStop !== true && mergedOptions.keep_mutation_observer_on_stop !== true) {
      removeMutationObserver();
    }
    runtime.stopTracks();
    return runtime.stop();
  }

  function dispose() {
    const stopped = runtime.stop();
    runtime.stopTracks();
    removeMutationObserver();
    removeLifecycleHandlers();
    return stopped;
  }

  function sampleTracks(sampleOptions = {}) {
    return runtime.sampleTracks(inputProvider(), sampleOptions);
  }

  function tickTracks(sampleOptions = {}) {
    return runtime.tickTracks(inputProvider(), sampleOptions);
  }

  function startTracks(startOptions = {}) {
    return runtime.startTracks(() => inputProvider(), startOptions);
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
    sampleTracks,
    tickTracks,
    startTracks,
    stopTracks() {
      return runtime.stopTracks();
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
