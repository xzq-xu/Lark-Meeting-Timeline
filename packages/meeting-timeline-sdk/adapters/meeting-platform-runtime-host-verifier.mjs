import {
  MeetingTimelineSdkError,
  compactObject,
  normalizeAbsoluteMs,
} from '../index.mjs';
import {
  createMeetingAppBrowserRuntime,
} from './meeting-app-browser-runtime.mjs';
import {
  MEETING_APP_FIXTURE_PLATFORMS,
  buildMeetingAppFixtureSnapshot,
} from './meeting-app-fixtures.mjs';
import {
  buildMeetingPlatformRuntimeHostConfig,
  createMeetingPlatformRuntimeHost,
} from './meeting-platform-runtime-host.mjs';
import {
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_RUNTIME_HOST_VERIFICATION_SCHEMA = 'meeting_platform_runtime_host_verification';
export const MEETING_PLATFORM_RUNTIME_HOST_VERIFICATION_MATRIX_SCHEMA = 'meeting_platform_runtime_host_verification_matrix';
export const MEETING_PLATFORM_RUNTIME_HOST_VERIFICATION_SCHEMA_VERSION = 1;

const DEFAULT_START_MS = 1_783_356_000_000;
const DEFAULT_END_OFFSET_MS = 1_000;

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
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_APP_FIXTURE_PLATFORMS))
    .map((platform) => normalizeMeetingPlatform(platform))
    .filter((platform) => MEETING_APP_FIXTURE_PLATFORMS.includes(platform)));
}

function startMs(options = {}) {
  return normalizeAbsoluteMs(firstNonEmpty(
    options.startMs,
    options.start_ms,
    options.observedAtMs,
    options.observed_at_ms,
    DEFAULT_START_MS,
  ), 'meeting_platform_runtime_host_verification_start_ms');
}

function endMs(start, options = {}) {
  const explicit = firstNonEmpty(options.endMs, options.end_ms);
  if (explicit != null) return normalizeAbsoluteMs(explicit, 'meeting_platform_runtime_host_verification_end_ms');
  const offset = Number(firstNonEmpty(options.endOffsetMs, options.end_offset_ms, DEFAULT_END_OFFSET_MS));
  return start + (Number.isFinite(offset) ? Math.max(1, offset) : DEFAULT_END_OFFSET_MS);
}

function simpleText(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function dataKey(name) {
  return String(name).replace(/-([a-z])/g, (_, char) => char.toUpperCase());
}

class FixtureNode {
  constructor(kind, attrs = {}, text = '') {
    this.nodeType = 1;
    this.tagName = kind === 'control' ? 'BUTTON' : 'DIV';
    this.nodeName = this.tagName;
    this.attributes = Object.fromEntries(Object.entries(attrs).filter(([, value]) => value != null && value !== ''));
    this.dataset = {};
    for (const [name, value] of Object.entries(this.attributes)) {
      if (name.startsWith('data-')) this.dataset[dataKey(name.slice(5))] = value;
    }
    this.className = this.attributes.class ?? '';
    this.innerText = simpleText(text || this.attributes['aria-label'] || this.attributes.title || this.attributes['data-tooltip'] || this.attributes['data-display-name']);
    this.textContent = this.innerText;
  }

  getAttribute(name) {
    return this.attributes[name] ?? null;
  }
}

function attrMatches(node, chunk) {
  const match = String(chunk).match(/^\[([^\]*~=\s]+)(\*=|=)?(?:"([^"]*)"|'([^']*)'|([^\]\s]+))?(?:\s+i)?\]$/i);
  if (!match) return true;
  const [, name, operator, doubleQuoted, singleQuoted, bare] = match;
  const actual = node.getAttribute?.(name);
  if (operator == null) return actual != null && actual !== '';
  if (actual == null) return false;
  const expected = doubleQuoted ?? singleQuoted ?? bare ?? '';
  const insensitive = /\s+i\]$/i.test(chunk);
  const left = insensitive ? String(actual).toLowerCase() : String(actual);
  const right = insensitive ? String(expected).toLowerCase() : String(expected);
  return operator === '*=' ? left.includes(right) : left === right;
}

function selectorMatches(node, selector) {
  const text = String(selector ?? '').trim();
  if (!text) return false;
  if (text === '*') return true;
  const classMatch = text.match(/^\.([a-zA-Z0-9_-]+)$/);
  if (classMatch) return String(node.className ?? '').split(/\s+/).includes(classMatch[1]);
  const tag = text.match(/^[a-z][a-z0-9-]*/i)?.[0];
  if (tag && !text.startsWith('[') && !text.startsWith('.') && String(node.tagName).toLowerCase() !== tag.toLowerCase()) {
    return false;
  }
  const chunks = text.match(/\[[^\]]+\]/g) ?? [];
  if (chunks.length > 0) return chunks.every((chunk) => attrMatches(node, chunk));
  return tag ? String(node.tagName).toLowerCase() === tag.toLowerCase() : false;
}

class FixtureDocument {
  constructor(snapshot = {}) {
    this.nodeType = 9;
    this.title = snapshot.title ?? '';
    this.hidden = false;
    this.visibilityState = 'visible';
    this.location = {
      href: snapshot.url ?? snapshot.meeting_url ?? '',
      toString() {
        return this.href;
      },
    };
    this.nodes = [
      ...(snapshot.page?.buttons ?? snapshot.page?.controls ?? []).map((control) => new FixtureNode('control', {
        role: control.role ?? 'button',
        'aria-label': control.ariaLabel ?? control.aria_label ?? control.label,
        title: control.title ?? control.label,
        'data-tooltip': control.label,
      }, control.label ?? control.text)),
      ...(snapshot.page?.participants ?? snapshot.page?.tiles ?? []).map((participant) => new FixtureNode('participant', {
        'data-participant-id': participant.id,
        'data-display-name': participant.name ?? participant.display_name,
        'data-speaking': Number(participant.audioLevel ?? participant.audio_level ?? 0) >= 0.2 || participant.speaking === true || participant.isSpeaking === true ? 'true' : undefined,
        'aria-label': participant.ariaLabel ?? participant.aria_label ?? participant.label ?? participant.name,
      }, participant.label ?? participant.name ?? participant.display_name)),
      ...(snapshot.page?.texts ?? snapshot.dom?.texts ?? []).map((item) => new FixtureNode('text', {
        role: item.role ?? 'status',
        'aria-live': item.ariaLive ?? item.aria_live ?? 'polite',
        'aria-label': item.label,
      }, item.text ?? item.label)),
    ];
  }

  get body() {
    return this;
  }

  get documentElement() {
    return this;
  }

  querySelectorAll(selector) {
    return this.nodes.filter((node) => selectorMatches(node, selector));
  }
}

export function createMeetingPlatformRuntimeHostVerificationClient(options = {}) {
  const calls = [];
  const client = {
    calls,
    async startMeeting(input) {
      calls.push({ action: 'startMeeting', input });
      return { ok: true, action: 'startMeeting', input };
    },
    async endMeeting(input) {
      calls.push({ action: 'endMeeting', input });
      return { ok: true, action: 'endMeeting', input };
    },
    async insertMark(input) {
      calls.push({ action: 'insertMark', input });
      return { ok: true, action: 'insertMark', input };
    },
    async insertMarks(input) {
      calls.push({ action: 'insertMarks', input });
      return { ok: true, action: 'insertMarks', input };
    },
    getState() {
      return {
        call_count: calls.length,
        actions: calls.map((call) => call.action),
      };
    },
  };
  if (options.exposeCalls === false || options.expose_calls === false) return client;
  Object.defineProperty(client, '__meetingTimelineCalls', {
    value: calls,
    enumerable: false,
  });
  return client;
}

export function createMeetingPlatformRuntimeHostFixtureEnvironment(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  let currentMs = startMs(options);
  let currentSnapshot = buildMeetingAppFixtureSnapshot(key, {
    ...options,
    state: firstNonEmpty(options.state, 'active'),
    observedAtMs: currentMs,
  });
  let currentDocument = new FixtureDocument(currentSnapshot);
  const win = {
    navigator: {
      userAgent: firstNonEmpty(options.userAgent, options.user_agent, 'MeetingTimelineSdkFixtureBrowser'),
    },
    get document() {
      return currentDocument;
    },
    get location() {
      return currentDocument.location;
    },
  };

  function setState(state = 'active', nextOptions = {}) {
    currentMs = normalizeAbsoluteMs(firstNonEmpty(
      nextOptions.observedAtMs,
      nextOptions.observed_at_ms,
      nextOptions.nowMs,
      nextOptions.now_ms,
      currentMs,
    ), 'meeting_platform_runtime_host_fixture_time');
    currentSnapshot = buildMeetingAppFixtureSnapshot(key, {
      ...options,
      ...nextOptions,
      state,
      observedAtMs: currentMs,
    });
    currentDocument = new FixtureDocument(currentSnapshot);
    return getState();
  }

  function getState() {
    return {
      platform: key,
      now_ms: currentMs,
      fixture_state: currentSnapshot.fixture_state,
      snapshot: currentSnapshot,
      document: currentDocument,
      window: win,
      location: currentDocument.location,
    };
  }

  return {
    platform: key,
    window: win,
    setState,
    setActive(nextOptions = {}) {
      return setState('active', nextOptions);
    },
    setEnded(nextOptions = {}) {
      return setState('prejoin', nextOptions);
    },
    now() {
      return currentMs;
    },
    get document() {
      return currentDocument;
    },
    get location() {
      return currentDocument.location;
    },
    get snapshot() {
      return currentSnapshot;
    },
    getState,
  };
}

function signalTypes(result = {}) {
  return [
    ...(result.result?.result?.signals ?? []),
    ...(result.result?.result?.rawSignals ?? []),
  ].map((signal) => signal?.type).filter(Boolean);
}

function resultActions(result = {}) {
  return (result.result?.result?.results ?? []).map((row) => row.action).filter(Boolean);
}

function captureSummary(result = {}) {
  const capture = result.result?.snapshot?.capture ?? {};
  return {
    profile: capture.profile,
    control_count: capture.control_count ?? 0,
    participant_count: capture.participant_count ?? 0,
    text_count: capture.text_count ?? 0,
  };
}

function callTimestamp(call = {}) {
  return firstNonEmpty(call.input?.start_time_ms, call.input?.captured_at_ms, call.input?.end_time_ms);
}

function buildCoverage({ platform, calls, startResult, endResult, startAtMs, endAtMs, config }) {
  const startSignals = signalTypes(startResult);
  const endSignals = signalTypes(endResult);
  const actions = calls.map((call) => call.action);
  const startCall = calls.find((call) => call.action === 'startMeeting');
  const endCall = calls.find((call) => call.action === 'endMeeting');
  const speakerCall = calls.find((call) => call.action === 'insertMark' && call.input?.kind === 'speaker_started');
  const startCapture = captureSummary(startResult);
  const endCapture = captureSummary(endResult);
  return {
    host_ready: config.readiness?.host_ready === true,
    runtime_capture_profile: startCapture.profile === platform && endCapture.profile === platform,
    meeting_start_signal: startSignals.includes('meeting_started'),
    speaker_start_signal: startSignals.includes('speaker_started'),
    meeting_end_signal: endSignals.includes('meeting_ended'),
    start_meeting_written: actions.includes('startMeeting'),
    speaker_mark_written: Boolean(speakerCall),
    end_meeting_written: actions.includes('endMeeting'),
    start_timestamp_aligned: callTimestamp(startCall) === startAtMs,
    speaker_timestamp_aligned: callTimestamp(speakerCall) === startAtMs,
    end_timestamp_aligned: callTimestamp(endCall) === endAtMs,
    participant_capture: startCapture.participant_count > 0,
    control_capture: startCapture.control_count > 0 && endCapture.control_count > 0,
  };
}

function missingCoverage(coverage = {}) {
  return Object.entries(coverage)
    .filter(([, ok]) => ok !== true)
    .map(([key]) => key);
}

export async function runMeetingPlatformRuntimeHostVerification(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  if (!MEETING_APP_FIXTURE_PLATFORMS.includes(key)) {
    throw new MeetingTimelineSdkError(`Unsupported runtime host verification platform: ${String(platform || '(empty)')}`, {
      platform,
      supported_platforms: MEETING_APP_FIXTURE_PLATFORMS,
    });
  }
  const startAtMs = startMs(options);
  const endAtMs = endMs(startAtMs, options);
  const explicitEnv = firstNonEmpty(options.environment, options.fixtureEnvironment, options.fixture_environment);
  const env = explicitEnv?.setActive && explicitEnv?.setEnded
    ? explicitEnv
    : createMeetingPlatformRuntimeHostFixtureEnvironment(key, {
    ...options,
    observedAtMs: startAtMs,
  });
  const client = options.client ?? createMeetingPlatformRuntimeHostVerificationClient(options.clientOptions ?? options.client_options);
  const calls = client.__meetingTimelineCalls ?? client.calls ?? [];
  const runtime = options.runtime ?? createMeetingAppBrowserRuntime(client, {
    platform: key,
    window: env.window,
    now: () => env.now(),
    applyOptions: {
      speakerAsAnnotation: true,
      ...(options.applyOptions ?? {}),
      ...(options.apply_options ?? {}),
    },
    speakerOptions: {
      minStableMs: 0,
      ...(options.speakerOptions ?? {}),
      ...(options.speaker_options ?? {}),
    },
    ...(options.runtimeOptions ?? {}),
    ...(options.runtime_options ?? {}),
  });
  const config = options.config?.schema === 'meeting_platform_runtime_host_config'
    ? options.config
    : buildMeetingPlatformRuntimeHostConfig(key, options.configOptions ?? options.config_options ?? options);
  const host = options.host ?? createMeetingPlatformRuntimeHost(runtime, config, {
    now: () => env.now(),
    ...(options.hostOptions ?? {}),
    ...(options.host_options ?? {}),
  });

  env.setActive({ observedAtMs: startAtMs });
  const startResult = await host.sample('dom_mutation', undefined, {
    force: true,
    ...(options.startSampleOptions ?? {}),
    ...(options.start_sample_options ?? {}),
  });
  env.setEnded({ observedAtMs: endAtMs });
  const endResult = await host.sample('meeting_candidate_missing', undefined, {
    force: true,
    ...(options.endSampleOptions ?? {}),
    ...(options.end_sample_options ?? {}),
  });
  host.stop?.();

  const coverage = buildCoverage({
    platform: key,
    calls,
    startResult,
    endResult,
    startAtMs,
    endAtMs,
    config,
  });
  const missing = missingCoverage(coverage);
  return compactObject({
    type: 'meeting_platform_runtime_host_verification',
    schema: MEETING_PLATFORM_RUNTIME_HOST_VERIFICATION_SCHEMA,
    version: MEETING_PLATFORM_RUNTIME_HOST_VERIFICATION_SCHEMA_VERSION,
    platform: key,
    accepted: missing.length === 0,
    start_at_ms: startAtMs,
    end_at_ms: endAtMs,
    call_count: calls.length,
    actions: calls.map((call) => call.action),
    signal_types: unique([...signalTypes(startResult), ...signalTypes(endResult)]),
    result_actions: unique([...resultActions(startResult), ...resultActions(endResult)]),
    coverage,
    missing,
    capture: {
      start: captureSummary(startResult),
      end: captureSummary(endResult),
    },
    calls,
    host_state: host.getState?.(),
    next_actions: missing.length === 0
      ? ['replace_fixture_environment_with_live_meeting_app_snapshot_capture']
      : missing,
  });
}

export async function runMeetingPlatformRuntimeHostVerificationMatrix(options = {}) {
  const platforms = selectedPlatforms(options);
  const reports = [];
  for (const platform of platforms) {
    reports.push(await runMeetingPlatformRuntimeHostVerification(platform, {
      ...options,
      platform,
      platforms: undefined,
      platform_keys: undefined,
    }));
  }
  return {
    type: 'meeting_platform_runtime_host_verification_matrix',
    schema: MEETING_PLATFORM_RUNTIME_HOST_VERIFICATION_MATRIX_SCHEMA,
    version: MEETING_PLATFORM_RUNTIME_HOST_VERIFICATION_SCHEMA_VERSION,
    platform_count: reports.length,
    accepted_count: reports.filter((report) => report.accepted).length,
    platforms,
    rows: reports.map((report) => ({
      platform: report.platform,
      accepted: report.accepted,
      call_count: report.call_count,
      actions: report.actions,
      signal_types: report.signal_types,
      missing: report.missing,
      start_capture_profile: report.capture?.start?.profile,
      end_capture_profile: report.capture?.end?.profile,
      participant_count: report.capture?.start?.participant_count ?? 0,
    })),
    reports,
    next_actions: unique(reports.flatMap((report) => report.next_actions ?? [])),
  };
}

export function assertMeetingPlatformRuntimeHostVerification(report = {}) {
  if (report.accepted !== true) {
    throw new MeetingTimelineSdkError('meeting platform runtime host verification failed', {
      platform: report.platform,
      missing: report.missing,
    });
  }
  return report;
}

export function assertMeetingPlatformRuntimeHostVerificationMatrix(matrix = {}) {
  const failed = (matrix.reports ?? []).filter((report) => report.accepted !== true);
  if (failed.length > 0) {
    throw new MeetingTimelineSdkError('meeting platform runtime host verification matrix failed', {
      failed: failed.map((report) => ({ platform: report.platform, missing: report.missing })),
    });
  }
  return matrix;
}
