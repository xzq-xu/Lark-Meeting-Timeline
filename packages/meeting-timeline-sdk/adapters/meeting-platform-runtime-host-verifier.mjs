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
  normalizeMeetingAppSnapshot,
} from './meeting-apps.mjs';
import {
  createMeetingSourceAggregator,
} from './meeting-source.mjs';
import {
  buildMeetingPlatformRuntimeHostConfig,
  createMeetingPlatformRuntimeHost,
} from './meeting-platform-runtime-host.mjs';
import {
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_RUNTIME_HOST_VERIFICATION_SCHEMA = 'meeting_platform_runtime_host_verification';
export const MEETING_PLATFORM_RUNTIME_HOST_VERIFICATION_MATRIX_SCHEMA = 'meeting_platform_runtime_host_verification_matrix';
export const MEETING_PLATFORM_RUNTIME_HOST_REPLAY_SCHEMA = 'meeting_platform_runtime_host_replay';
export const MEETING_PLATFORM_RUNTIME_HOST_REPLAY_MATRIX_SCHEMA = 'meeting_platform_runtime_host_replay_matrix';
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

function observedResult(result = {}) {
  let node = result;
  for (let depth = 0; depth < 4; depth += 1) {
    if (!node || typeof node !== 'object') return {};
    if (Array.isArray(node.signals) || Array.isArray(node.rawSignals) || Array.isArray(node.results)) return node;
    node = node.result;
  }
  return {};
}

function signalTypes(result = {}) {
  const observed = observedResult(result);
  return [
    ...(observed.signals ?? []),
    ...(observed.rawSignals ?? []),
  ].map((signal) => signal?.type).filter(Boolean);
}

function resultActions(result = {}) {
  return (observedResult(result).results ?? []).map((row) => row.action).filter(Boolean);
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

function normalizeReplayArgs(platformOrInput, inputOrOptions = {}, maybeOptions = {}) {
  if (platformOrInput && typeof platformOrInput === 'object' && !Array.isArray(platformOrInput)) {
    const input = platformOrInput;
    return {
      input,
      platform: firstNonEmpty(input.platform, input.provider, input.adapter, input.rollout_plan?.platform),
      options: inputOrOptions ?? {},
    };
  }
  return {
    input: inputOrOptions ?? {},
    platform: platformOrInput,
    options: maybeOptions ?? {},
  };
}

function tryNormalizePlatform(value) {
  if (!value || typeof value === 'object') return undefined;
  try {
    return normalizeMeetingPlatform(value);
  } catch {
    return undefined;
  }
}

function replayInputPlatform(value = {}) {
  return tryNormalizePlatform(firstNonEmpty(
    value.platform,
    value.provider,
    value.adapter,
    value.capture?.profile,
    value.snapshot?.platform,
    value.snapshot?.provider,
    value.snapshot?.capture?.profile,
    value.rollout_plan?.platform,
  ));
}

function matchesReplayPlatform(value, platform) {
  const key = replayInputPlatform(value);
  return !key || key === platform;
}

function recordPhase(record = {}, snapshot = {}) {
  const raw = firstNonEmpty(record.phase, record.state, record.fixture_state, snapshot.phase, snapshot.state, snapshot.fixture_state);
  const text = raw == null ? '' : String(raw).trim().toLowerCase().replace(/[-\s]+/g, '_');
  if (['active', 'joined', 'in_meeting', 'start', 'started'].includes(text)) return 'active';
  if (['ended', 'left', 'prejoin', 'pre_join', 'inactive', 'end', 'finished'].includes(text)) return 'ended';
  return text || undefined;
}

function recordTimeMs(record = {}, snapshot = {}, fallbackMs = DEFAULT_START_MS) {
  return normalizeAbsoluteMs(firstNonEmpty(
    record.capturedAtMs,
    record.captured_at_ms,
    record.observedAtMs,
    record.observed_at_ms,
    record.timestamp_ms,
    record.timestampMs,
    record.createdAtMs,
    record.created_at_ms,
    snapshot.observedAtMs,
    snapshot.observed_at_ms,
    snapshot.capturedAtMs,
    snapshot.captured_at_ms,
    fallbackMs,
  ), 'meeting_platform_runtime_host_replay_time');
}

function snapshotFromReplayRecord(record = {}) {
  if (!record || typeof record !== 'object') return null;
  const explicit = firstNonEmpty(
    record.snapshot,
    record.meeting_app_snapshot,
    record.meetingAppSnapshot,
    record.dom_snapshot,
    record.domSnapshot,
    record.payload?.snapshot,
    record.payload?.meeting_app_snapshot,
    record.payload?.meetingAppSnapshot,
    record.raw?.snapshot,
    record.raw?.meeting_app_snapshot,
    record.input?.snapshot,
  );
  if (explicit && typeof explicit === 'object') return explicit;
  if (
    record.page
    || record.dom
    || record.capture
    || record.participants
    || record.tiles
    || record.activeSpeaker
    || record.active_speaker
    || record.inMeeting != null
    || record.in_meeting != null
    || record.url
    || record.meeting_url
  ) {
    return record;
  }
  return null;
}

function pushReplayCollection(target, platform, collection, seen = new Set()) {
  if (!collection) return target;
  if (Array.isArray(collection)) {
    for (const item of collection) pushReplayCollection(target, platform, item, seen);
    return target;
  }
  if (typeof collection !== 'object') return target;
  if (collection.records && Array.isArray(collection.records)) {
    pushReplayCollection(target, platform, collection.records, seen);
    return target;
  }
  const keyedCollections = [
    collection[platform],
    collection[platform.replace(/_/g, '-')],
    platform === 'microsoft_teams' ? collection.teams : undefined,
    platform === 'google_meet' ? collection.googleMeet : undefined,
  ];
  if (!snapshotFromReplayRecord(collection) && keyedCollections.some(Boolean)) {
    for (const item of keyedCollections) pushReplayCollection(target, platform, item, seen);
    return target;
  }
  if (snapshotFromReplayRecord(collection)) {
    if (matchesReplayPlatform(collection, platform) && !seen.has(collection)) {
      seen.add(collection);
      target.push(collection);
    }
    return target;
  }
  for (const key of [
    'items',
    'snapshots',
    'records',
    'snapshot_records',
    'meetingAppRecordSet',
    'meeting_app_record_set',
    'recordSet',
    'record_set',
    'meetingAppRecords',
    'meeting_app_records',
    'meetingAppSnapshotRecords',
    'meeting_app_snapshot_records',
    'meetingAppSnapshots',
    'meeting_app_snapshots',
    'domSnapshots',
    'dom_snapshots',
    'evidencePackages',
    'evidence_packages',
    'packages',
  ]) {
    pushReplayCollection(target, platform, collection[key], seen);
  }
  return target;
}

function replayRecordsForPlatform(platform, input = {}, options = {}) {
  const rows = [];
  const seen = new Set();
  for (const source of [
    input,
    options.input,
    options.inputs,
    options.evidenceByPlatform,
    options.evidence_by_platform,
    options.inputByPlatform,
    options.input_by_platform,
  ]) {
    pushReplayCollection(rows, platform, source, seen);
  }
  return rows;
}

function replayRowsForPlatform(platform, input = {}, options = {}) {
  return replayRecordsForPlatform(platform, input, options)
    .map((record, index) => {
      const sourceSnapshot = snapshotFromReplayRecord(record);
      if (!sourceSnapshot) return null;
      const phase = recordPhase(record, sourceSnapshot);
      const observedAtMs = recordTimeMs(record, sourceSnapshot, startMs(options) + index * 1_000);
      const snapshot = {
        ...sourceSnapshot,
        platform,
        provider: platform,
        observedAtMs,
        observed_at_ms: observedAtMs,
        inMeeting: sourceSnapshot.inMeeting ?? sourceSnapshot.in_meeting ?? (phase === 'active' ? true : phase === 'ended' ? false : undefined),
        in_meeting: sourceSnapshot.in_meeting ?? sourceSnapshot.inMeeting ?? (phase === 'active' ? true : phase === 'ended' ? false : undefined),
      };
      const normalized = normalizeMeetingAppSnapshot(snapshot, {
        platform,
        observedAtMs,
      });
      return normalized ? compactObject({
        index,
        id: firstNonEmpty(record.id, record.record_id, `${platform}-replay-${index}`),
        phase,
        captured_at_ms: observedAtMs,
        source: firstNonEmpty(record.source, sourceSnapshot.source),
        snapshot,
        normalized,
      }) : null;
    })
    .filter(Boolean)
    .sort((left, right) => left.captured_at_ms - right.captured_at_ms || left.index - right.index);
}

function selectReplayActiveRow(rows = []) {
  return rows.find((row) => row.normalized?.inMeeting === true)
    ?? rows.find((row) => row.phase === 'active')
    ?? null;
}

function selectReplayEndRow(rows = [], activeRow = null) {
  const afterActive = activeRow ? rows.filter((row) => row.captured_at_ms >= activeRow.captured_at_ms && row.index !== activeRow.index) : rows;
  return afterActive.find((row) => row.normalized?.inMeeting === false)
    ?? afterActive.find((row) => row.phase === 'ended')
    ?? null;
}

function replayRowSummary(row = {}) {
  const speaker = row.normalized?.activeSpeaker ?? {};
  return compactObject({
    index: row.index,
    id: row.id,
    phase: row.phase,
    captured_at_ms: row.captured_at_ms,
    platform: row.normalized?.platform,
    in_meeting: row.normalized?.inMeeting,
    meeting_id: row.normalized?.meeting_id,
    meeting_url: row.normalized?.meeting_url ?? row.normalized?.url,
    title: row.normalized?.title,
    participant_count: asArray(row.normalized?.participants).length,
    active_speaker_id: speaker.id,
    active_speaker_name: speaker.name ?? speaker.display_name,
    has_active_speaker: Boolean(speaker.id || speaker.name || speaker.display_name),
  });
}

function createReplayRuntime(client, options = {}) {
  const sources = createMeetingSourceAggregator(client, {
    source: 'meeting_platform_runtime_host_replay',
    ...options,
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
  });
  return {
    sources,
    sample(input = {}, sampleOptions = {}) {
      return sources.observeMeetingApp(input, {
        source: 'meeting_platform_runtime_host_replay',
        ...sampleOptions,
      });
    },
    getState() {
      return {
        replay_runtime: true,
        sources: sources.getState?.(),
      };
    },
    stop() {
      return { stopped: true };
    },
  };
}

function buildReplayCoverage({ calls, startResult, endResult, activeRow, endRow, config }) {
  const startSignals = signalTypes(startResult);
  const endSignals = signalTypes(endResult);
  const actions = calls.map((call) => call.action);
  const startCall = calls.find((call) => call.action === 'startMeeting');
  const endCall = calls.find((call) => call.action === 'endMeeting');
  const speakerCall = calls.find((call) => call.action === 'insertMark' && call.input?.kind === 'speaker_started');
  return {
    host_ready: config.readiness?.host_ready === true,
    replay_input_available: Boolean(activeRow || endRow),
    active_snapshot_available: Boolean(activeRow),
    end_snapshot_available: Boolean(endRow),
    active_snapshot_in_meeting: activeRow?.normalized?.inMeeting === true,
    end_snapshot_out_of_meeting: endRow?.normalized?.inMeeting === false,
    active_speaker_available: Boolean(activeRow?.normalized?.activeSpeaker?.id || activeRow?.normalized?.activeSpeaker?.name || activeRow?.normalized?.activeSpeaker?.display_name),
    meeting_start_signal: startSignals.includes('meeting_started'),
    speaker_start_signal: startSignals.includes('speaker_started'),
    meeting_end_signal: endSignals.includes('meeting_ended'),
    start_meeting_written: actions.includes('startMeeting'),
    speaker_mark_written: Boolean(speakerCall),
    end_meeting_written: actions.includes('endMeeting'),
    start_timestamp_aligned: callTimestamp(startCall) === activeRow?.captured_at_ms,
    speaker_timestamp_aligned: callTimestamp(speakerCall) === activeRow?.captured_at_ms,
    end_timestamp_aligned: callTimestamp(endCall) === endRow?.captured_at_ms,
  };
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

export async function runMeetingPlatformRuntimeHostReplay(platformOrInput, inputOrOptions = {}, maybeOptions = {}) {
  const { input, platform, options } = normalizeReplayArgs(platformOrInput, inputOrOptions, maybeOptions);
  const key = normalizeMeetingPlatform(platform);
  if (!MEETING_APP_FIXTURE_PLATFORMS.includes(key)) {
    throw new MeetingTimelineSdkError(`Unsupported runtime host replay platform: ${String(platform || '(empty)')}`, {
      platform,
      supported_platforms: MEETING_APP_FIXTURE_PLATFORMS,
    });
  }
  const rows = replayRowsForPlatform(key, input, options);
  const activeRow = selectReplayActiveRow(rows);
  const endRow = selectReplayEndRow(rows, activeRow);
  const client = options.client ?? createMeetingPlatformRuntimeHostVerificationClient(options.clientOptions ?? options.client_options);
  const calls = client.__meetingTimelineCalls ?? client.calls ?? [];
  const config = options.config?.schema === 'meeting_platform_runtime_host_config'
    ? options.config
    : buildMeetingPlatformRuntimeHostConfig(key, options.configOptions ?? options.config_options ?? options);

  let currentMs = activeRow?.captured_at_ms ?? rows[0]?.captured_at_ms ?? startMs(options);
  let startResult = null;
  let endResult = null;
  let host = null;

  if (activeRow && endRow) {
    const runtime = options.runtime ?? createReplayRuntime(client, {
      ...options,
      platform: key,
      provider: key,
    });
    host = options.host ?? createMeetingPlatformRuntimeHost(runtime, config, {
      now: () => currentMs,
      runtimeInputMode: 'explicit',
      runtime_input_mode: 'explicit',
      ...(options.hostOptions ?? {}),
      ...(options.host_options ?? {}),
    });
    currentMs = activeRow.captured_at_ms;
    startResult = await host.sample('dom_mutation', activeRow.normalized, {
      force: true,
      ...(options.startSampleOptions ?? {}),
      ...(options.start_sample_options ?? {}),
    });
    currentMs = endRow.captured_at_ms;
    endResult = await host.sample('meeting_candidate_missing', endRow.normalized, {
      force: true,
      ...(options.endSampleOptions ?? {}),
      ...(options.end_sample_options ?? {}),
    });
    host.stop?.();
  }

  const coverage = buildReplayCoverage({
    calls,
    startResult,
    endResult,
    activeRow,
    endRow,
    config,
  });
  const missing = missingCoverage(coverage);
  return compactObject({
    type: 'meeting_platform_runtime_host_replay',
    schema: MEETING_PLATFORM_RUNTIME_HOST_REPLAY_SCHEMA,
    version: MEETING_PLATFORM_RUNTIME_HOST_VERIFICATION_SCHEMA_VERSION,
    platform: key,
    accepted: missing.length === 0,
    input_record_count: replayRecordsForPlatform(key, input, options).length,
    replay_row_count: rows.length,
    active_at_ms: activeRow?.captured_at_ms,
    end_at_ms: endRow?.captured_at_ms,
    active_row: activeRow ? replayRowSummary(activeRow) : undefined,
    end_row: endRow ? replayRowSummary(endRow) : undefined,
    rows: rows.map((row) => replayRowSummary(row)),
    call_count: calls.length,
    actions: calls.map((call) => call.action),
    signal_types: unique([...signalTypes(startResult), ...signalTypes(endResult)]),
    result_actions: unique([...resultActions(startResult), ...resultActions(endResult)]),
    coverage,
    missing,
    calls,
    host_state: host?.getState?.(),
    next_actions: missing.length === 0
      ? ['replace_runtime_host_replay_input_with_fresh_live_capture_for_each_target_meeting_platform']
      : missing,
  });
}

function replayInputForPlatform(platform, options = {}) {
  const keyed = firstNonEmpty(
    options.evidenceByPlatform?.[platform],
    options.evidence_by_platform?.[platform],
    options.inputByPlatform?.[platform],
    options.input_by_platform?.[platform],
    options.inputs?.[platform],
    options.evidenceByPlatform?.[platform.replace(/_/g, '-')],
    options.evidence_by_platform?.[platform.replace(/_/g, '-')],
    options.inputByPlatform?.[platform.replace(/_/g, '-')],
    options.input_by_platform?.[platform.replace(/_/g, '-')],
    options.inputs?.[platform.replace(/_/g, '-')],
  );
  return keyed ?? firstNonEmpty(options.input, options.inputs, options.evidence, options.package, options.packages, {});
}

export async function runMeetingPlatformRuntimeHostReplayMatrix(options = {}) {
  const platforms = selectedPlatforms(options);
  const reports = [];
  for (const platform of platforms) {
    reports.push(await runMeetingPlatformRuntimeHostReplay(platform, replayInputForPlatform(platform, options), {
      ...options,
      platform,
      platforms: undefined,
      platform_keys: undefined,
    }));
  }
  return {
    type: 'meeting_platform_runtime_host_replay_matrix',
    schema: MEETING_PLATFORM_RUNTIME_HOST_REPLAY_MATRIX_SCHEMA,
    version: MEETING_PLATFORM_RUNTIME_HOST_VERIFICATION_SCHEMA_VERSION,
    platform_count: reports.length,
    accepted_count: reports.filter((report) => report.accepted).length,
    platforms,
    rows: reports.map((report) => ({
      platform: report.platform,
      accepted: report.accepted,
      input_record_count: report.input_record_count,
      replay_row_count: report.replay_row_count,
      active_at_ms: report.active_at_ms,
      end_at_ms: report.end_at_ms,
      call_count: report.call_count,
      actions: report.actions,
      signal_types: report.signal_types,
      missing: report.missing,
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

export function assertMeetingPlatformRuntimeHostReplay(report = {}) {
  if (report.accepted !== true) {
    throw new MeetingTimelineSdkError('meeting platform runtime host replay failed', {
      platform: report.platform,
      missing: report.missing,
    });
  }
  return report;
}

export function assertMeetingPlatformRuntimeHostReplayMatrix(matrix = {}) {
  const failed = (matrix.reports ?? []).filter((report) => report.accepted !== true);
  if (failed.length > 0) {
    throw new MeetingTimelineSdkError('meeting platform runtime host replay matrix failed', {
      failed: failed.map((report) => ({ platform: report.platform, missing: report.missing })),
    });
  }
  return matrix;
}
