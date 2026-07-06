import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  MEETING_PLATFORM_ADAPTER_RUNNER_SCHEMA,
  createMeetingPlatformAdapterRunner,
} from './platform-adapter-runner.mjs';

export const MEETING_PLATFORM_ADAPTER_MESSAGE_BRIDGE_SCHEMA = 'meeting_platform_adapter_message_bridge';
export const MEETING_PLATFORM_ADAPTER_MESSAGE_BRIDGE_EVENT_SCHEMA = 'meeting_platform_adapter_message_bridge_event';
export const MEETING_PLATFORM_ADAPTER_MESSAGE_BRIDGE_SCHEMA_VERSION = 1;

const OPEN_ACTIONS = new Set([
  'open_session',
  'observe_candidates',
  'observe_platform_candidates',
  'meeting_timeline.open_session',
  'meeting_timeline.observe_candidates',
  'meeting_timeline.observe_platform_candidates',
]);

const INSERT_ACTIONS = new Set([
  'insert_mark',
  'insert_annotation',
  'annotation',
  'insertAnnotation',
  'meeting_timeline.insert_mark',
  'meeting_timeline.insert_annotation',
  'meeting_timeline.annotation',
]);

const SPEAKER_ACTIONS = new Set([
  'speaker_track',
  'speaker',
  'meeting_timeline.speaker_track',
]);

const PARTICIPANT_ACTIONS = new Set([
  'participant_track',
  'participant',
  'meeting_timeline.participant_track',
]);

const PROVIDER_ACTIONS = new Set([
  'provider_event',
  'ingest_provider',
  'provider_reconcile',
  'meeting_timeline.provider_event',
  'meeting_timeline.ingest_provider',
]);

const STATUS_ACTIONS = new Set([
  'status',
  'get_state',
  'meeting_timeline.status',
  'meeting_timeline.get_state',
]);

const RESET_ACTIONS = new Set([
  'reset',
  'meeting_timeline.reset',
]);

const LAUNCH_PLAN_ACTIONS = new Set([
  'launch_plan',
  'meeting_timeline.launch_plan',
]);

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date);
}

function isRunner(value) {
  return value?.schema === MEETING_PLATFORM_ADAPTER_RUNNER_SCHEMA
    && typeof value.open === 'function'
    && typeof value.insertAnnotation === 'function';
}

function bridgeId(input = {}, options = {}) {
  return String(firstNonEmpty(
    options.bridgeId,
    options.bridge_id,
    input.bridgeId,
    input.bridge_id,
    'meeting-platform-adapter-message-bridge',
  ));
}

function messageType(message = {}) {
  return String(firstNonEmpty(
    message.type,
    message.action,
    message.kind,
    message.message_type,
    message.messageType,
    '',
  )).trim();
}

function requestId(message = {}, payload = {}) {
  return firstNonEmpty(
    message.id,
    message.requestId,
    message.request_id,
    payload.id,
    payload.requestId,
    payload.request_id,
  );
}

function payloadFrom(message = {}) {
  const payload = firstNonEmpty(message.payload, message.data, message.detail, message.body);
  return isPlainObject(payload) ? payload : {};
}

function scoreCandidate(candidate = {}, index = 0) {
  const active = candidate.active === true || candidate.current === true || candidate.selected === true ? 100 : 0;
  const inMeeting = candidate.in_meeting === true || candidate.inMeeting === true || candidate.meeting?.in_meeting === true ? 40 : 0;
  const audible = candidate.audible === true || candidate.has_audio === true || candidate.hasAudio === true ? 10 : 0;
  return active + inMeeting + audible - index;
}

function compactCandidate(value = {}) {
  if (typeof value === 'string' || value instanceof URL) return { url: String(value) };
  if (!isPlainObject(value)) return {};
  return compactObject({
    platform: firstNonEmpty(value.platform, value.provider, value.adapter, value.meeting?.platform),
    url: firstNonEmpty(
      value.url,
      value.href,
      value.meeting_url,
      value.meetingUrl,
      value.join_url,
      value.joinUrl,
      value.meeting?.url,
      value.meeting?.meeting_url,
      value.meeting?.meetingUrl,
      value.meeting?.join_url,
      value.meeting?.joinUrl,
      value.tab?.url,
      value.window?.url,
    ),
    title: firstNonEmpty(
      value.title,
      value.topic,
      value.name,
      value.meeting?.title,
      value.meeting?.topic,
      value.tab?.title,
      value.window?.title,
    ),
    active: firstNonEmpty(value.active, value.current, value.selected, value.tab?.active, value.window?.active),
    audible: firstNonEmpty(value.audible, value.has_audio, value.hasAudio, value.tab?.audible),
    in_meeting: firstNonEmpty(value.in_meeting, value.inMeeting, value.meeting?.in_meeting, value.meeting?.inMeeting),
  });
}

function candidateListFrom(payload = {}, message = {}, options = {}) {
  const candidates = [];
  const push = (value) => {
    const candidate = compactCandidate(value);
    if (candidate.url || candidate.platform || candidate.title) candidates.push(candidate);
  };
  push(options);
  push(message);
  push(payload);
  push(payload.input);
  push(payload.meeting);
  push(payload.current_meeting);
  push(payload.currentMeeting);
  push(payload.tab);
  push(payload.window);
  for (const candidate of asArray(payload.candidates)) push(candidate);
  for (const tab of asArray(payload.tabs)) push(tab);
  for (const window of asArray(payload.windows)) {
    push(window);
    for (const tab of asArray(window?.tabs)) push({ ...tab, window });
  }
  return candidates
    .map((candidate, index) => ({ ...candidate, score: scoreCandidate(candidate, index) }))
    .sort((left, right) => right.score - left.score);
}

function launchInputFrom(payload = {}, message = {}, options = {}) {
  const candidates = candidateListFrom(payload, message, options);
  const best = candidates[0] ?? {};
  return compactObject({
    platform: firstNonEmpty(payload.platform, message.platform, options.platform, best.platform),
    url: firstNonEmpty(payload.url, payload.href, message.url, options.url, best.url),
    title: firstNonEmpty(payload.title, payload.topic, message.title, options.title, best.title),
    active: firstNonEmpty(payload.active, best.active),
    in_meeting: firstNonEmpty(payload.in_meeting, payload.inMeeting, best.in_meeting, true),
    captured_at_ms: firstNonEmpty(payload.captured_at_ms, payload.capturedAtMs, message.captured_at_ms, options.captured_at_ms),
    candidates,
    tabs: payload.tabs,
    windows: payload.windows,
  });
}

function markInputFrom(payload = {}, message = {}, options = {}) {
  const mark = firstNonEmpty(payload.mark, payload.annotation, payload.input, payload.payload);
  const base = isPlainObject(mark) ? mark : payload;
  return compactObject({
    ...base,
    platform: firstNonEmpty(base.platform, payload.platform, message.platform, options.platform),
    captured_at_ms: firstNonEmpty(
      base.captured_at_ms,
      base.capturedAtMs,
      payload.captured_at_ms,
      payload.capturedAtMs,
      message.captured_at_ms,
      options.captured_at_ms,
    ),
    source: firstNonEmpty(base.source, payload.source, options.source, 'meeting_platform_adapter_message_bridge'),
  });
}

function trackInputFrom(payload = {}, message = {}, options = {}) {
  const track = firstNonEmpty(payload.track, payload.speaker, payload.participant, payload.input, payload.payload);
  const base = isPlainObject(track) ? track : payload;
  return compactObject({
    ...base,
    platform: firstNonEmpty(base.platform, payload.platform, message.platform, options.platform),
    captured_at_ms: firstNonEmpty(
      base.captured_at_ms,
      base.capturedAtMs,
      payload.captured_at_ms,
      payload.capturedAtMs,
      message.captured_at_ms,
      options.captured_at_ms,
    ),
    source: firstNonEmpty(base.source, payload.source, options.source, 'meeting_platform_adapter_message_bridge'),
  });
}

function providerInputFrom(payload = {}, message = {}, options = {}) {
  const provider = firstNonEmpty(payload.event, payload.provider_event, payload.input, payload.payload);
  const base = isPlainObject(provider) ? provider : payload;
  return compactObject({
    ...base,
    platform: firstNonEmpty(base.platform, payload.platform, message.platform, options.platform),
    captured_at_ms: firstNonEmpty(
      base.captured_at_ms,
      base.capturedAtMs,
      payload.captured_at_ms,
      payload.capturedAtMs,
      message.captured_at_ms,
      options.captured_at_ms,
    ),
    source: firstNonEmpty(base.source, payload.source, options.source, 'meeting_platform_adapter_message_bridge'),
  });
}

function shouldAutoOpen(payload = {}, options = {}) {
  if (options.autoOpen === false || options.auto_open === false) return false;
  if (options.autoOpenOnInsert === false || options.auto_open_on_insert === false) return false;
  return Boolean(
    payload.url
    || payload.href
    || payload.platform
    || payload.meeting_url
    || payload.meetingUrl
    || payload.tab
    || payload.window
    || payload.tabs
    || payload.windows
    || payload.candidates
  );
}

function event(action, bridge, message = {}, payload = {}, result, handled = true) {
  const state = bridge.runner?.getState?.();
  return compactObject({
    type: 'meeting_platform_adapter_message_bridge_event',
    schema: MEETING_PLATFORM_ADAPTER_MESSAGE_BRIDGE_EVENT_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_MESSAGE_BRIDGE_SCHEMA_VERSION,
    handled,
    bridge_id: bridge.id,
    action,
    message_type: messageType(message),
    request_id: requestId(message, payload),
    platform: result?.platform ?? result?.payload?.platform ?? state?.current_launch_plan?.platform,
    selected_surface: result?.selected_surface ?? result?.payload?.selected_surface ?? state?.current_launch_plan?.selected_surface,
    result,
    runner_state: state,
  });
}

async function ensureOpenForBridge(bridge, payload = {}, message = {}, options = {}) {
  if (bridge.runner.getState().current_session) return undefined;
  if (!shouldAutoOpen(payload, options)) return undefined;
  return bridge.runner.open(launchInputFrom(payload, message, options), options);
}

export function createMeetingPlatformAdapterMessageBridge(manifestOrRunner = {}, clientOrOptions = {}, options = {}) {
  const runner = isRunner(manifestOrRunner)
    ? manifestOrRunner
    : createMeetingPlatformAdapterRunner(manifestOrRunner, clientOrOptions, options);
  const defaults = isRunner(manifestOrRunner)
    ? { ...clientOrOptions, ...options }
    : { ...options };
  const bridge = {
    type: 'meeting_platform_adapter_message_bridge',
    schema: MEETING_PLATFORM_ADAPTER_MESSAGE_BRIDGE_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_MESSAGE_BRIDGE_SCHEMA_VERSION,
    id: bridgeId(manifestOrRunner, defaults),
    runner,
    getState() {
      return compactObject({
        type: 'meeting_platform_adapter_message_bridge_state',
        schema: 'meeting_platform_adapter_message_bridge_state',
        schema_version: MEETING_PLATFORM_ADAPTER_MESSAGE_BRIDGE_SCHEMA_VERSION,
        bridge_id: bridge.id,
        runner: runner.getState(),
      });
    },
    async handleMessage(message = {}, messageOptions = {}) {
      const mergedOptions = { ...defaults, ...messageOptions };
      const type = messageType(message);
      const payload = payloadFrom(message);
      if (OPEN_ACTIONS.has(type)) {
        const result = await runner.open(launchInputFrom(payload, message, mergedOptions), mergedOptions);
        return event('open_session', bridge, message, payload, result);
      }
      if (INSERT_ACTIONS.has(type)) {
        await ensureOpenForBridge(bridge, payload, message, mergedOptions);
        const result = await runner.insertAnnotation(markInputFrom(payload, message, mergedOptions), mergedOptions);
        return event('insert_annotation', bridge, message, payload, result);
      }
      if (SPEAKER_ACTIONS.has(type)) {
        await ensureOpenForBridge(bridge, payload, message, mergedOptions);
        const result = await runner.speakerTrack(trackInputFrom(payload, message, mergedOptions), mergedOptions);
        return event('speaker_track', bridge, message, payload, result);
      }
      if (PARTICIPANT_ACTIONS.has(type)) {
        await ensureOpenForBridge(bridge, payload, message, mergedOptions);
        const result = await runner.participantTrack(trackInputFrom(payload, message, mergedOptions), mergedOptions);
        return event('participant_track', bridge, message, payload, result);
      }
      if (PROVIDER_ACTIONS.has(type)) {
        await ensureOpenForBridge(bridge, payload, message, mergedOptions);
        const result = await runner.providerReconcile(providerInputFrom(payload, message, mergedOptions), mergedOptions);
        return event('provider_reconcile', bridge, message, payload, result);
      }
      if (STATUS_ACTIONS.has(type)) {
        return event('status', bridge, message, payload, bridge.getState());
      }
      if (RESET_ACTIONS.has(type)) {
        return event('reset', bridge, message, payload, runner.reset());
      }
      if (LAUNCH_PLAN_ACTIONS.has(type)) {
        const result = runner.launchPlan(launchInputFrom(payload, message, mergedOptions), mergedOptions);
        return event('launch_plan', bridge, message, payload, result);
      }
      return event('unsupported', bridge, message, payload, {
        reason: 'unsupported_message_type',
        message_type: type,
      }, false);
    },
    dispatchMessage(message = {}, messageOptions = {}) {
      return bridge.handleMessage(message, messageOptions);
    },
    async open(input = {}, openOptions = {}) {
      const result = await runner.open(input, { ...defaults, ...openOptions });
      return event('open_session', bridge, { type: 'open_session' }, input, result);
    },
    insertAnnotation(input = {}, markOptions = {}) {
      return bridge.handleMessage({
        type: 'meeting_timeline.insert_annotation',
        payload: { annotation: input },
      }, markOptions);
    },
    insertMark(input = {}, markOptions = {}) {
      return bridge.insertAnnotation(input, markOptions);
    },
  };
  return bridge;
}

export function buildMeetingPlatformAdapterMessageBridgeHandoff(manifestOrInput = {}, options = {}) {
  return {
    type: 'meeting_platform_adapter_message_bridge_handoff',
    schema: 'meeting_platform_adapter_message_bridge_handoff',
    schema_version: MEETING_PLATFORM_ADAPTER_MESSAGE_BRIDGE_SCHEMA_VERSION,
    bridge_factory: 'createMeetingPlatformAdapterMessageBridge',
    required_runner_factory: 'createMeetingPlatformAdapterRunner',
    install_manifest_schema: manifestOrInput?.schema,
    message_types: [
      'meeting_timeline.observe_candidates',
      'meeting_timeline.open_session',
      'meeting_timeline.insert_mark',
      'meeting_timeline.insert_annotation',
      'meeting_timeline.speaker_track',
      'meeting_timeline.participant_track',
      'meeting_timeline.provider_event',
      'meeting_timeline.status',
      'meeting_timeline.reset',
    ],
    runtime_sequence: [
      'background_or_native_host_sends_observe_candidates',
      'bridge_opens_adapter_runner_session',
      'content_or_device_sends_insert_mark_with_captured_at_ms',
      'bridge_routes_mark_to_current_runner_session',
    ],
    supported_surfaces: [
      'browser_extension_background',
      'browser_extension_content_script',
      'electron_webview_preload',
      'native_detector',
    ],
    auto_open_on_insert: options.autoOpenOnInsert !== false && options.auto_open_on_insert !== false,
    next_actions: [
      'create_message_bridge_from_adapter_install_manifest',
      'route_meeting_timeline_observe_candidates_from_host',
      'route_meeting_timeline_insert_mark_from_content_or_device',
    ],
  };
}

export function assertMeetingPlatformAdapterMessageBridge(bridge) {
  if (bridge?.schema !== MEETING_PLATFORM_ADAPTER_MESSAGE_BRIDGE_SCHEMA) {
    throw new MeetingTimelineSdkError('Invalid meeting platform adapter message bridge', {
      expected_schema: MEETING_PLATFORM_ADAPTER_MESSAGE_BRIDGE_SCHEMA,
      actual_schema: bridge?.schema,
    });
  }
  return bridge;
}
