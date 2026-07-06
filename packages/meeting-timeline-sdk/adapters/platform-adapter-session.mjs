import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  MEETING_PLATFORM_ADAPTER_LAUNCH_PLAN_SCHEMA,
  assertMeetingPlatformAdapterLaunchPlan,
} from './platform-adapter-launch-plan.mjs';

export const MEETING_PLATFORM_ADAPTER_SESSION_SCHEMA = 'meeting_platform_adapter_session';
export const MEETING_PLATFORM_ADAPTER_SESSION_EVENT_SCHEMA = 'meeting_platform_adapter_session_event';
export const MEETING_PLATFORM_ADAPTER_SESSION_SCHEMA_VERSION = 1;

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

function nowMs(options = {}) {
  const clock = firstNonEmpty(options.clock, options.now);
  if (typeof clock === 'function') return Number(clock());
  return Date.now();
}

function launchPlanFrom(input = {}, options = {}) {
  const plan = input?.schema === MEETING_PLATFORM_ADAPTER_LAUNCH_PLAN_SCHEMA
    ? input
    : firstNonEmpty(options.launchPlan, options.launch_plan, input.launchPlan, input.launch_plan, input.plan);
  return assertMeetingPlatformAdapterLaunchPlan(plan ?? input);
}

function sessionId(plan = {}, options = {}) {
  return String(firstNonEmpty(
    options.sessionId,
    options.session_id,
    plan.detected_meeting?.meeting_id,
    plan.detected_meeting?.external_meeting_id,
    `${plan.platform || 'meeting'}-${plan.selected_surface || 'surface'}`,
  ));
}

function observePayload(plan = {}, input = {}, options = {}) {
  const candidate = compactObject({
    platform: plan.platform,
    url: firstNonEmpty(input.url, input.href, plan.current_url),
    title: firstNonEmpty(input.title, plan.detected_meeting?.title),
    active: firstNonEmpty(input.active, true),
    in_meeting: firstNonEmpty(input.in_meeting, input.inMeeting, true),
    surface: plan.selected_surface,
    source: firstNonEmpty(options.source, input.source, 'meeting_platform_adapter_session'),
    captured_at_ms: firstNonEmpty(input.captured_at_ms, input.capturedAtMs, options.captured_at_ms, options.capturedAtMs, nowMs(options)),
  });
  return compactObject({
    platform: plan.platform,
    source: candidate.source,
    captured_at_ms: candidate.captured_at_ms,
    current_url: candidate.url,
    surface: plan.selected_surface,
    detected_meeting: plan.detected_meeting,
    candidates: asArray(firstNonEmpty(input.candidates, input.tabs, input.windows, [candidate])),
  });
}

function markPayload(plan = {}, input = {}, options = {}) {
  return {
    ...input,
    platform: firstNonEmpty(input.platform, plan.platform),
    captured_at_ms: firstNonEmpty(input.captured_at_ms, input.capturedAtMs, options.captured_at_ms, options.capturedAtMs, nowMs(options)),
    source: firstNonEmpty(input.source, options.source, 'meeting_platform_adapter_session'),
    surface: firstNonEmpty(input.surface, plan.selected_surface),
  };
}

function trackPayload(plan = {}, input = {}, options = {}) {
  return {
    ...input,
    platform: firstNonEmpty(input.platform, plan.platform),
    captured_at_ms: firstNonEmpty(input.captured_at_ms, input.capturedAtMs, options.captured_at_ms, options.capturedAtMs, nowMs(options)),
    source: firstNonEmpty(input.source, options.source, 'meeting_platform_adapter_session'),
  };
}

async function callInsert(client = {}, platform, payload = {}, options = {}) {
  if (typeof client.insertAnnotation === 'function') {
    if (options.platformArgument === false || options.platform_argument === false || client.insertAnnotation.length < 2) {
      return client.insertAnnotation(payload, options);
    }
    return client.insertAnnotation(platform, payload, options);
  }
  if (typeof client.insertMark === 'function') return client.insertMark(payload, options);
  throw new MeetingTimelineSdkError('Adapter session client cannot insert annotations', {
    required_method: 'insertAnnotation',
  });
}

async function callObserve(client = {}, payload = {}, options = {}) {
  if (typeof client.observePlatformCandidates === 'function') return client.observePlatformCandidates(payload, options);
  if (typeof client.observeCandidates === 'function') return client.observeCandidates(payload.candidates ?? [], options);
  if (typeof client.observe === 'function') return client.observe(payload, options);
  throw new MeetingTimelineSdkError('Adapter session client cannot observe platform candidates', {
    required_method: 'observePlatformCandidates',
  });
}

async function callTrack(client = {}, method, platform, payload = {}, options = {}) {
  if (typeof client[method] !== 'function') {
    throw new MeetingTimelineSdkError(`Adapter session client cannot call ${method}`, {
      required_method: method,
    });
  }
  if (options.platformArgument === false || options.platform_argument === false || client[method].length < 2) {
    return client[method](payload, options);
  }
  return client[method](platform, payload, options);
}

function event(action, session, payload = {}, result) {
  return compactObject({
    type: 'meeting_platform_adapter_session_event',
    schema: MEETING_PLATFORM_ADAPTER_SESSION_EVENT_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_SESSION_SCHEMA_VERSION,
    session_id: session.id,
    action,
    platform: session.platform,
    selected_surface: session.selected_surface,
    payload,
    result,
    captured_at_ms: payload.captured_at_ms,
  });
}

export function createMeetingPlatformAdapterSession(launchPlanOrInput = {}, clientOrOptions = {}, options = {}) {
  const client = firstNonEmpty(options.client, options.sdk, clientOrOptions.client, clientOrOptions.sdk, clientOrOptions);
  const defaults = {
    ...options,
    ...(clientOrOptions.options ?? {}),
  };
  const plan = launchPlanFrom(launchPlanOrInput, defaults);
  const state = {
    axis_observed: false,
    last_observe_event: undefined,
    last_annotation_event: undefined,
  };
  const session = {
    type: 'meeting_platform_adapter_session',
    schema: MEETING_PLATFORM_ADAPTER_SESSION_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_SESSION_SCHEMA_VERSION,
    id: sessionId(plan, defaults),
    platform: plan.platform,
    selected_surface: plan.selected_surface,
    launch_plan: plan,
    axis_contract: plan.axis_contract,
    runtime_actions: plan.runtime_actions,
    getState() {
      return {
        ...state,
        session_id: session.id,
        platform: session.platform,
        selected_surface: session.selected_surface,
      };
    },
    async observeAxis(input = {}, observeOptions = {}) {
      const mergedOptions = { ...defaults, ...observeOptions };
      const payload = observePayload(plan, input, mergedOptions);
      const result = await callObserve(client, payload, mergedOptions);
      state.axis_observed = true;
      state.last_observe_event = event('observe_axis', session, payload, result);
      return state.last_observe_event;
    },
    async insertAnnotation(input = {}, markOptions = {}) {
      const mergedOptions = { ...defaults, ...markOptions };
      if (state.axis_observed !== true && mergedOptions.allowUnobservedAxis !== true && mergedOptions.allow_unobserved_axis !== true) {
        throw new MeetingTimelineSdkError('Adapter session must observe local axis before inserting realtime annotation', {
          code: 'adapter_session_axis_not_observed',
          session_id: session.id,
          platform: session.platform,
        });
      }
      const payload = markPayload(plan, input, mergedOptions);
      const result = await callInsert(client, plan.platform, payload, mergedOptions);
      state.last_annotation_event = event('insert_annotation', session, payload, result);
      return state.last_annotation_event;
    },
    async insertMark(input = {}, markOptions = {}) {
      return session.insertAnnotation(input, markOptions);
    },
    async speakerTrack(input = {}, trackOptions = {}) {
      const mergedOptions = { ...defaults, ...trackOptions };
      const payload = trackPayload(plan, input, mergedOptions);
      const result = await callTrack(client, 'speakerTrack', plan.platform, payload, mergedOptions);
      return event('speaker_track', session, payload, result);
    },
    async participantTrack(input = {}, trackOptions = {}) {
      const mergedOptions = { ...defaults, ...trackOptions };
      const payload = trackPayload(plan, input, mergedOptions);
      const result = await callTrack(client, 'participantTrack', plan.platform, payload, mergedOptions);
      return event('participant_track', session, payload, result);
    },
    async providerReconcile(input = {}, reconcileOptions = {}) {
      if (typeof client.ingestProvider !== 'function') {
        throw new MeetingTimelineSdkError('Adapter session client cannot ingest provider reconcile events', {
          required_method: 'ingestProvider',
        });
      }
      const mergedOptions = { ...defaults, ...reconcileOptions };
      const payload = trackPayload(plan, input, mergedOptions);
      const result = client.ingestProvider.length < 2
        ? await client.ingestProvider(payload, mergedOptions)
        : await client.ingestProvider(plan.platform, payload, mergedOptions);
      return event('provider_reconcile', session, payload, result);
    },
  };
  return session;
}

export function buildMeetingPlatformAdapterSessionHandoff(launchPlanOrInput = {}, options = {}) {
  const plan = launchPlanFrom(launchPlanOrInput, options);
  return {
    type: 'meeting_platform_adapter_session_handoff',
    schema: 'meeting_platform_adapter_session_handoff',
    schema_version: MEETING_PLATFORM_ADAPTER_SESSION_SCHEMA_VERSION,
    platform: plan.platform,
    selected_surface: plan.selected_surface,
    session_factory: 'createMeetingPlatformAdapterSession',
    required_client_methods: ['observePlatformCandidates', 'insertAnnotation'],
    optional_client_methods: ['speakerTrack', 'participantTrack', 'ingestProvider'],
    launch_plan_schema: plan.schema,
    runtime_actions: plan.runtime_actions,
    axis_contract: plan.axis_contract,
    next_actions: [
      'create_session_from_launch_plan',
      'call_session_observeAxis_before_first_mark',
      'call_session_insertAnnotation_for_realtime_marks',
    ],
  };
}
