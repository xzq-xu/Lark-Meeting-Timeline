import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  MEETING_PLATFORM_ADAPTER_CANDIDATE_LAUNCH_PLAN_SCHEMA,
  MEETING_PLATFORM_ADAPTER_LAUNCH_PLAN_SCHEMA,
  assertMeetingPlatformAdapterCandidateLaunchPlan,
  assertMeetingPlatformAdapterLaunchPlan,
} from './platform-adapter-launch-plan.mjs';
import {
  MEETING_PLATFORM_ADAPTER_RUNTIME_TARGET_SCHEMA,
  assertMeetingPlatformAdapterRuntimeTarget,
} from './platform-adapter-runtime-recipe.mjs';

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

function isLaunchPlan(value = {}) {
  return value?.schema === MEETING_PLATFORM_ADAPTER_LAUNCH_PLAN_SCHEMA;
}

function isCandidateLaunchPlan(value = {}) {
  return value?.schema === MEETING_PLATFORM_ADAPTER_CANDIDATE_LAUNCH_PLAN_SCHEMA;
}

function isRuntimeTarget(value = {}) {
  return value?.schema === MEETING_PLATFORM_ADAPTER_RUNTIME_TARGET_SCHEMA;
}

function planKind(plan = {}) {
  if (isCandidateLaunchPlan(plan)) return 'candidate_launch_plan';
  if (isRuntimeTarget(plan)) return 'runtime_target';
  if (isLaunchPlan(plan)) return 'launch_plan';
  return 'unknown';
}

function sessionPlanFrom(input = {}, options = {}) {
  if (isRuntimeTarget(input)) return assertMeetingPlatformAdapterRuntimeTarget(input);
  if (isCandidateLaunchPlan(input)) return assertMeetingPlatformAdapterCandidateLaunchPlan(input);
  if (isLaunchPlan(input)) return assertMeetingPlatformAdapterLaunchPlan(input);
  const candidate = firstNonEmpty(
    options.runtimeTarget,
    options.runtime_target,
    options.candidateLaunchPlan,
    options.candidate_launch_plan,
    options.launchPlan,
    options.launch_plan,
    input.runtimeTarget,
    input.runtime_target,
    input.candidateLaunchPlan,
    input.candidate_launch_plan,
    input.launchPlan,
    input.launch_plan,
    input.target,
    input.plan,
    input,
  );
  if (isRuntimeTarget(candidate)) return assertMeetingPlatformAdapterRuntimeTarget(candidate);
  if (isCandidateLaunchPlan(candidate)) return assertMeetingPlatformAdapterCandidateLaunchPlan(candidate);
  if (isLaunchPlan(candidate)) return assertMeetingPlatformAdapterLaunchPlan(candidate);
  return assertMeetingPlatformAdapterLaunchPlan(candidate);
}

function nestedLaunchPlan(plan = {}) {
  return isCandidateLaunchPlan(plan) ? plan.launch_plan ?? {} : plan;
}

function detectedMeetingFrom(plan = {}) {
  return plan.detected_meeting ?? nestedLaunchPlan(plan).detected_meeting;
}

function currentUrlFrom(plan = {}) {
  return plan.current_url ?? plan.launch_input?.url ?? nestedLaunchPlan(plan).current_url;
}

function selectedEvidenceFrom(plan = {}) {
  return plan.selected_evidence ?? plan.launch_input?.preflight_evidence;
}

function axisContractFrom(plan = {}) {
  const axisContract = plan.axis_contract ?? {};
  return compactObject({
    ...axisContract,
    timestamp_field: firstNonEmpty(axisContract.timestamp_field, plan.timestamp_field, plan.runtime_contract?.timestamp_field),
    provider_events_block_realtime: firstNonEmpty(
      axisContract.provider_events_block_realtime,
      plan.runtime_contract?.provider_events_block_realtime,
    ),
    transcript_blocks_realtime: firstNonEmpty(
      axisContract.transcript_blocks_realtime,
      plan.runtime_contract?.transcript_blocks_realtime,
    ),
  });
}

function sessionId(plan = {}, options = {}) {
  const detectedMeeting = detectedMeetingFrom(plan) ?? {};
  return String(firstNonEmpty(
    options.sessionId,
    options.session_id,
    detectedMeeting.meeting_id,
    detectedMeeting.external_meeting_id,
    `${plan.platform || 'meeting'}-${plan.selected_surface || 'surface'}`,
  ));
}

function observePayload(plan = {}, input = {}, options = {}) {
  const evidence = selectedEvidenceFrom(plan);
  const detectedMeeting = detectedMeetingFrom(plan);
  const candidate = compactObject({
    platform: plan.platform,
    url: firstNonEmpty(input.url, input.href, currentUrlFrom(plan)),
    title: firstNonEmpty(input.title, detectedMeeting?.title),
    active: firstNonEmpty(input.active, true),
    in_meeting: firstNonEmpty(input.in_meeting, input.inMeeting, evidence?.interaction_in_call, evidence?.interaction?.in_call, true),
    surface: plan.selected_surface,
    source: firstNonEmpty(options.source, input.source, 'meeting_platform_adapter_session'),
    captured_at_ms: firstNonEmpty(input.captured_at_ms, input.capturedAtMs, options.captured_at_ms, options.capturedAtMs, nowMs(options)),
    interaction: evidence?.interaction,
    semantic_signal_types: evidence?.semantic_signal_types,
    control_signal_summary: evidence?.control_signal_summary,
    control_signal_gaps: evidence?.control_signal_gaps,
    control_signal_gap_summary: evidence?.control_signal_gap_summary,
    active_speaker_candidate: evidence?.active_speaker_candidate,
  });
  return compactObject({
    platform: plan.platform,
    source: candidate.source,
    captured_at_ms: candidate.captured_at_ms,
    current_url: candidate.url,
    surface: plan.selected_surface,
    detected_meeting: detectedMeeting,
    preflight_evidence: evidence,
    candidates: asArray(firstNonEmpty(input.candidates, input.tabs, input.windows, [candidate])),
  });
}

function rawSignalValidationPayload(plan = {}, input = {}, options = {}) {
  const capturedAtMs = firstNonEmpty(input.captured_at_ms, input.capturedAtMs, options.captured_at_ms, options.capturedAtMs, nowMs(options));
  const evidence = selectedEvidenceFrom(plan);
  const detectedMeeting = detectedMeetingFrom(plan);
  const snapshot = compactObject({
    kind: 'meeting_app_snapshot',
    platform: plan.platform,
    source: firstNonEmpty(input.source, options.source, 'meeting_platform_adapter_session'),
    url: firstNonEmpty(input.url, input.href, currentUrlFrom(plan)),
    title: firstNonEmpty(input.title, detectedMeeting?.title),
    observed_at_ms: capturedAtMs,
    in_meeting: firstNonEmpty(input.in_meeting, input.inMeeting, evidence?.interaction_in_call, evidence?.interaction?.in_call, true),
    active_speaker: firstNonEmpty(input.active_speaker, input.activeSpeaker, input.speaker, evidence?.active_speaker_candidate),
    interaction: evidence?.interaction,
    semantic_signal_types: evidence?.semantic_signal_types,
    control_signal_summary: evidence?.control_signal_summary,
    control_signal_gaps: evidence?.control_signal_gaps,
    control_signal_gap_summary: evidence?.control_signal_gap_summary,
    current_meeting: detectedMeeting,
  });
  return compactObject({
    platform: plan.platform,
    source: snapshot.source,
    captured_at_ms: capturedAtMs,
    validation: plan.raw_signal_validation,
    raw_signal_validation_path: plan.raw_signal_validation?.path,
    expected_runtime_actions: plan.raw_signal_validation?.runtime_actions,
    preflight_evidence: evidence,
    raw_signals: asArray(firstNonEmpty(
      input.raw_signals,
      input.rawSignals,
      input.samples,
      input.signals,
      [snapshot],
    )),
    filter_active_speaker_samples: true,
  });
}

function adapterSelectionPayload(plan = {}, input = {}, options = {}) {
  const capturedAtMs = firstNonEmpty(input.captured_at_ms, input.capturedAtMs, options.captured_at_ms, options.capturedAtMs, nowMs(options));
  const runtimeAction = asArray(plan.runtime_actions).find((action) => action.id === 'read_adapter_selection');
  return compactObject({
    platform: plan.platform,
    source: firstNonEmpty(input.source, options.source, 'meeting_platform_adapter_session'),
    captured_at_ms: capturedAtMs,
    selected_surface: plan.selected_surface,
    adapter_selection: plan.adapter_selection,
    adapter_selection_path: plan.adapter_selection?.path ?? runtimeAction?.artifact_path,
    axis_source: plan.adapter_selection?.axis_source ?? plan.axis_contract?.adapter_selection_axis_source,
    axis_surface: plan.adapter_selection?.axis_surface ?? plan.axis_contract?.adapter_selection_axis_surface,
    timestamp_field: plan.adapter_selection?.timestamp_field ?? plan.axis_contract?.timestamp_field,
    runtime_action: runtimeAction,
  });
}

function markPayload(plan = {}, input = {}, options = {}) {
  const evidence = selectedEvidenceFrom(plan);
  return {
    ...input,
    platform: firstNonEmpty(input.platform, plan.platform),
    captured_at_ms: firstNonEmpty(input.captured_at_ms, input.capturedAtMs, options.captured_at_ms, options.capturedAtMs, nowMs(options)),
    source: firstNonEmpty(input.source, options.source, 'meeting_platform_adapter_session'),
    surface: firstNonEmpty(input.surface, plan.selected_surface),
    preflight_evidence: firstNonEmpty(input.preflight_evidence, input.preflightEvidence, evidence),
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

async function callValidateRawSignal(client = {}, payload = {}, options = {}) {
  if (typeof client.platformRawSignalBatch === 'function') return client.platformRawSignalBatch(payload, options);
  if (typeof client.rawSignalBatch === 'function') return client.rawSignalBatch(payload, options);
  if (typeof client.validateRawSignal === 'function') return client.validateRawSignal(payload, options);
  return {
    ok: true,
    mode: 'launch_plan_contract',
    status: payload.validation?.status ?? 'ready',
    runtime_actions: payload.expected_runtime_actions,
    signal_count: payload.raw_signals?.length ?? 0,
  };
}

async function callAdapterSelection(client = {}, platform, payload = {}, options = {}) {
  if (typeof client.platformAdapterSelection === 'function') {
    if (options.platformArgument === false || options.platform_argument === false || client.platformAdapterSelection.length < 2) {
      return client.platformAdapterSelection(payload, options);
    }
    return client.platformAdapterSelection(platform, payload, options);
  }
  if (typeof client.adapterSelection === 'function') {
    if (options.platformArgument === false || options.platform_argument === false || client.adapterSelection.length < 2) {
      return client.adapterSelection(payload, options);
    }
    return client.adapterSelection(platform, payload, options);
  }
  return {
    ok: true,
    mode: 'launch_plan_contract',
    status: payload.adapter_selection?.ready === true ? 'ready' : 'not_available',
    adapter_selection: payload.adapter_selection,
  };
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
  const plan = sessionPlanFrom(launchPlanOrInput, defaults);
  const launchPlan = nestedLaunchPlan(plan);
  const axisContract = axisContractFrom(plan);
  const kind = planKind(plan);
  const state = {
    adapter_selection_read: false,
    raw_signal_validated: false,
    axis_observed: false,
    last_adapter_selection_event: undefined,
    last_raw_signal_event: undefined,
    last_observe_event: undefined,
    last_annotation_event: undefined,
  };
  const session = {
    type: 'meeting_platform_adapter_session',
    schema: MEETING_PLATFORM_ADAPTER_SESSION_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_SESSION_SCHEMA_VERSION,
    id: sessionId(plan, defaults),
    plan_kind: kind,
    platform: plan.platform,
    selected_surface: plan.selected_surface,
    host_kind: plan.host_kind ?? launchPlan.host_kind,
    bridge_kind: plan.bridge_kind ?? launchPlan.bridge_kind,
    adapter_module: plan.adapter_module ?? launchPlan.adapter_module,
    adapter_selection: plan.adapter_selection ?? launchPlan.adapter_selection,
    selected_evidence: selectedEvidenceFrom(plan),
    launch_plan: isLaunchPlan(plan) ? plan : isCandidateLaunchPlan(plan) ? launchPlan : undefined,
    candidate_launch_plan: isCandidateLaunchPlan(plan) ? plan : undefined,
    runtime_target: isRuntimeTarget(plan) ? plan : undefined,
    axis_contract: axisContract,
    runtime_actions: plan.runtime_actions ?? launchPlan.runtime_actions,
    getState() {
      return {
        ...state,
        session_id: session.id,
        plan_kind: session.plan_kind,
        platform: session.platform,
        selected_surface: session.selected_surface,
        host_kind: session.host_kind,
        bridge_kind: session.bridge_kind,
        adapter_selection: session.adapter_selection,
        selected_evidence: session.selected_evidence,
      };
    },
    async readAdapterSelection(input = {}, selectionOptions = {}) {
      const mergedOptions = { ...defaults, ...selectionOptions };
      const payload = adapterSelectionPayload(plan, input, mergedOptions);
      const result = await callAdapterSelection(client, plan.platform, payload, mergedOptions);
      state.adapter_selection_read = true;
      state.last_adapter_selection_event = event('read_adapter_selection', session, payload, result);
      return state.last_adapter_selection_event;
    },
    async validateRawSignal(input = {}, validationOptions = {}) {
      const mergedOptions = { ...defaults, ...validationOptions };
      if (state.adapter_selection_read !== true && mergedOptions.readAdapterSelection !== false && mergedOptions.read_adapter_selection !== false) {
        await session.readAdapterSelection(input, mergedOptions);
      }
      const payload = rawSignalValidationPayload(plan, input, mergedOptions);
      const result = await callValidateRawSignal(client, payload, mergedOptions);
      state.raw_signal_validated = true;
      state.last_raw_signal_event = event('validate_raw_signal', session, payload, result);
      return state.last_raw_signal_event;
    },
    async observeAxis(input = {}, observeOptions = {}) {
      const mergedOptions = { ...defaults, ...observeOptions };
      if (state.adapter_selection_read !== true && mergedOptions.readAdapterSelection !== false && mergedOptions.read_adapter_selection !== false) {
        await session.readAdapterSelection(input, mergedOptions);
      }
      if (state.raw_signal_validated !== true && mergedOptions.validateRawSignal !== false && mergedOptions.validate_raw_signal !== false) {
        await session.validateRawSignal(input, mergedOptions);
      }
      const payload = observePayload(plan, input, mergedOptions);
      const result = await callObserve(client, payload, mergedOptions);
      state.axis_observed = true;
      state.last_observe_event = event('observe_axis', session, payload, result);
      return state.last_observe_event;
    },
    async insertAnnotation(input = {}, markOptions = {}) {
      const mergedOptions = { ...defaults, ...markOptions };
      if (state.raw_signal_validated !== true && mergedOptions.allowUnvalidatedRawSignal !== true && mergedOptions.allow_unvalidated_raw_signal !== true) {
        throw new MeetingTimelineSdkError('Adapter session must validate raw signals before inserting realtime annotation', {
          code: 'adapter_session_raw_signal_not_validated',
          session_id: session.id,
          platform: session.platform,
        });
      }
      if (state.axis_observed !== true && mergedOptions.allowUnobservedAxis !== true && mergedOptions.allow_unobserved_axis !== true) {
        throw new MeetingTimelineSdkError('Adapter session must observe local axis before inserting realtime annotation', {
          code: 'adapter_session_axis_not_observed',
          session_id: session.id,
          platform: session.platform,
        });
      }
      if (state.adapter_selection_read !== true && mergedOptions.allowUnreadAdapterSelection !== true && mergedOptions.allow_unread_adapter_selection !== true) {
        throw new MeetingTimelineSdkError('Adapter session must read adapter selection before inserting realtime annotation', {
          code: 'adapter_session_adapter_selection_not_read',
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
  const plan = sessionPlanFrom(launchPlanOrInput, options);
  const launchPlan = nestedLaunchPlan(plan);
  const kind = planKind(plan);
  const axisContract = axisContractFrom(plan);
  return {
    type: 'meeting_platform_adapter_session_handoff',
    schema: 'meeting_platform_adapter_session_handoff',
    schema_version: MEETING_PLATFORM_ADAPTER_SESSION_SCHEMA_VERSION,
    platform: plan.platform,
    selected_surface: plan.selected_surface,
    plan_kind: kind,
    host_kind: plan.host_kind ?? launchPlan.host_kind,
    bridge_kind: plan.bridge_kind ?? launchPlan.bridge_kind,
    session_factory: 'createMeetingPlatformAdapterSession',
    required_client_methods: ['observePlatformCandidates', 'insertAnnotation'],
    optional_client_methods: ['platformAdapterSelection', 'adapterSelection', 'platformRawSignalBatch', 'rawSignalBatch', 'speakerTrack', 'participantTrack', 'ingestProvider'],
    input_schema: plan.schema,
    launch_plan_schema: isLaunchPlan(plan) ? plan.schema : undefined,
    candidate_launch_plan_schema: isCandidateLaunchPlan(plan) ? plan.schema : undefined,
    runtime_target_schema: isRuntimeTarget(plan) ? plan.schema : undefined,
    selected_evidence: selectedEvidenceFrom(plan),
    runtime_actions: plan.runtime_actions ?? launchPlan.runtime_actions,
    axis_contract: axisContract,
    next_actions: [
      kind === 'runtime_target'
        ? 'create_session_from_runtime_target'
        : kind === 'candidate_launch_plan'
          ? 'create_session_from_candidate_launch_plan'
          : 'create_session_from_launch_plan',
      'call_session_readAdapterSelection_before_runtime_wiring',
      'call_session_validateRawSignal_before_adapter_preflight',
      'call_session_observeAxis_before_first_mark',
      'call_session_insertAnnotation_for_realtime_marks',
    ],
  };
}
