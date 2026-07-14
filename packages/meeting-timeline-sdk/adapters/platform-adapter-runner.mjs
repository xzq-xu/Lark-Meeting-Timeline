import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  MEETING_PLATFORM_ADAPTER_CANDIDATE_LAUNCH_PLAN_SCHEMA,
  MEETING_PLATFORM_ADAPTER_LAUNCH_PLAN_SCHEMA,
  assertMeetingPlatformAdapterCandidateLaunchPlan,
  assertMeetingPlatformAdapterLaunchPlan,
  buildMeetingPlatformAdapterCandidateLaunchPlan,
  buildMeetingPlatformAdapterLaunchPlan,
} from './platform-adapter-launch-plan.mjs';
import {
  MEETING_PLATFORM_ADAPTER_SESSION_SCHEMA,
  createMeetingPlatformAdapterSession,
} from './platform-adapter-session.mjs';
import {
  MEETING_PLATFORM_ADAPTER_RUNTIME_MANIFEST_SCHEMA,
  MEETING_PLATFORM_ADAPTER_RUNTIME_TARGET_SCHEMA,
  assertMeetingPlatformAdapterRuntimeTarget,
  buildMeetingPlatformAdapterRuntimeTarget,
} from './platform-adapter-runtime-recipe.mjs';

export const MEETING_PLATFORM_ADAPTER_RUNNER_SCHEMA = 'meeting_platform_adapter_runner';
export const MEETING_PLATFORM_ADAPTER_OPEN_SESSION_EVENT_SCHEMA = 'meeting_platform_adapter_open_session_event';
export const MEETING_PLATFORM_ADAPTER_RUNNER_SCHEMA_VERSION = 1;

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function isLaunchPlan(value) {
  return value?.schema === MEETING_PLATFORM_ADAPTER_LAUNCH_PLAN_SCHEMA;
}

function isCandidateLaunchPlan(value) {
  return value?.schema === MEETING_PLATFORM_ADAPTER_CANDIDATE_LAUNCH_PLAN_SCHEMA;
}

function isInstallManifest(value) {
  return value?.schema === 'meeting_platform_adapter_install_manifest';
}

function isRuntimeManifest(value) {
  return value?.schema === MEETING_PLATFORM_ADAPTER_RUNTIME_MANIFEST_SCHEMA;
}

function isRuntimeTarget(value) {
  return value?.schema === MEETING_PLATFORM_ADAPTER_RUNTIME_TARGET_SCHEMA;
}

function isEmptyInput(value) {
  if (value == null || value === '') return true;
  return typeof value === 'object'
    && !Array.isArray(value)
    && !(value instanceof Date)
    && Object.keys(value).length === 0;
}

function isRunnerClient(value) {
  return Boolean(value)
    && typeof value === 'object'
    && (
      typeof value.observePlatformCandidates === 'function'
      || typeof value.observeCandidates === 'function'
      || typeof value.observe === 'function'
      || typeof value.insertAnnotation === 'function'
      || typeof value.insertMark === 'function'
    );
}

function clientFrom(clientOrOptions = {}, options = {}) {
  return firstNonEmpty(
    options.client,
    options.sdk,
    clientOrOptions.client,
    clientOrOptions.sdk,
    isRunnerClient(clientOrOptions) ? clientOrOptions : undefined,
  );
}

function defaultsFrom(clientOrOptions = {}, options = {}) {
  return {
    ...(isRunnerClient(clientOrOptions) ? {} : (clientOrOptions ?? {})),
    ...options,
  };
}

function installManifestFrom(manifestOrInput = {}, options = {}) {
  if (isInstallManifest(manifestOrInput)) return manifestOrInput;
  return firstNonEmpty(
    options.installManifest,
    options.install_manifest,
    options.manifest,
    manifestOrInput.installManifest,
    manifestOrInput.install_manifest,
    manifestOrInput.manifest,
    manifestOrInput.adapterInstallManifest,
    manifestOrInput.adapter_install_manifest,
  );
}

function runtimeManifestFrom(manifestOrInput = {}, options = {}) {
  if (isRuntimeManifest(manifestOrInput)) return manifestOrInput;
  return firstNonEmpty(
    options.runtimeManifest,
    options.runtime_manifest,
    options.runtime_manifest_input,
    manifestOrInput.runtimeManifest,
    manifestOrInput.runtime_manifest,
    manifestOrInput.runtime_manifest_input,
  );
}

function runnerId(input = {}, options = {}) {
  return String(firstNonEmpty(
    options.runnerId,
    options.runner_id,
    input.runnerId,
    input.runner_id,
    'meeting-platform-adapter-runner',
  ));
}

function shouldObserveAxis(options = {}) {
  return options.observe !== false
    && options.observeAxis !== false
    && options.observe_axis !== false
    && options.autoObserve !== false
    && options.auto_observe !== false;
}

function shouldValidateRawSignal(options = {}) {
  return options.validateRawSignal !== false
    && options.validate_raw_signal !== false
    && options.autoValidateRawSignal !== false
    && options.auto_validate_raw_signal !== false;
}

function shouldReadAdapterSelection(options = {}) {
  return options.readAdapterSelection !== false
    && options.read_adapter_selection !== false
    && options.autoReadAdapterSelection !== false
    && options.auto_read_adapter_selection !== false;
}

function launchPlanFromCandidate(plan) {
  return assertMeetingPlatformAdapterCandidateLaunchPlan(plan).launch_plan;
}

function runtimeTargetCapturedAtMs(target = {}) {
  const value = Number(target.mark_template?.captured_at_ms);
  return Number.isFinite(value) && value > 0 ? value : undefined;
}

function observeInputFor(input = {}) {
  if (isRuntimeTarget(input)) {
    const capturedAtMs = runtimeTargetCapturedAtMs(input);
    return compactObject({
      platform: input.platform,
      url: input.current_url,
      title: input.detected_meeting?.title,
      active: true,
      in_meeting: true,
      surface: input.selected_surface,
      candidates: input.current_url ? [compactObject({
        platform: input.platform,
        url: input.current_url,
        title: input.detected_meeting?.title,
        active: true,
        in_meeting: true,
        surface: input.selected_surface,
        source: 'meeting_platform_adapter_runner',
        captured_at_ms: capturedAtMs,
      })] : undefined,
      captured_at_ms: capturedAtMs,
    });
  }
  if (!isCandidateLaunchPlan(input)) return input;
  const evidence = input.selected_evidence ?? input.launch_input?.preflight_evidence;
  const selectedCandidate = input.selected_candidate ? compactObject({
    ...input.selected_candidate,
    in_meeting: evidence?.interaction_in_call ?? evidence?.interaction?.in_call,
    interaction: evidence?.interaction,
    semantic_signal_types: evidence?.semantic_signal_types,
    control_signal_summary: evidence?.control_signal_summary,
    control_signal_gaps: evidence?.control_signal_gaps,
    control_signal_gap_summary: evidence?.control_signal_gap_summary,
    active_speaker_candidate: evidence?.active_speaker_candidate,
    preflight_evidence: evidence,
  }) : undefined;
  return compactObject({
    platform: input.platform,
    url: input.selected_candidate?.url,
    title: input.selected_candidate?.title,
    active: input.selected_candidate?.active,
    candidates: selectedCandidate ? [selectedCandidate] : undefined,
    preflight_evidence: evidence,
    captured_at_ms: input.launch_input?.captured_at_ms ?? input.launch_input?.capturedAtMs,
  });
}

function event(action, runner, payload = {}, result) {
  const plan = payload.candidate_launch_plan ?? payload.runtime_target ?? payload.launch_plan;
  const planKind = payload.candidate_launch_plan
    ? 'candidate_launch_plan'
    : payload.runtime_target
      ? 'runtime_target'
      : 'launch_plan';
  return compactObject({
    type: 'meeting_platform_adapter_open_session_event',
    schema: MEETING_PLATFORM_ADAPTER_OPEN_SESSION_EVENT_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_RUNNER_SCHEMA_VERSION,
    runner_id: runner.id,
    action,
    platform: plan?.platform,
    selected_surface: plan?.selected_surface,
    plan_kind: planKind,
    payload,
    result,
    captured_at_ms: payload.observe_event?.captured_at_ms
      ?? payload.raw_signal_event?.captured_at_ms
      ?? payload.adapter_selection_event?.captured_at_ms
      ?? runtimeTargetCapturedAtMs(payload.runtime_target)
      ?? payload.candidate_launch_plan?.launch_plan?.mark_template?.captured_at_ms
      ?? payload.launch_plan?.mark_template?.captured_at_ms,
  });
}

function requireCurrentSession(state = {}) {
  if (state.current_session) return state.current_session;
  throw new MeetingTimelineSdkError('Adapter runner has no open session', {
    code: 'adapter_runner_session_not_open',
    required_method: 'open',
  });
}

export function createMeetingPlatformAdapterRunner(manifestOrInput = {}, clientOrOptions = {}, options = {}) {
  const defaults = defaultsFrom(clientOrOptions, options);
  const client = clientFrom(clientOrOptions, defaults);
  const installManifest = installManifestFrom(manifestOrInput, defaults);
  const runtimeManifest = runtimeManifestFrom(manifestOrInput, defaults);
  const initialLaunchPlan = isLaunchPlan(manifestOrInput)
    ? assertMeetingPlatformAdapterLaunchPlan(manifestOrInput)
    : isCandidateLaunchPlan(manifestOrInput)
      ? launchPlanFromCandidate(manifestOrInput)
    : undefined;
  const initialRuntimeTarget = isRuntimeTarget(manifestOrInput)
    ? assertMeetingPlatformAdapterRuntimeTarget(manifestOrInput)
    : undefined;
  const initialCandidateLaunchPlan = isCandidateLaunchPlan(manifestOrInput)
    ? assertMeetingPlatformAdapterCandidateLaunchPlan(manifestOrInput)
    : undefined;
  const state = {
    opened: false,
    current_launch_plan: initialLaunchPlan,
    current_runtime_target: initialRuntimeTarget,
    current_candidate_launch_plan: initialCandidateLaunchPlan,
    current_session: undefined,
    last_open_event: undefined,
    last_session_event: undefined,
  };
  const runner = {
    type: 'meeting_platform_adapter_runner',
    schema: MEETING_PLATFORM_ADAPTER_RUNNER_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_RUNNER_SCHEMA_VERSION,
    id: runnerId(manifestOrInput, defaults),
    install_manifest: installManifest,
    runtime_manifest: runtimeManifest,
    getState() {
      return compactObject({
        opened: state.opened,
        runner_id: runner.id,
        current_launch_plan: state.current_launch_plan,
        current_runtime_target: state.current_runtime_target,
        current_candidate_launch_plan: state.current_candidate_launch_plan,
        current_session: state.current_session?.getState?.(),
        last_open_event: state.last_open_event,
        last_session_event: state.last_session_event,
      });
    },
    runtimeTarget(input = {}, targetOptions = {}) {
      if (isRuntimeTarget(input)) return assertMeetingPlatformAdapterRuntimeTarget(input);
      if (state.current_runtime_target && isEmptyInput(input)) return state.current_runtime_target;
      const manifest = runtimeManifestFrom(input, targetOptions) ?? runtimeManifest;
      if (!manifest) {
        throw new MeetingTimelineSdkError('Adapter runner requires a runtime manifest to build a runtime target', {
          code: 'adapter_runner_missing_runtime_manifest',
        });
      }
      return buildMeetingPlatformAdapterRuntimeTarget(manifest, input, {
        ...defaults,
        ...targetOptions,
      });
    },
    launchPlan(input = {}, launchOptions = {}) {
      if (isLaunchPlan(input)) return assertMeetingPlatformAdapterLaunchPlan(input);
      if (isCandidateLaunchPlan(input)) return launchPlanFromCandidate(input);
      if (state.current_launch_plan && isEmptyInput(input)) return state.current_launch_plan;
      if (!installManifest && !launchOptions.installManifest && !launchOptions.install_manifest) {
        throw new MeetingTimelineSdkError('Adapter runner requires an install manifest to build a launch plan', {
          code: 'adapter_runner_missing_install_manifest',
        });
      }
      return buildMeetingPlatformAdapterLaunchPlan(installManifest, input, {
        ...defaults,
        ...launchOptions,
      });
    },
    candidateLaunchPlan(input = {}, launchOptions = {}) {
      if (isCandidateLaunchPlan(input)) return assertMeetingPlatformAdapterCandidateLaunchPlan(input);
      if (state.current_candidate_launch_plan && isEmptyInput(input)) return state.current_candidate_launch_plan;
      if (!installManifest && !launchOptions.installManifest && !launchOptions.install_manifest) {
        throw new MeetingTimelineSdkError('Adapter runner requires an install manifest to build a candidate launch plan', {
          code: 'adapter_runner_missing_install_manifest',
        });
      }
      return buildMeetingPlatformAdapterCandidateLaunchPlan(installManifest, input, {
        ...defaults,
        ...launchOptions,
      });
    },
    async open(input = {}, openOptions = {}) {
      const mergedOptions = { ...defaults, ...openOptions };
      const candidateLaunchPlan = isCandidateLaunchPlan(input)
        ? assertMeetingPlatformAdapterCandidateLaunchPlan(input)
        : state.current_candidate_launch_plan && isEmptyInput(input)
          ? state.current_candidate_launch_plan
        : undefined;
      const runtimeTarget = isRuntimeTarget(input)
        ? assertMeetingPlatformAdapterRuntimeTarget(input)
        : state.current_runtime_target && isEmptyInput(input)
          ? state.current_runtime_target
          : runtimeManifest && !isLaunchPlan(input) && !candidateLaunchPlan
            ? assertMeetingPlatformAdapterRuntimeTarget(runner.runtimeTarget(input, mergedOptions))
            : undefined;
      const launchPlan = runtimeTarget
        ? undefined
        : isLaunchPlan(input)
        ? assertMeetingPlatformAdapterLaunchPlan(input)
        : candidateLaunchPlan
          ? candidateLaunchPlan.launch_plan
        : state.current_launch_plan && isEmptyInput(input)
          ? state.current_launch_plan
        : assertMeetingPlatformAdapterLaunchPlan(runner.launchPlan(input, mergedOptions));
      const sessionPlan = runtimeTarget ?? candidateLaunchPlan ?? launchPlan;
      const session = createMeetingPlatformAdapterSession(sessionPlan, client, mergedOptions);
      const validationInput = runtimeTarget
        ? observeInputFor(runtimeTarget)
        : candidateLaunchPlan
          ? observeInputFor(candidateLaunchPlan)
          : observeInputFor(input);
      const adapterSelectionEvent = shouldReadAdapterSelection(mergedOptions)
        ? await session.readAdapterSelection(validationInput, mergedOptions)
        : undefined;
      const rawSignalEvent = shouldValidateRawSignal(mergedOptions)
        ? await session.validateRawSignal(validationInput, mergedOptions)
        : undefined;
      const observeEvent = shouldObserveAxis(mergedOptions)
        ? await session.observeAxis(validationInput, mergedOptions)
        : undefined;
      state.opened = true;
      state.current_launch_plan = launchPlan;
      state.current_runtime_target = runtimeTarget;
      state.current_candidate_launch_plan = candidateLaunchPlan;
      state.current_session = session;
      state.last_session_event = observeEvent;
      state.last_open_event = event('open_session', runner, {
        launch_plan: launchPlan,
        candidate_launch_plan: candidateLaunchPlan,
        runtime_target: runtimeTarget,
        session: {
          schema: MEETING_PLATFORM_ADAPTER_SESSION_SCHEMA,
          id: session.id,
          plan_kind: session.plan_kind,
          platform: session.platform,
          selected_surface: session.selected_surface,
          host_kind: session.host_kind,
          bridge_kind: session.bridge_kind,
          selected_evidence: session.selected_evidence,
        },
        adapter_selection_event: adapterSelectionEvent,
        raw_signal_event: rawSignalEvent,
        observe_event: observeEvent,
      }, observeEvent?.result);
      return state.last_open_event;
    },
    async openCandidate(input = {}, openOptions = {}) {
      const mergedOptions = { ...defaults, ...openOptions };
      const candidateLaunchPlan = assertMeetingPlatformAdapterCandidateLaunchPlan(
        isCandidateLaunchPlan(input) ? input : runner.candidateLaunchPlan(input, mergedOptions),
      );
      return runner.open(candidateLaunchPlan, mergedOptions);
    },
    async ensureOpen(input = {}, openOptions = {}) {
      if (state.current_session) return state.last_open_event;
      return runner.open(input, openOptions);
    },
    currentSession() {
      return requireCurrentSession(state);
    },
    async observeAxis(input = {}, observeOptions = {}) {
      const session = requireCurrentSession(state);
      state.last_session_event = await session.observeAxis(input, observeOptions);
      return state.last_session_event;
    },
    async validateRawSignal(input = {}, validationOptions = {}) {
      const session = requireCurrentSession(state);
      state.last_session_event = await session.validateRawSignal(input, validationOptions);
      return state.last_session_event;
    },
    async readAdapterSelection(input = {}, selectionOptions = {}) {
      const session = requireCurrentSession(state);
      state.last_session_event = await session.readAdapterSelection(input, selectionOptions);
      return state.last_session_event;
    },
    async insertAnnotation(input = {}, markOptions = {}) {
      const session = requireCurrentSession(state);
      state.last_session_event = await session.insertAnnotation(input, markOptions);
      return state.last_session_event;
    },
    async insertMark(input = {}, markOptions = {}) {
      return runner.insertAnnotation(input, markOptions);
    },
    async speakerTrack(input = {}, trackOptions = {}) {
      const session = requireCurrentSession(state);
      state.last_session_event = await session.speakerTrack(input, trackOptions);
      return state.last_session_event;
    },
    async participantTrack(input = {}, trackOptions = {}) {
      const session = requireCurrentSession(state);
      state.last_session_event = await session.participantTrack(input, trackOptions);
      return state.last_session_event;
    },
    async providerReconcile(input = {}, reconcileOptions = {}) {
      const session = requireCurrentSession(state);
      state.last_session_event = await session.providerReconcile(input, reconcileOptions);
      return state.last_session_event;
    },
    reset() {
      state.opened = false;
      state.current_launch_plan = initialLaunchPlan;
      state.current_runtime_target = initialRuntimeTarget;
      state.current_candidate_launch_plan = initialCandidateLaunchPlan;
      state.current_session = undefined;
      state.last_open_event = undefined;
      state.last_session_event = undefined;
      return runner.getState();
    },
  };
  return runner;
}

export async function openMeetingPlatformAdapterSession(
  manifestOrInput = {},
  clientOrOptions = {},
  input = {},
  options = {},
) {
  const runner = createMeetingPlatformAdapterRunner(manifestOrInput, clientOrOptions, options);
  return runner.open(input, options);
}

export async function openMeetingPlatformAdapterRuntimeSession(
  runtimeManifestOrTarget = {},
  clientOrOptions = {},
  input = {},
  options = {},
) {
  const runner = createMeetingPlatformAdapterRunner(runtimeManifestOrTarget, clientOrOptions, options);
  return runner.open(input, options);
}

export function buildMeetingPlatformAdapterRunnerHandoff(manifestOrInput = {}, options = {}) {
  const installManifest = installManifestFrom(manifestOrInput, options);
  const runtimeManifest = runtimeManifestFrom(manifestOrInput, options);
  const runtimeTarget = isRuntimeTarget(manifestOrInput) ? manifestOrInput : undefined;
  return {
    type: 'meeting_platform_adapter_runner_handoff',
    schema: 'meeting_platform_adapter_runner_handoff',
    schema_version: MEETING_PLATFORM_ADAPTER_RUNNER_SCHEMA_VERSION,
    runner_factory: 'createMeetingPlatformAdapterRunner',
    convenience_method: runtimeManifest || runtimeTarget
      ? 'openMeetingPlatformAdapterRuntimeSession'
      : 'openMeetingPlatformAdapterSession',
    install_manifest_schema: installManifest?.schema,
    runtime_manifest_schema: runtimeManifest?.schema,
    runtime_target_schema: runtimeTarget?.schema,
    required_client_methods: ['observePlatformCandidates', 'insertAnnotation'],
    optional_client_methods: ['platformAdapterSelection', 'adapterSelection', 'platformRawSignalBatch', 'rawSignalBatch', 'speakerTrack', 'participantTrack', 'ingestProvider'],
    runtime_sequence: [
      runtimeManifest || runtimeTarget ? 'build_runtime_target_from_current_url_or_platform' : 'build_launch_plan_from_current_url_or_platform',
      runtimeManifest || runtimeTarget ? 'create_session_from_runtime_target' : 'create_session_from_launch_plan',
      'read_adapter_selection_before_runtime_wiring',
      'validate_raw_signal_before_adapter_preflight',
      'observe_axis_before_first_mark',
      'insert_realtime_annotations_with_captured_at_ms',
    ],
    next_actions: [
      runtimeManifest || runtimeTarget ? 'create_runner_with_adapter_runtime_manifest' : 'create_runner_with_adapter_install_manifest',
      'call_runner_open_with_current_meeting_url',
      'reuse_runner_insertAnnotation_for_realtime_marks',
    ],
  };
}
