import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  MEETING_PLATFORM_ADAPTER_LAUNCH_PLAN_SCHEMA,
  assertMeetingPlatformAdapterLaunchPlan,
  buildMeetingPlatformAdapterLaunchPlan,
} from './platform-adapter-launch-plan.mjs';
import {
  MEETING_PLATFORM_ADAPTER_SESSION_SCHEMA,
  createMeetingPlatformAdapterSession,
} from './platform-adapter-session.mjs';

export const MEETING_PLATFORM_ADAPTER_RUNNER_SCHEMA = 'meeting_platform_adapter_runner';
export const MEETING_PLATFORM_ADAPTER_OPEN_SESSION_EVENT_SCHEMA = 'meeting_platform_adapter_open_session_event';
export const MEETING_PLATFORM_ADAPTER_RUNNER_SCHEMA_VERSION = 1;

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function isLaunchPlan(value) {
  return value?.schema === MEETING_PLATFORM_ADAPTER_LAUNCH_PLAN_SCHEMA;
}

function isInstallManifest(value) {
  return value?.schema === 'meeting_platform_adapter_install_manifest';
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

function event(action, runner, payload = {}, result) {
  return compactObject({
    type: 'meeting_platform_adapter_open_session_event',
    schema: MEETING_PLATFORM_ADAPTER_OPEN_SESSION_EVENT_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_RUNNER_SCHEMA_VERSION,
    runner_id: runner.id,
    action,
    platform: payload.launch_plan?.platform,
    selected_surface: payload.launch_plan?.selected_surface,
    payload,
    result,
    captured_at_ms: payload.observe_event?.captured_at_ms ?? payload.launch_plan?.mark_template?.captured_at_ms,
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
  const initialLaunchPlan = isLaunchPlan(manifestOrInput)
    ? assertMeetingPlatformAdapterLaunchPlan(manifestOrInput)
    : undefined;
  const state = {
    opened: false,
    current_launch_plan: initialLaunchPlan,
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
    getState() {
      return compactObject({
        opened: state.opened,
        runner_id: runner.id,
        current_launch_plan: state.current_launch_plan,
        current_session: state.current_session?.getState?.(),
        last_open_event: state.last_open_event,
        last_session_event: state.last_session_event,
      });
    },
    launchPlan(input = {}, launchOptions = {}) {
      if (isLaunchPlan(input)) return assertMeetingPlatformAdapterLaunchPlan(input);
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
    async open(input = {}, openOptions = {}) {
      const mergedOptions = { ...defaults, ...openOptions };
      const launchPlan = isLaunchPlan(input)
        ? assertMeetingPlatformAdapterLaunchPlan(input)
        : state.current_launch_plan && isEmptyInput(input)
          ? state.current_launch_plan
        : assertMeetingPlatformAdapterLaunchPlan(runner.launchPlan(input, mergedOptions));
      const session = createMeetingPlatformAdapterSession(launchPlan, client, mergedOptions);
      const observeEvent = shouldObserveAxis(mergedOptions)
        ? await session.observeAxis(input, mergedOptions)
        : undefined;
      state.opened = true;
      state.current_launch_plan = launchPlan;
      state.current_session = session;
      state.last_session_event = observeEvent;
      state.last_open_event = event('open_session', runner, {
        launch_plan: launchPlan,
        session: {
          schema: MEETING_PLATFORM_ADAPTER_SESSION_SCHEMA,
          id: session.id,
          platform: session.platform,
          selected_surface: session.selected_surface,
        },
        observe_event: observeEvent,
      }, observeEvent?.result);
      return state.last_open_event;
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

export function buildMeetingPlatformAdapterRunnerHandoff(manifestOrInput = {}, options = {}) {
  const installManifest = installManifestFrom(manifestOrInput, options);
  return {
    type: 'meeting_platform_adapter_runner_handoff',
    schema: 'meeting_platform_adapter_runner_handoff',
    schema_version: MEETING_PLATFORM_ADAPTER_RUNNER_SCHEMA_VERSION,
    runner_factory: 'createMeetingPlatformAdapterRunner',
    convenience_method: 'openMeetingPlatformAdapterSession',
    install_manifest_schema: installManifest?.schema,
    required_client_methods: ['observePlatformCandidates', 'insertAnnotation'],
    optional_client_methods: ['speakerTrack', 'participantTrack', 'ingestProvider'],
    runtime_sequence: [
      'build_launch_plan_from_current_url_or_platform',
      'create_session_from_launch_plan',
      'observe_axis_before_first_mark',
      'insert_realtime_annotations_with_captured_at_ms',
    ],
    next_actions: [
      'create_runner_with_adapter_install_manifest',
      'call_runner_open_with_current_meeting_url',
      'reuse_runner_insertAnnotation_for_realtime_marks',
    ],
  };
}
