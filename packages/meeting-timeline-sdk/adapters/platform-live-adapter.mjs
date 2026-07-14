import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import { createMeetingSourceAggregator } from './meeting-source.mjs';
import { verifyMeetingPlatformEvidencePackage } from './platform-evidence-package.mjs';
import { createMeetingPlatformEvidenceSession } from './platform-evidence-session.mjs';
import {
  MEETING_PLATFORM_KEYS,
  buildPlatformIntegrationPlan,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';
import { buildMeetingPlatformRealtimeAnnotation } from './platform-realtime-annotation.mjs';
import { buildMeetingPlatformAdaptationRunbook } from './platform-rollout.mjs';
import { buildMeetingPlatformAdaptationStrategy } from './platform-strategy.mjs';
import { MEETING_APP_EXTENSION_MESSAGE_TYPES } from './meeting-app-extension.mjs';
import {
  MEETING_PLATFORM_RUNTIME_EVENT_ENDPOINT,
  buildMeetingPlatformRuntimeEventPlan,
} from './platform-runtime-event.mjs';

export const MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA = 'meeting_platform_live_adapter';
export const MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION = 1;
export const MEETING_PLATFORM_LIVE_ADAPTER_PLAN_SCHEMA = 'meeting_platform_live_adapter_plan';
export const MEETING_PLATFORM_LIVE_ADAPTER_MATRIX_SCHEMA = 'meeting_platform_live_adapter_matrix';
export const MEETING_PLATFORM_LIVE_ADAPTER_READINESS_SCHEMA = 'meeting_platform_live_adapter_readiness';
export const MEETING_PLATFORM_LIVE_ADAPTER_READINESS_MATRIX_SCHEMA = 'meeting_platform_live_adapter_readiness_matrix';
export const MEETING_PLATFORM_LIVE_ADAPTER_HANDOFF_SCHEMA = 'meeting_platform_live_adapter_handoff';
export const MEETING_PLATFORM_LIVE_ADAPTER_HANDOFF_BUNDLE_SCHEMA = 'meeting_platform_live_adapter_handoff_bundle';
export const MEETING_PLATFORM_CANDIDATE_OBSERVATION_ENDPOINT = '/api/meeting-platform/observe-candidates';
export const MEETING_PLATFORM_LIVE_ADAPTER_REQUIRED_METHODS = Object.freeze([
  'observeMeetingApp',
  'ingestProvider',
  'insertAnnotation',
  'summary',
  'exportPackage',
  'verify',
]);

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

function uniqueList(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function matchesPlatform(value, platform) {
  if (!value) return true;
  try {
    return normalizeMeetingPlatform(value) === platform;
  } catch {
    return false;
  }
}

function isTimelineClient(value) {
  return Boolean(value)
    && typeof value === 'object'
    && typeof value.startMeeting === 'function'
    && typeof value.endMeeting === 'function'
    && typeof value.insertMark === 'function';
}

function liveSessionOptions(options = {}) {
  return compactObject({
    requireMeetingEnd: firstNonEmpty(options.requireMeetingEnd, options.require_meeting_end, false),
    ...options,
  });
}

function providerPayload(input, payload) {
  if (input && typeof input === 'object' && !Array.isArray(input)) {
    return firstNonEmpty(input.payload, input.body, input.event, input.raw, payload);
  }
  return payload;
}

function boolFlag(options = {}, keys = [], fallback) {
  const value = firstNonEmpty(...keys.map((key) => options[key]));
  return value == null ? fallback : value !== false && value !== 'false';
}

function directMeetingInput(input = {}, options = {}) {
  return firstNonEmpty(
    input.current_meeting,
    input.currentMeeting,
    input.current_axis,
    input.currentAxis,
    input.meeting,
    options.current_meeting,
    options.currentMeeting,
    options.current_axis,
    options.currentAxis,
    options.meeting,
  );
}

function meetingFromActiveState(platform, source) {
  const state = typeof source.getState === 'function' ? source.getState() : {};
  const active = (state.reconciler?.active_meetings ?? [])
    .filter((meeting) => matchesPlatform(meeting.platform, platform))
    .at(-1);
  if (!active) return undefined;
  return compactObject({
    platform,
    meeting_id: active.meeting_id,
    external_meeting_id: active.external_meeting_id,
    meeting_url: active.meeting_url,
    start_time_ms: active.occurred_at_ms,
    source: active.source,
    source_event_id: active.source_event_id,
  });
}

function stripRoutingFields(input = {}) {
  const {
    annotation,
    mark,
    item,
    payload,
    current_meeting: currentMeetingSnake,
    currentMeeting,
    current_axis: currentAxisSnake,
    currentAxis,
    meeting,
    local_observer: localObserverSnake,
    localObserver,
    local_observers: localObserversSnake,
    localObservers,
    provider_signal: providerSignalSnake,
    providerSignal,
    provider_signals: providerSignalsSnake,
    providerSignals,
    provider_event: providerEventSnake,
    providerEvent,
    provider_events: providerEventsSnake,
    providerEvents,
    signal,
    signals,
    samples,
    clock_samples: clockSamplesSnake,
    clockSamples,
    time_sync_samples: timeSyncSamplesSnake,
    timeSyncSamples,
    clock_sync: clockSyncSnake,
    clockSync,
    time_sync: timeSyncSnake,
    timeSync,
    ...annotationFields
  } = input;
  void annotation;
  void mark;
  void item;
  void payload;
  void currentMeetingSnake;
  void currentMeeting;
  void currentAxisSnake;
  void currentAxis;
  void meeting;
  void localObserverSnake;
  void localObserver;
  void localObserversSnake;
  void localObservers;
  void providerSignalSnake;
  void providerSignal;
  void providerSignalsSnake;
  void providerSignals;
  void providerEventSnake;
  void providerEvent;
  void providerEventsSnake;
  void providerEvents;
  void signal;
  void signals;
  void samples;
  void clockSamplesSnake;
  void clockSamples;
  void timeSyncSamplesSnake;
  void timeSyncSamples;
  void clockSyncSnake;
  void clockSync;
  void timeSyncSnake;
  void timeSync;
  return annotationFields;
}

function annotationPayload(input = {}, platform) {
  const explicit = firstNonEmpty(input.annotation, input.mark, input.item, input.payload?.annotation);
  if (isPlainObject(explicit)) {
    return {
      platform,
      ...explicit,
    };
  }
  return {
    platform,
    ...stripRoutingFields(input),
  };
}

function enrichAnnotationWithMeeting(annotation = {}, meeting = {}) {
  if (!meeting || Object.keys(meeting).length === 0) return annotation;
  return compactObject({
    ...annotation,
    meeting_id: firstNonEmpty(annotation.meeting_id, annotation.meetingId, annotation.session_id, annotation.sessionId, meeting.meeting_id),
    external_meeting_id: firstNonEmpty(annotation.external_meeting_id, annotation.externalMeetingId, meeting.external_meeting_id),
    meeting_url: firstNonEmpty(annotation.meeting_url, annotation.meetingUrl, annotation.url, meeting.meeting_url),
    meeting_title: firstNonEmpty(annotation.meeting_title, annotation.meetingTitle, meeting.title),
  });
}

function realtimeAnnotationOptions(baseOptions = {}, markOptions = {}) {
  const configured = {
    ...(baseOptions.realtimeAnnotation ?? {}),
    ...(baseOptions.realtimeAnnotationOptions ?? {}),
    ...(baseOptions.realtime_annotation ?? {}),
    ...(baseOptions.realtime_annotation_options ?? {}),
    ...(markOptions.realtimeAnnotation ?? {}),
    ...(markOptions.realtimeAnnotationOptions ?? {}),
    ...(markOptions.realtime_annotation ?? {}),
    ...(markOptions.realtime_annotation_options ?? {}),
  };
  const requireClockSync = boolFlag({
    ...baseOptions,
    ...markOptions,
    ...configured,
  }, ['requireClockSync', 'require_clock_sync'], false);
  return compactObject({
    ...configured,
    allowUnsyncedCapturedAtMs: firstNonEmpty(
      configured.allowUnsyncedCapturedAtMs,
      configured.allow_unsynced_captured_at_ms,
      requireClockSync ? false : true,
    ),
  });
}

function realtimeAnnotationInput(platform, source, input = {}, markOptions = {}) {
  const currentMeeting = directMeetingInput(input, markOptions) ?? meetingFromActiveState(platform, source);
  const annotation = enrichAnnotationWithMeeting(annotationPayload(input, platform), currentMeeting);
  const extra = firstNonEmpty(
    markOptions.realtimeAnnotationInput,
    markOptions.realtime_annotation_input,
    input.realtimeAnnotationInput,
    input.realtime_annotation_input,
    {},
  );
  return compactObject({
    ...markOptions,
    ...input,
    ...extra,
    platform,
    annotation,
    current_meeting: currentMeeting,
    clock_sync: firstNonEmpty(input.clock_sync, input.clockSync, markOptions.clock_sync, markOptions.clockSync, extra.clock_sync, extra.clockSync),
    samples: firstNonEmpty(input.samples, markOptions.samples, extra.samples),
    clock_samples: firstNonEmpty(input.clock_samples, input.clockSamples, markOptions.clock_samples, markOptions.clockSamples, extra.clock_samples, extra.clockSamples),
    time_sync_samples: firstNonEmpty(input.time_sync_samples, input.timeSyncSamples, markOptions.time_sync_samples, markOptions.timeSyncSamples, extra.time_sync_samples, extra.timeSyncSamples),
  });
}

function meetingAppRecordOptions(input = {}, options = {}) {
  return {
    ...options,
    phase: firstNonEmpty(
      options.phase,
      options.state,
      input.phase,
      input.state,
      input.fixture_state,
      input.snapshot?.phase,
      input.snapshot?.state,
      input.snapshot?.fixture_state,
    ),
    capturedAtMs: firstNonEmpty(
      options.capturedAtMs,
      options.captured_at_ms,
      options.observedAtMs,
      options.observed_at_ms,
      input.capturedAtMs,
      input.captured_at_ms,
      input.observedAtMs,
      input.observed_at_ms,
      input.snapshot?.observedAtMs,
      input.snapshot?.observed_at_ms,
    ),
  };
}

async function executeRealtimeAnnotationPipeline(source, platform, input = {}, markOptions = {}, baseOptions = {}) {
  const usePipeline = boolFlag({
    ...baseOptions,
    ...markOptions,
  }, ['useRealtimeAnnotationPipeline', 'use_realtime_annotation_pipeline'], true);
  if (!usePipeline) {
    const direct = await source.insertAnnotation({
      platform,
      ...input,
    }, markOptions);
    return compactObject({
      mode: 'direct_insert',
      ok: true,
      insert_result: direct,
    });
  }
  const decision = buildMeetingPlatformRealtimeAnnotation(
    platform,
    realtimeAnnotationInput(platform, source, input, markOptions),
    realtimeAnnotationOptions(baseOptions, markOptions),
  );
  let startResult;
  let insertResult;
  if (decision.actions.includes('start_meeting_session')) {
    startResult = await source.startMeeting(decision.start_payload);
  }
  if (decision.actions.includes('insert_mark')) {
    insertResult = await source.insertAnnotation(decision.insert_payload, markOptions);
  }
  return compactObject({
    mode: 'realtime_annotation_pipeline',
    ok: decision.accepted_for_realtime,
    status: decision.status,
    actions: decision.actions,
    pipeline: decision,
    start_result: startResult,
    insert_result: insertResult,
    pending_payload: decision.pending_payload,
  });
}

function resultWithEvidence(result = {}, evidenceSession, action, options = {}) {
  return compactObject({
    ...result,
    action,
    live_evidence: evidenceSession.summary(options.evidenceOptions ?? options.evidence_options ?? {}),
    evidence_state: evidenceSession.getState(),
  });
}

function operationWithEvidence(action, result, evidenceSession, options = {}) {
  return compactObject({
    action,
    result,
    live_evidence: evidenceSession.summary(options.evidenceOptions ?? options.evidence_options ?? {}),
    evidence_state: evidenceSession.getState(),
  });
}

function selectedPlatforms(options = {}) {
  return asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_PLATFORM_KEYS));
}

function liveAdapterSdkContract() {
  return {
    package: '@ai-annotation/meeting-timeline-sdk',
    module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-live-adapter',
    kit_module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit',
    factory: 'createMeetingPlatformLiveAdapter',
    suite_factory: 'createMeetingPlatformLiveAdapterSuite',
    readiness_methods: [
      'buildMeetingPlatformLiveAdapterReadiness',
      'buildMeetingPlatformLiveAdapterReadinessMatrix',
      'assertMeetingPlatformLiveAdapterReadiness',
      'assertMeetingPlatformLiveAdapterReadinessMatrix',
    ],
    kit_methods: [
      'platformLiveAdapter',
      'platformLiveAdapterSuite',
      'platformLiveAdapterHandoff',
      'platformLiveAdapterHandoffBundle',
    ],
  };
}

function liveAdapterCommands(platform) {
  const key = normalizeMeetingPlatform(platform);
  return {
    build_capture_runtime: 'npm run meeting-app:extension:build',
    validate_dom_evidence: `npm run meeting-app:evidence-gate -- --input=data/meeting-app-evidence/${key}.json --report-file=data/meeting-app-live-gate-report.json`,
    validate_dom_matrix: 'npm run meeting-app:evidence-matrix',
    validate_provider_events: 'buildPlatformLaunchGate(platform, { records: providerRecords, env, baseUrl })',
    validate_rollout: 'npm run meeting-platform:rollout-matrix',
    verify_evidence_package: `npm run meeting-platform:evidence-package -- --input=data/meeting-platform-evidence-packages/${key}.json`,
    validate_live_readiness: 'npm run meeting-platform:live-readiness',
    package_smoke: 'npm run sdk:package-smoke',
  };
}

function liveAdapterEvidencePaths(platform) {
  const key = normalizeMeetingPlatform(platform);
  return {
    meeting_app_evidence: `data/meeting-app-evidence/${key}.json`,
    provider_evidence: `data/provider-evidence/${key}.json`,
    evidence_package: `data/meeting-platform-evidence-packages/${key}.json`,
    live_readiness_report: 'data/meeting-platform-live-readiness-report.json',
  };
}

function hostContract(plan = {}, candidateObservation = {}) {
  return {
    annotation_timestamp_field: plan.live_adapter?.timestamp_field ?? 'captured_at_ms',
    realtime_source_order: plan.source_priority ?? [],
    meeting_axis_primary_source: plan.realtime_axis?.primary_source,
    provider_events_block_realtime: plan.realtime_axis?.provider_events_block_realtime === true,
    transcript_blocks_realtime: plan.realtime_axis?.transcript_blocks_realtime === true,
    must_keep_annotations_per_meeting: true,
    must_use_absolute_capture_time_ms: true,
    can_insert_before_provider_event: plan.realtime_axis?.provider_events_block_realtime === false,
    candidate_observation_required_for_host_axis_binding: true,
    candidate_observation_message_type: candidateObservation.message_type,
    candidate_observation_runtime_action: candidateObservation.runtime_event_action,
    candidate_observation_runtime_event_endpoint: candidateObservation.runtime_event_endpoint,
    candidate_observation_endpoint: candidateObservation.endpoint,
    candidate_observation_required_permission: candidateObservation.required_permission,
    candidate_observation: compactObject({
      ready: candidateObservation.ready,
      message_type: candidateObservation.message_type,
      runtime_event_action: candidateObservation.runtime_event_action,
      runtime_event_endpoint: candidateObservation.runtime_event_endpoint,
      endpoint: candidateObservation.endpoint,
      required_permission: candidateObservation.required_permission,
      runtime_event_client_method: candidateObservation.runtime_event_client_method,
    }),
  };
}

function candidateObservationOverride(options = {}) {
  return firstNonEmpty(
    options.candidateObservation,
    options.candidate_observation,
    options.candidateObserver,
    options.candidate_observer,
    {},
  ) ?? {};
}

function buildCandidateObservationContract(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const runtimeEventPlan = buildMeetingPlatformRuntimeEventPlan(key, options);
  const override = candidateObservationOverride(options);
  const action = firstNonEmpty(
    override.runtime_event_action,
    override.runtimeEventAction,
    override.action,
    'observe_platform_candidates',
  );
  const messageType = firstNonEmpty(
    override.message_type,
    override.messageType,
    MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
  );
  const requiredPermission = firstNonEmpty(
    override.required_permission,
    override.requiredPermission,
    override.permission,
    'tabs',
  );
  const endpoint = firstNonEmpty(
    override.endpoint,
    MEETING_PLATFORM_CANDIDATE_OBSERVATION_ENDPOINT,
  );
  const runtimeEventEndpoint = firstNonEmpty(
    override.runtime_event_endpoint,
    override.runtimeEventEndpoint,
    runtimeEventPlan.endpoint,
    MEETING_PLATFORM_RUNTIME_EVENT_ENDPOINT,
  );
  const clientMethod = firstNonEmpty(
    override.runtime_event_client_method,
    override.runtimeEventClientMethod,
    override.client_method,
    override.clientMethod,
    'observePlatformCandidates',
  );
  const actionSupported = runtimeEventPlan.supported_actions?.includes?.('observe_platform_candidates')
    || runtimeEventPlan.actions?.some?.((row) => row.action === 'observe_platform_candidates');
  const issues = [
    action !== 'observe_platform_candidates' ? 'invalid_runtime_event_action' : undefined,
    messageType !== MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates ? 'invalid_message_type' : undefined,
    requiredPermission !== 'tabs' ? 'missing_tabs_permission' : undefined,
    endpoint !== MEETING_PLATFORM_CANDIDATE_OBSERVATION_ENDPOINT ? 'invalid_candidate_observation_endpoint' : undefined,
    runtimeEventEndpoint !== runtimeEventPlan.endpoint ? 'invalid_runtime_event_endpoint' : undefined,
    clientMethod !== 'observePlatformCandidates' ? 'invalid_runtime_event_client_method' : undefined,
    actionSupported !== true ? 'runtime_event_plan_missing_candidate_action' : undefined,
  ].filter(Boolean);
  return compactObject({
    type: 'meeting_platform_candidate_observation_contract',
    platform: key,
    ready: issues.length === 0,
    message_type: messageType,
    required_permission: requiredPermission,
    runtime_event_action: action,
    runtime_event_endpoint: runtimeEventEndpoint,
    endpoint,
    producer: firstNonEmpty(override.producer, 'browser_extension_background_or_native_host'),
    runtime_event_client_method: clientMethod,
    required_snapshot_fields: ['url', 'title', 'active', 'captured_at_ms'],
    action_supported: actionSupported,
    runtime_event_plan_schema: runtimeEventPlan.schema,
    issues,
  });
}

function selectedEvidencePackage(platform, options = {}) {
  const input = firstNonEmpty(
    options.evidencePackage,
    options.evidence_package,
    options.handoffPackage,
    options.handoff_package,
    options.package,
  );
  if (!input) return undefined;
  if (Array.isArray(input)) {
    return input.find((item) => matchesPlatform(item?.platform ?? item?.rollout_plan?.platform, platform));
  }
  if (input.schema || input.rollout_plan || input.provider_records || input.meeting_app_record_set) {
    return matchesPlatform(input.platform ?? input.rollout_plan?.platform, platform) ? input : undefined;
  }
  if (typeof input === 'object') {
    for (const [key, value] of Object.entries(input)) {
      if (matchesPlatform(key, platform)) return value;
    }
  }
  return undefined;
}

function readinessTarget(options = {}) {
  const target = String(firstNonEmpty(
    options.target,
    options.readinessTarget,
    options.readiness_target,
    options.requireProductionReady === true || options.require_production_ready === true ? 'production' : 'pilot',
  ));
  return target === 'production' ? 'production' : 'pilot';
}

function check(code, passed, severity, message, details = {}) {
  return compactObject({
    code,
    passed: Boolean(passed),
    severity: passed ? 'info' : severity,
    message,
    ...details,
  });
}

function adapterMethodCheck(adapter, requiredMethods = MEETING_PLATFORM_LIVE_ADAPTER_REQUIRED_METHODS) {
  if (!adapter) return undefined;
  const missing = requiredMethods.filter((method) => typeof adapter[method] !== 'function');
  return check(
    'adapter_methods_available',
    missing.length === 0,
    'error',
    'Live adapter instance must expose the realtime and evidence handoff methods.',
    { required_methods: requiredMethods, missing_methods: missing },
  );
}

function readinessStatus(blocking = [], warnings = []) {
  if (blocking.length > 0) return 'blocked';
  if (warnings.length > 0) return 'warning';
  return 'ready';
}

function planWithRollout(plan = {}, rollout = {}) {
  if (!rollout || Object.keys(rollout).length === 0) return plan;
  return {
    ...plan,
    rollout_status: rollout.status ?? plan.rollout_status,
    production_ready: Boolean(rollout.production_ready),
    ready_for_realtime_annotations: Boolean(rollout.ready_for_realtime_annotations),
    pilot_ready: Boolean(rollout.ready_for_realtime_annotations),
    recommended_mode: rollout.recommended_mode ?? plan.recommended_mode,
    next_actions: uniqueList([
      ...(rollout.next_actions ?? []),
      ...(plan.next_actions ?? []),
    ]),
  };
}

export function buildMeetingPlatformLiveAdapterPlan(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const integration = buildPlatformIntegrationPlan(key, options);
  const strategy = buildMeetingPlatformAdaptationStrategy(key, options);
  return compactObject({
    schema: MEETING_PLATFORM_LIVE_ADAPTER_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION,
    type: 'meeting_platform_live_adapter_plan',
    platform: key,
    display_name: strategy.display_name ?? integration.display_name,
    rollout_status: strategy.rollout_status,
    production_ready: strategy.production_ready,
    ready_for_realtime_annotations: strategy.ready_for_realtime_annotations,
    pilot_ready: strategy.ready_for_realtime_annotations === true,
    recommended_mode: strategy.recommendation,
    source_priority: strategy.source_priority,
    live_adapter: {
      module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-live-adapter',
      factory: 'createMeetingPlatformLiveAdapter',
      kit_method: 'platformLiveAdapter',
      realtime_methods: ['observeMeetingApp', 'ingestProvider', 'insertAnnotation'],
      evidence_methods: ['summary', 'correlation', 'exportPackage', 'verify'],
      timestamp_field: strategy.realtime_axis?.annotation_time_field ?? 'captured_at_ms',
    },
    realtime_axis: strategy.realtime_axis,
    local_observer: strategy.local_observer,
    provider_events: strategy.provider_events,
    speaker_activity: strategy.speaker_activity,
    post_meeting_transcript: strategy.post_meeting_transcript,
    handoff: {
      evidence_session_module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-session',
      evidence_package_module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-package',
      package_schema: 'meeting_platform_evidence_package',
      verifier: 'verifyMeetingPlatformEvidencePackage',
    },
    endpoints: {
      provider_events: integration.provider_events?.endpoint,
      transcript_import: integration.post_meeting_transcript?.import_endpoint,
      status: integration.provider_events?.status_endpoint,
    },
    next_actions: uniqueList(strategy.next_actions ?? []),
  });
}

export function buildMeetingPlatformLiveAdapterMatrix(options = {}) {
  const plans = selectedPlatforms(options).map((platform) => buildMeetingPlatformLiveAdapterPlan(platform, options));
  return {
    type: 'meeting_platform_live_adapter_matrix',
    schema: MEETING_PLATFORM_LIVE_ADAPTER_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION,
    platform_count: plans.length,
    production_ready_count: plans.filter((plan) => plan.production_ready).length,
    pilot_ready_count: plans.filter((plan) => plan.pilot_ready).length,
    local_first_count: plans.filter((plan) => plan.realtime_axis?.primary_source === 'local_observer').length,
    non_blocking_provider_count: plans.filter((plan) => plan.realtime_axis?.provider_events_block_realtime === false).length,
    platforms: plans.map((plan) => plan.platform),
    rows: plans.map((plan) => ({
      platform: plan.platform,
      display_name: plan.display_name,
      rollout_status: plan.rollout_status,
      recommended_mode: plan.recommended_mode,
      pilot_ready: plan.pilot_ready,
      production_ready: plan.production_ready,
      primary_axis_source: plan.realtime_axis?.primary_source,
      provider_blocks_realtime: plan.realtime_axis?.provider_events_block_realtime,
      transcript_blocks_realtime: plan.realtime_axis?.transcript_blocks_realtime,
      next_actions: plan.next_actions,
    })),
    plans,
  };
}

export function buildMeetingPlatformLiveAdapterReadiness(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const target = readinessTarget(options);
  const evidencePackage = selectedEvidencePackage(key, options);
  const verification = evidencePackage
    ? verifyMeetingPlatformEvidencePackage(evidencePackage, {
      ...options,
      requireProductionReady: target === 'production',
      require_production_ready: target === 'production',
      includePackage: true,
    })
    : undefined;
  const strategy = buildMeetingPlatformAdaptationStrategy(key, {
    ...options,
    includeRolloutPlan: true,
  });
  const rollout = verification?.verified_package?.rollout_plan ?? strategy.rollout_plan ?? {};
  const plan = planWithRollout(buildMeetingPlatformLiveAdapterPlan(key, options), rollout);
  const candidateObservation = buildCandidateObservationContract(key, options);
  const requiredMethods = asArray(firstNonEmpty(options.requiredMethods, options.required_methods, MEETING_PLATFORM_LIVE_ADAPTER_REQUIRED_METHODS));
  const checks = [
    check(
      'adapter_contract_declared',
      requiredMethods.every((method) => [
        ...(plan.live_adapter?.realtime_methods ?? []),
        ...(plan.live_adapter?.evidence_methods ?? []),
      ].includes(method)),
      'error',
      'Live adapter plan must declare the methods needed by a host project.',
      { required_methods: requiredMethods },
    ),
    adapterMethodCheck(firstNonEmpty(options.adapter, options.liveAdapter, options.live_adapter), requiredMethods),
    check(
      'timestamp_field_contract',
      plan.live_adapter?.timestamp_field === 'captured_at_ms',
      'error',
      'Realtime marks must use captured_at_ms so annotations land on the current meeting axis.',
      { timestamp_field: plan.live_adapter?.timestamp_field },
    ),
    check(
      'provider_events_non_blocking',
      plan.realtime_axis?.provider_events_block_realtime === false,
      'error',
      'Provider events must not block realtime annotation insertion.',
    ),
    check(
      'candidate_observation_contract_ready',
      candidateObservation.ready === true,
      'error',
      'Host handoff must expose candidate observation so active meeting windows can create and maintain current axes.',
      {
        message_type: candidateObservation.message_type,
        required_permission: candidateObservation.required_permission,
        runtime_event_action: candidateObservation.runtime_event_action,
        endpoint: candidateObservation.endpoint,
        runtime_event_endpoint: candidateObservation.runtime_event_endpoint,
        issues: candidateObservation.issues,
      },
    ),
    check(
      'post_meeting_transcript_non_blocking',
      plan.realtime_axis?.transcript_blocks_realtime === false,
      'error',
      'Post-meeting transcript import must not block realtime annotation insertion.',
    ),
    check(
      'pilot_realtime_axis_ready',
      plan.ready_for_realtime_annotations === true,
      'error',
      'At least one low-latency axis source must be verified before realtime annotations can be enabled.',
      { rollout_status: plan.rollout_status },
    ),
    check(
      'production_evidence_ready',
      plan.production_ready === true,
      target === 'production' ? 'error' : 'warning',
      'Production rollout requires both local observer evidence and provider reconcile evidence.',
      { rollout_status: plan.rollout_status },
    ),
    evidencePackage ? check(
      'evidence_package_verified',
      verification?.passed === true,
      'error',
      'Evidence package must verify against the current SDK gates.',
      {
        requirement: verification?.requirement,
        correlation_passed: verification?.correlation_passed,
        embedded_plan_matches: verification?.embedded_plan_matches,
      },
    ) : check(
      'evidence_package_available',
      false,
      target === 'production' ? 'error' : 'warning',
      'No evidence package was provided; SDK wiring can be inspected but handoff cannot be audited.',
    ),
  ].filter(Boolean);
  const blockingChecks = checks.filter((item) => item.passed !== true && item.severity === 'error');
  const warnings = checks.filter((item) => item.passed !== true && item.severity !== 'error');
  const status = readinessStatus(blockingChecks, warnings);
  return compactObject({
    type: 'meeting_platform_live_adapter_readiness',
    schema: MEETING_PLATFORM_LIVE_ADAPTER_READINESS_SCHEMA,
    schema_version: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION,
    platform: key,
    display_name: plan.display_name,
    target,
    status,
    passed: blockingChecks.length === 0,
    ready_for_realtime_annotations: plan.ready_for_realtime_annotations,
    production_ready: plan.production_ready,
    candidate_observation_ready: candidateObservation.ready === true,
    candidate_observation: candidateObservation,
    rollout_status: plan.rollout_status,
    recommended_mode: plan.recommended_mode,
    blocking_count: blockingChecks.length,
    warning_count: warnings.length,
    checks,
    blocking_checks: blockingChecks,
    warnings,
    verification,
    plan,
    next_actions: uniqueList([
      ...blockingChecks.map((item) => item.code),
      ...warnings.map((item) => item.code),
      ...(verification?.next_actions ?? []),
      ...(plan.next_actions ?? []),
    ]),
  });
}

export function buildMeetingPlatformLiveAdapterReadinessMatrix(options = {}) {
  const rows = selectedPlatforms(options).map((platform) => buildMeetingPlatformLiveAdapterReadiness(platform, options));
  return {
    type: 'meeting_platform_live_adapter_readiness_matrix',
    schema: MEETING_PLATFORM_LIVE_ADAPTER_READINESS_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION,
    platform_count: rows.length,
    passed_count: rows.filter((row) => row.passed).length,
    ready_count: rows.filter((row) => row.status === 'ready').length,
    warning_count: rows.filter((row) => row.status === 'warning').length,
    blocked_count: rows.filter((row) => row.status === 'blocked').length,
    realtime_ready_count: rows.filter((row) => row.ready_for_realtime_annotations).length,
    production_ready_count: rows.filter((row) => row.production_ready).length,
    candidate_observer_count: rows.filter((row) => row.candidate_observation_ready).length,
    platforms: rows.map((row) => row.platform),
    rows: rows.map((row) => ({
      platform: row.platform,
      display_name: row.display_name,
      target: row.target,
      status: row.status,
      passed: row.passed,
      rollout_status: row.rollout_status,
      ready_for_realtime_annotations: row.ready_for_realtime_annotations,
      production_ready: row.production_ready,
      candidate_observation_ready: row.candidate_observation_ready,
      candidate_observer_message_type: row.candidate_observation?.message_type,
      candidate_observer_permission: row.candidate_observation?.required_permission,
      candidate_observer_endpoint: row.candidate_observation?.endpoint,
      blocking_count: row.blocking_count,
      warning_count: row.warning_count,
      next_actions: row.next_actions,
    })),
    reports: rows,
  };
}

export function assertMeetingPlatformLiveAdapterReadiness(platform, options = {}) {
  const readiness = buildMeetingPlatformLiveAdapterReadiness(platform, options);
  if (readiness.passed !== true) {
    throw new MeetingTimelineSdkError(`Meeting platform live adapter readiness failed for ${readiness.platform}`, {
      platform: readiness.platform,
      target: readiness.target,
      status: readiness.status,
      passed: readiness.passed,
      production_ready: readiness.production_ready,
      ready_for_realtime_annotations: readiness.ready_for_realtime_annotations,
      rollout_status: readiness.rollout_status,
      blocking_checks: readiness.blocking_checks,
      warnings: readiness.warnings,
      next_actions: readiness.next_actions,
      readiness,
    });
  }
  return readiness;
}

export function assertMeetingPlatformLiveAdapterReadinessMatrix(options = {}) {
  const matrix = buildMeetingPlatformLiveAdapterReadinessMatrix(options);
  if (matrix.reports.some((row) => row.passed !== true)) {
    throw new MeetingTimelineSdkError('Meeting platform live adapter readiness matrix failed', {
      platform_count: matrix.platform_count,
      passed_count: matrix.passed_count,
      ready_count: matrix.ready_count,
      warning_count: matrix.warning_count,
      blocked_count: matrix.blocked_count,
      realtime_ready_count: matrix.realtime_ready_count,
      production_ready_count: matrix.production_ready_count,
      failed_platforms: matrix.reports.filter((row) => row.passed !== true).map((row) => row.platform),
      next_actions: uniqueList(matrix.reports.flatMap((row) => row.next_actions ?? [])),
      matrix,
    });
  }
  return matrix;
}

export function buildMeetingPlatformLiveAdapterHandoff(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const plan = buildMeetingPlatformLiveAdapterPlan(key, options);
  const readiness = buildMeetingPlatformLiveAdapterReadiness(key, options);
  const candidateObservation = readiness.candidate_observation ?? buildCandidateObservationContract(key, options);
  const integration = buildPlatformIntegrationPlan(key, options);
  const runbook = buildMeetingPlatformAdaptationRunbook(key, options);
  return compactObject({
    type: 'meeting_platform_live_adapter_handoff',
    schema: MEETING_PLATFORM_LIVE_ADAPTER_HANDOFF_SCHEMA,
    schema_version: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION,
    platform: key,
    display_name: plan.display_name,
    status: readiness.status,
    passed: readiness.passed,
    target: readiness.target,
    rollout_status: readiness.rollout_status,
    recommended_mode: plan.recommended_mode,
    production_ready: readiness.production_ready,
    ready_for_realtime_annotations: readiness.ready_for_realtime_annotations,
    candidate_observation_ready: candidateObservation.ready === true,
    candidate_observation_contract: candidateObservation,
    sdk: liveAdapterSdkContract(),
    host_contract: hostContract(plan, candidateObservation),
    commands: liveAdapterCommands(key),
    evidence_paths: liveAdapterEvidencePaths(key),
    required_host_inputs: [
      'timeline client implementing startMeeting/endMeeting/insertMark/insertMarks/importTranscript',
      'candidate meeting window snapshots delivered through meeting_timeline.observe_candidates or observe_platform_candidates',
      'local meeting app observer snapshots with captured_at_ms',
      'annotation packets with captured_at_ms from the writing device or host app',
      'provider webhook records for later reconcile when the platform exposes them',
      'post-meeting transcript or artifact import only after the meeting if available',
    ],
    realtime_flow: [
      'observe local meeting app state',
      'observe active meeting window candidates before relying on provider events',
      'create or update the current meeting axis immediately',
      'insert annotation marks with captured_at_ms onto the active axis',
      'ingest provider events asynchronously for reconciliation',
      'import transcript/artifacts after the meeting without blocking realtime marks',
    ],
    outputs: [
      'meeting axis start/end events',
      'annotation marks aligned by captured_at_ms',
      'speaker activity markers when local observer can emit them',
      'meeting_platform_evidence_package for handoff and audit',
      'meeting_platform_live_adapter_readiness report for CI gating',
    ],
    plan,
    readiness,
    integration_plan: integration,
    runbook,
    next_actions: uniqueList([
      ...(readiness.next_actions ?? []),
      ...(plan.next_actions ?? []),
      ...(runbook.next_actions ?? []),
    ]),
  });
}

export function buildMeetingPlatformLiveAdapterHandoffBundle(options = {}) {
  const platforms = selectedPlatforms(options).map((platform) => normalizeMeetingPlatform(platform));
  const handoffs = platforms.map((platform) => buildMeetingPlatformLiveAdapterHandoff(platform, {
    ...options,
    platforms,
  }));
  const liveAdapterMatrix = buildMeetingPlatformLiveAdapterMatrix({
    ...options,
    platforms,
  });
  const readinessMatrix = buildMeetingPlatformLiveAdapterReadinessMatrix({
    ...options,
    platforms,
  });
  return {
    type: 'meeting_platform_live_adapter_handoff_bundle',
    schema: MEETING_PLATFORM_LIVE_ADAPTER_HANDOFF_BUNDLE_SCHEMA,
    schema_version: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION,
    platform_count: platforms.length,
    passed_count: readinessMatrix.passed_count,
    ready_count: readinessMatrix.ready_count,
    blocked_count: readinessMatrix.blocked_count,
    realtime_ready_count: readinessMatrix.realtime_ready_count,
    production_ready_count: readinessMatrix.production_ready_count,
    candidate_observer_count: readinessMatrix.candidate_observer_count,
    platforms,
    sdk: liveAdapterSdkContract(),
    commands: {
      build_capture_runtime: 'npm run meeting-app:extension:build',
      validate_dom_matrix: 'npm run meeting-app:evidence-matrix',
      validate_rollout: 'npm run meeting-platform:rollout-matrix',
      validate_live_readiness: 'npm run meeting-platform:live-readiness',
      package_smoke: 'npm run sdk:package-smoke',
    },
    evidence_paths: {
      meeting_app_evidence_dir: 'data/meeting-app-evidence',
      provider_evidence_dir: 'data/provider-evidence',
      evidence_package_dir: 'data/meeting-platform-evidence-packages',
      live_readiness_report: 'data/meeting-platform-live-readiness-report.json',
    },
    host_contract: {
      annotation_timestamp_field: 'captured_at_ms',
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      per_meeting_annotation_isolation_required: true,
      realtime_marks_do_not_wait_for_transcript: true,
      candidate_observation_required_for_host_axis_binding: true,
      candidate_observation_endpoint: MEETING_PLATFORM_CANDIDATE_OBSERVATION_ENDPOINT,
      candidate_observation_message_type: MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
      candidate_observation_runtime_action: 'observe_platform_candidates',
    },
    rows: handoffs.map((handoff) => ({
      platform: handoff.platform,
      display_name: handoff.display_name,
      status: handoff.status,
      passed: handoff.passed,
      rollout_status: handoff.rollout_status,
      recommended_mode: handoff.recommended_mode,
      ready_for_realtime_annotations: handoff.ready_for_realtime_annotations,
      production_ready: handoff.production_ready,
      candidate_observation_ready: handoff.candidate_observation_ready,
      candidate_observer_message_type: handoff.candidate_observation_contract?.message_type,
      candidate_observer_permission: handoff.candidate_observation_contract?.required_permission,
      candidate_observer_endpoint: handoff.candidate_observation_contract?.endpoint,
      next_actions: handoff.next_actions,
    })),
    handoffs,
    live_adapter_matrix: liveAdapterMatrix,
    readiness_matrix: readinessMatrix,
    next_actions: uniqueList(handoffs.flatMap((handoff) => handoff.next_actions ?? [])),
  };
}

export function createMeetingPlatformLiveAdapter(platform, clientOrOptions = {}, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const sourceInput = isTimelineClient(clientOrOptions)
    ? clientOrOptions
    : {
      ...(clientOrOptions ?? {}),
      ...(options.clientOptions ?? options.client_options ?? {}),
      baseUrl: firstNonEmpty(options.baseUrl, options.base_url, clientOrOptions?.baseUrl, clientOrOptions?.base_url),
    };
  const sourceOptions = {
    ...options,
    platform: key,
  };
  const source = createMeetingSourceAggregator(sourceInput, sourceOptions);
  const evidenceSession = createMeetingPlatformEvidenceSession(key, liveSessionOptions({
    ...options,
    id: firstNonEmpty(options.id, options.sessionId, options.session_id, `${key}-live-adapter`),
    source: firstNonEmpty(options.source, 'meeting_platform_live_adapter'),
  }));

  async function observeMeetingApp(input = {}, observeOptions = {}) {
    if (observeOptions.captureEvidence !== false && observeOptions.capture_evidence !== false) {
      evidenceSession.captureMeetingAppSnapshot(input, meetingAppRecordOptions(input, observeOptions));
    }
    const result = await source.observeMeetingApp(input, observeOptions);
    return resultWithEvidence(result, evidenceSession, 'observe_meeting_app', observeOptions);
  }

  async function ingestProvider(input = {}, payload, ingestOptions = {}) {
    if (ingestOptions.captureEvidence !== false && ingestOptions.capture_evidence !== false) {
      evidenceSession.captureProviderWebhook(input, payload, ingestOptions);
    }
    const result = await source.ingestProvider(key, providerPayload(input, payload), ingestOptions);
    return resultWithEvidence(result, evidenceSession, 'ingest_provider', ingestOptions);
  }

  return {
    schema: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA,
    schema_version: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION,
    platform: key,
    source,
    evidenceSession,
    session: evidenceSession,
    observeMeetingApp,
    observeApp: observeMeetingApp,
    async observeBrowser(input = {}, observeOptions = {}) {
      const result = await source.observeBrowser(input, observeOptions);
      return resultWithEvidence(result, evidenceSession, 'observe_browser', observeOptions);
    },
    async observeNative(input = {}, observeOptions = {}) {
      const result = await source.observeNative(input, observeOptions);
      return resultWithEvidence(result, evidenceSession, 'observe_native', observeOptions);
    },
    async observeLocal(input = {}, observeOptions = {}) {
      const result = await source.observeLocal(input, observeOptions);
      return resultWithEvidence(result, evidenceSession, 'observe_local', observeOptions);
    },
    async observeLocalCandidates(candidates = [], observeOptions = {}) {
      const result = await source.observeLocalCandidates(candidates, observeOptions);
      return resultWithEvidence(result, evidenceSession, 'observe_local_candidates', observeOptions);
    },
    ingestProvider,
    async ingestSignals(signals = [], ingestOptions = {}) {
      const result = await source.ingestSignals(signals, ingestOptions);
      return resultWithEvidence(result, evidenceSession, 'ingest_signals', ingestOptions);
    },
    captureMeetingAppSnapshot(input = {}, captureOptions = {}) {
      const record = evidenceSession.captureMeetingAppSnapshot(input, meetingAppRecordOptions(input, captureOptions));
      return operationWithEvidence('capture_meeting_app_snapshot', record, evidenceSession, captureOptions);
    },
    captureProviderWebhook(input = {}, payload, captureOptions = {}) {
      const record = evidenceSession.captureProviderWebhook(input, payload, captureOptions);
      return operationWithEvidence('capture_provider_webhook', record, evidenceSession, captureOptions);
    },
    async insertAnnotation(input = {}, markOptions = {}) {
      const result = await executeRealtimeAnnotationPipeline(source, key, input, markOptions, options);
      return operationWithEvidence('insert_annotation', result, evidenceSession, markOptions);
    },
    async insertMark(input = {}, markOptions = {}) {
      const result = await source.insertMark({
        platform: key,
        ...input,
      }, markOptions);
      return operationWithEvidence('insert_mark', result, evidenceSession, markOptions);
    },
    async insertMarks(inputs = [], markOptions = {}) {
      const rows = Array.isArray(inputs) ? inputs : [inputs];
      const result = await source.insertMarks(rows.map((input) => ({
        platform: key,
        ...input,
      })), markOptions);
      return operationWithEvidence('insert_marks', result, evidenceSession, markOptions);
    },
    importTranscript(input = {}, transcriptOptions = {}) {
      return source.importTranscript({
        platform: key,
        ...input,
      }, transcriptOptions);
    },
    startMeeting(input = {}) {
      return source.startMeeting({
        platform: key,
        ...input,
      });
    },
    endMeeting(input = {}) {
      return source.endMeeting(input);
    },
    summary(summaryOptions = {}) {
      return evidenceSession.summary(summaryOptions);
    },
    correlation(correlationOptions = {}) {
      return evidenceSession.correlation(correlationOptions);
    },
    strategy(strategyOptions = {}) {
      return evidenceSession.strategy(strategyOptions);
    },
    exportPackage(packageOptions = {}) {
      return evidenceSession.exportPackage(packageOptions);
    },
    verify(verifyOptions = {}) {
      return evidenceSession.verify(verifyOptions);
    },
    getState() {
      return {
        type: 'meeting_platform_live_adapter_state',
        schema: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA,
        schema_version: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION,
        platform: key,
        source: source.getState(),
        evidence: evidenceSession.getState(),
      };
    },
    reset(nextState = {}) {
      return {
        source: source.reset(nextState.source ?? nextState.meeting_source ?? {}),
        evidence: evidenceSession.reset(),
      };
    },
  };
}

export function createMeetingPlatformLiveAdapterSuite(clientOrOptions = {}, options = {}) {
  const cache = new Map();
  const platforms = selectedPlatforms(options).map((platform) => normalizeMeetingPlatform(platform));

  function adapter(platform, adapterOptions = {}) {
    const key = normalizeMeetingPlatform(platform);
    const cacheKey = `${key}:${adapterOptions.id ?? adapterOptions.sessionId ?? adapterOptions.session_id ?? 'default'}`;
    if (!cache.has(cacheKey)) {
      cache.set(cacheKey, createMeetingPlatformLiveAdapter(key, clientOrOptions, {
        ...options,
        ...adapterOptions,
      }));
    }
    return cache.get(cacheKey);
  }

  return {
    schema: MEETING_PLATFORM_LIVE_ADAPTER_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION,
    platforms,
    adapter,
    liveAdapter: adapter,
    adapters(adapterOptions = {}) {
      return Object.fromEntries(platforms.map((platform) => [platform, adapter(platform, adapterOptions)]));
    },
    plan(platform, planOptions = {}) {
      return buildMeetingPlatformLiveAdapterPlan(platform, {
        ...options,
        ...planOptions,
      });
    },
    readiness(platform, readinessOptions = {}) {
      return buildMeetingPlatformLiveAdapterReadiness(platform, {
        ...options,
        ...readinessOptions,
        adapter: firstNonEmpty(
          readinessOptions.adapter,
          readinessOptions.liveAdapter,
          readinessOptions.live_adapter,
          adapter(platform, readinessOptions.adapterOptions ?? readinessOptions.adapter_options ?? {}),
        ),
      });
    },
    assertReadiness(platform, readinessOptions = {}) {
      return assertMeetingPlatformLiveAdapterReadiness(platform, {
        ...options,
        ...readinessOptions,
        adapter: firstNonEmpty(
          readinessOptions.adapter,
          readinessOptions.liveAdapter,
          readinessOptions.live_adapter,
          adapter(platform, readinessOptions.adapterOptions ?? readinessOptions.adapter_options ?? {}),
        ),
      });
    },
    matrix(matrixOptions = {}) {
      return buildMeetingPlatformLiveAdapterMatrix({
        ...options,
        ...matrixOptions,
        platforms: matrixOptions.platforms ?? matrixOptions.platform_keys ?? platforms,
      });
    },
    readinessMatrix(readinessOptions = {}) {
      return buildMeetingPlatformLiveAdapterReadinessMatrix({
        ...options,
        ...readinessOptions,
        platforms: readinessOptions.platforms ?? readinessOptions.platform_keys ?? platforms,
      });
    },
    assertReadinessMatrix(readinessOptions = {}) {
      return assertMeetingPlatformLiveAdapterReadinessMatrix({
        ...options,
        ...readinessOptions,
        platforms: readinessOptions.platforms ?? readinessOptions.platform_keys ?? platforms,
      });
    },
    handoff(platform, handoffOptions = {}) {
      return buildMeetingPlatformLiveAdapterHandoff(platform, {
        ...options,
        ...handoffOptions,
      });
    },
    handoffBundle(handoffOptions = {}) {
      return buildMeetingPlatformLiveAdapterHandoffBundle({
        ...options,
        ...handoffOptions,
        platforms: handoffOptions.platforms ?? handoffOptions.platform_keys ?? platforms,
      });
    },
    summary(summaryOptions = {}) {
      const matrix = buildMeetingPlatformLiveAdapterMatrix({
        ...options,
        ...summaryOptions,
        platforms: summaryOptions.platforms ?? summaryOptions.platform_keys ?? platforms,
      });
      return compactObject({
        type: 'meeting_platform_live_adapter_suite_summary',
        schema: MEETING_PLATFORM_LIVE_ADAPTER_MATRIX_SCHEMA,
        schema_version: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION,
        platform_count: matrix.platform_count,
        production_ready_count: matrix.production_ready_count,
        pilot_ready_count: matrix.pilot_ready_count,
        platforms: matrix.platforms,
        rows: matrix.rows,
      });
    },
    getState() {
      return {
        type: 'meeting_platform_live_adapter_suite_state',
        schema: MEETING_PLATFORM_LIVE_ADAPTER_MATRIX_SCHEMA,
        schema_version: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION,
        platforms,
        adapter_count: cache.size,
        adapters: Object.fromEntries([...cache.entries()].map(([key, item]) => [key, item.getState()])),
      };
    },
    reset(nextState = {}) {
      const states = Object.fromEntries([...cache.entries()].map(([key, item]) => [key, item.reset(nextState[key] ?? {})]));
      cache.clear();
      return {
        removed_adapters: Object.keys(states).length,
        adapters: states,
      };
    },
  };
}
