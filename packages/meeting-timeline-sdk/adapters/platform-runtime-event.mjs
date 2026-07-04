import {
  MeetingTimelineApiError,
  MeetingTimelineSdkError,
  compactObject,
  normalizeAbsoluteMs,
} from '../index.mjs';
import {
  MEETING_PLATFORM_KEYS,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA = 'meeting_platform_runtime_event';
export const MEETING_PLATFORM_RUNTIME_EVENT_PLAN_SCHEMA = 'meeting_platform_runtime_event_plan';
export const MEETING_PLATFORM_RUNTIME_EVENT_PLAN_MATRIX_SCHEMA = 'meeting_platform_runtime_event_plan_matrix';
export const MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA_VERSION = 1;
export const MEETING_PLATFORM_RUNTIME_EVENT_ENDPOINT = '/api/meeting-platform/runtime-events';

export const MEETING_PLATFORM_RUNTIME_EVENT_ACTIONS = Object.freeze([
  'observe_meeting_app',
  'observe_platform_candidates',
  'provider_event',
  'insert_annotation',
  'speaker_track',
  'participant_track',
  'timeline_view',
  'runtime_bundles',
  'registry',
  'manifest',
  'run_manifest',
  'readiness',
  'handoff_readiness',
  'run_handoff_readiness',
]);

const ACTION_ALIASES = new Map([
  ['observe', 'observe_meeting_app'],
  ['observe_app', 'observe_meeting_app'],
  ['observe_meeting_app', 'observe_meeting_app'],
  ['meeting_app_snapshot', 'observe_meeting_app'],
  ['snapshot', 'observe_meeting_app'],
  ['observe_candidates', 'observe_platform_candidates'],
  ['observe_platform_candidates', 'observe_platform_candidates'],
  ['observe_meeting_environment', 'observe_platform_candidates'],
  ['meeting_environment_snapshot', 'observe_platform_candidates'],
  ['platform_candidate_observation', 'observe_platform_candidates'],
  ['provider', 'provider_event'],
  ['provider_event', 'provider_event'],
  ['ingest_provider', 'provider_event'],
  ['webhook', 'provider_event'],
  ['annotation', 'insert_annotation'],
  ['mark', 'insert_annotation'],
  ['insert_annotation', 'insert_annotation'],
  ['insert_mark', 'insert_annotation'],
  ['speaker', 'speaker_track'],
  ['speaker_track', 'speaker_track'],
  ['participant', 'participant_track'],
  ['participant_track', 'participant_track'],
  ['view', 'timeline_view'],
  ['timeline_view', 'timeline_view'],
  ['runtime_bundle_matrix', 'runtime_bundles'],
  ['runtime_bundles', 'runtime_bundles'],
  ['platform_registry', 'registry'],
  ['registry', 'registry'],
  ['runtime_manifest', 'manifest'],
  ['integration_manifest', 'manifest'],
  ['manifest', 'manifest'],
  ['run_manifest', 'run_manifest'],
  ['runtime_manifest_run', 'run_manifest'],
  ['integration_manifest_run', 'run_manifest'],
  ['integration_runtime_run_manifest', 'run_manifest'],
  ['integration_runtime_manifest_run', 'run_manifest'],
  ['live_readiness', 'readiness'],
  ['readiness', 'readiness'],
  ['handoff_readiness', 'handoff_readiness'],
  ['run_handoff_readiness', 'run_handoff_readiness'],
  ['handoff_readiness_run', 'run_handoff_readiness'],
]);

const PLATFORM_REQUIRED_ACTIONS = new Set([
  'observe_meeting_app',
  'provider_event',
  'insert_annotation',
  'speaker_track',
  'participant_track',
  'timeline_view',
]);

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
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_PLATFORM_KEYS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date);
}

function maybeAbsoluteMs(value, fieldName) {
  if (value == null || value === '') return undefined;
  return normalizeAbsoluteMs(value, fieldName);
}

function nowMs(options = {}) {
  const now = firstNonEmpty(options.now, options.clock);
  if (typeof now === 'function') return maybeAbsoluteMs(now(), 'sent_at_ms');
  return Date.now();
}

function sampleUrlFor(platform, options = {}) {
  return firstNonEmpty(
    options.url,
    options.href,
    options.meeting_url,
    options.meetingUrl,
    {
      google_meet: 'https://meet.google.com/abc-defg-hij',
      microsoft_teams: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
      zoom: 'https://zoom.us/j/987654321',
      webex: 'https://example.webex.com/meet/sample',
      lark: 'https://vc.feishu.cn/j/123456789',
      local_detector: 'local://meeting-window/sample',
    }[platform],
  );
}

function providerStartEventFor(platform) {
  return {
    google_meet: 'google.workspace.meet.conference.v2.started',
    microsoft_teams: 'microsoft.graph.callStarted',
    zoom: 'meeting.started',
    webex: 'meeting.started',
    lark: 'vc.meeting.meeting_started_v1',
    local_detector: 'local.meeting.started',
  }[platform] ?? 'meeting.started';
}

function providerEndEventFor(platform) {
  return {
    google_meet: 'google.workspace.meet.conference.v2.ended',
    microsoft_teams: 'microsoft.graph.callEnded',
    zoom: 'meeting.ended',
    webex: 'meeting.ended',
    lark: 'vc.meeting.meeting_ended_v1',
    local_detector: 'local.meeting.ended',
  }[platform] ?? 'meeting.ended';
}

function valuePlatform(value) {
  if (value == null || value === '') return undefined;
  return normalizeMeetingPlatform(value);
}

function normalizedAction(value) {
  const key = String(value ?? '').trim().toLowerCase().replace(/[-.\s]+/g, '_');
  return ACTION_ALIASES.get(key) ?? key;
}

function payloadObject(payload) {
  return isPlainObject(payload) ? payload : {};
}

function eventPayload(input = {}, options = {}) {
  return firstNonEmpty(
    options.payload,
    input.payload,
    input.provider_event,
    input.providerEvent,
    input.snapshot,
    input.meeting_app_snapshot,
    input.meetingAppSnapshot,
    input.annotation,
    input.mark,
    input.raw,
  );
}

function eventIdFor(input = {}, options = {}) {
  return firstNonEmpty(
    options.event_id,
    options.eventId,
    input.event_id,
    input.eventId,
    input.id,
    input.source_event_id,
    input.sourceEventId,
  );
}

function correlationIdFor(input = {}, options = {}) {
  return firstNonEmpty(
    options.correlation_id,
    options.correlationId,
    input.correlation_id,
    input.correlationId,
    input.meeting_id,
    input.meetingId,
    input.meeting?.meeting_id,
    input.meeting?.meetingId,
    input.current_meeting?.meeting_id,
    input.currentMeeting?.meetingId,
  );
}

function sentAtFor(input = {}, options = {}) {
  return maybeAbsoluteMs(firstNonEmpty(
    options.sent_at_ms,
    options.sentAtMs,
    options.sent_at,
    options.sentAt,
    input.sent_at_ms,
    input.sentAtMs,
    input.sent_at,
    input.sentAt,
  ), 'sent_at_ms') ?? nowMs(options);
}

function platformFor(input = {}, options = {}, payload = undefined) {
  const value = firstNonEmpty(
    options.platform,
    input.platform,
    input.provider,
    input.adapter,
    payload?.platform,
    payload?.provider,
    payload?.meeting?.platform,
    payload?.current_meeting?.platform,
    payload?.currentMeeting?.platform,
  );
  return value == null ? undefined : valuePlatform(value);
}

function actionFor(input = {}, options = {}) {
  return normalizedAction(firstNonEmpty(
    input.action,
    input.kind,
    input.event_kind,
    input.eventKind,
    input.runtime_action,
    input.runtimeAction,
    input.type,
    options.action,
    options.kind,
  ));
}

function actionDefaults(event = {}, payload = undefined, input = {}) {
  const objectPayload = payloadObject(payload);
  if (event.action === 'observe_meeting_app' && !event.snapshot && !event.meeting_app_snapshot) {
    event.snapshot = firstNonEmpty(input.snapshot, input.meeting_app_snapshot, input.meetingAppSnapshot, input.dom, payload);
  }
  if (event.action === 'observe_platform_candidates') {
    event.environment = firstNonEmpty(input.environment, input.meeting_environment, input.meetingEnvironment, input.snapshot, payload);
    event.windows = firstNonEmpty(input.windows, objectPayload.windows, event.windows);
    event.tabs = firstNonEmpty(input.tabs, objectPayload.tabs, event.tabs);
    event.applications = firstNonEmpty(input.applications, input.apps, objectPayload.applications, objectPayload.apps, event.applications);
    event.candidates = firstNonEmpty(input.candidates, objectPayload.candidates, event.candidates);
  }
  if (event.action === 'provider_event' && !event.provider_event && !event.providerEvent && !event.raw) {
    event.provider_event = payload ?? input.raw ?? objectPayload;
  }
  if (event.action === 'insert_annotation' && !event.annotation && !event.mark) {
    event.annotation = firstNonEmpty(input.annotation, input.mark, objectPayload.annotation, objectPayload.mark, payload);
  }
  if (['speaker_track', 'participant_track'].includes(event.action) && !event.signals) {
    event.signals = firstNonEmpty(input.signals, objectPayload.signals);
  }
  if (event.action === 'timeline_view') {
    event.meeting = firstNonEmpty(input.meeting, objectPayload.meeting, event.meeting);
    event.annotations = firstNonEmpty(input.annotations, objectPayload.annotations, event.annotations);
    event.speakerTrack = firstNonEmpty(input.speakerTrack, input.speaker_track, objectPayload.speakerTrack, objectPayload.speaker_track);
    event.participantTrack = firstNonEmpty(input.participantTrack, input.participant_track, objectPayload.participantTrack, objectPayload.participant_track);
  }
  return event;
}

export function normalizeMeetingPlatformRuntimeEventAction(action) {
  const normalized = normalizedAction(action);
  if (!MEETING_PLATFORM_RUNTIME_EVENT_ACTIONS.includes(normalized)) {
    throw new MeetingTimelineSdkError('Unsupported meeting platform runtime event action', {
      action,
      normalized,
      supported_actions: MEETING_PLATFORM_RUNTIME_EVENT_ACTIONS,
    });
  }
  return normalized;
}

export function buildMeetingPlatformRuntimeEvent(input = {}, options = {}) {
  const payload = eventPayload(input, options);
  const objectPayload = payloadObject(payload);
  const action = normalizeMeetingPlatformRuntimeEventAction(actionFor(input, options));
  const platform = platformFor(input, options, objectPayload);
  const event = actionDefaults(compactObject({
    ...objectPayload,
    ...input,
    type: MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA,
    schema: MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA,
    schema_version: MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA_VERSION,
    action,
    platform,
    source: firstNonEmpty(options.source, input.source),
    event_id: eventIdFor(input, options),
    correlation_id: correlationIdFor(input, options),
    sent_at_ms: sentAtFor(input, options),
    payload,
  }), payload, input);
  return assertMeetingPlatformRuntimeEvent(event, options);
}

export function assertMeetingPlatformRuntimeEvent(event = {}, options = {}) {
  const action = normalizeMeetingPlatformRuntimeEventAction(firstNonEmpty(event.action, options.action));
  if (PLATFORM_REQUIRED_ACTIONS.has(action) && !event.platform) {
    throw new MeetingTimelineSdkError('Meeting platform runtime event requires platform', {
      action,
      required_field: 'platform',
    });
  }
  return {
    ...event,
    action,
  };
}

export function buildMeetingPlatformObserveRuntimeEvent(platform, snapshot = {}, options = {}) {
  return buildMeetingPlatformRuntimeEvent({
    action: 'observe_meeting_app',
    platform,
    snapshot,
    payload: snapshot,
  }, options);
}

export function buildMeetingPlatformCandidateObservationRuntimeEvent(input = {}, options = {}) {
  return buildMeetingPlatformRuntimeEvent({
    ...input,
    action: 'observe_platform_candidates',
    payload: input,
  }, options);
}

export function buildMeetingPlatformProviderRuntimeEvent(platform, payload = {}, options = {}) {
  return buildMeetingPlatformRuntimeEvent({
    action: 'provider_event',
    platform,
    provider_event: payload,
    payload,
  }, options);
}

export function buildMeetingPlatformAnnotationRuntimeEvent(platform, annotationInput = {}, options = {}) {
  const annotation = firstNonEmpty(annotationInput.annotation, annotationInput.mark, annotationInput);
  return buildMeetingPlatformRuntimeEvent({
    ...annotationInput,
    action: 'insert_annotation',
    platform,
    annotation,
    payload: annotationInput,
  }, options);
}

export function buildMeetingPlatformSpeakerTrackRuntimeEvent(platform, input = {}, options = {}) {
  return buildMeetingPlatformRuntimeEvent({
    ...input,
    action: 'speaker_track',
    platform,
    payload: input,
  }, options);
}

export function buildMeetingPlatformParticipantTrackRuntimeEvent(platform, input = {}, options = {}) {
  return buildMeetingPlatformRuntimeEvent({
    ...input,
    action: 'participant_track',
    platform,
    payload: input,
  }, options);
}

export function buildMeetingPlatformTimelineViewRuntimeEvent(platform, input = {}, options = {}) {
  return buildMeetingPlatformRuntimeEvent({
    ...input,
    action: 'timeline_view',
    platform,
    payload: input,
  }, options);
}

export function buildMeetingPlatformRunManifestRuntimeEvent(input = {}, options = {}) {
  return buildMeetingPlatformRuntimeEvent({
    ...input,
    action: 'run_manifest',
    payload: input,
  }, options);
}

export function buildMeetingPlatformRunHandoffReadinessRuntimeEvent(input = {}, options = {}) {
  return buildMeetingPlatformRuntimeEvent({
    ...input,
    action: 'run_handoff_readiness',
    payload: input,
  }, options);
}

function runtimeEventExamples(platform, options = {}) {
  const capturedAtMs = maybeAbsoluteMs(firstNonEmpty(
    options.capturedAtMs,
    options.captured_at_ms,
    options.sample_at_ms,
    1_782_614_400_000,
  ), 'captured_at_ms');
  const url = sampleUrlFor(platform, options);
  const currentMeeting = compactObject({
    platform,
    meeting_id: firstNonEmpty(options.meeting_id, options.meetingId, `${platform}-sample-meeting`),
    external_meeting_id: firstNonEmpty(options.external_meeting_id, options.externalMeetingId, `${platform}-external-sample`),
    meeting_url: url,
    title: firstNonEmpty(options.title, 'Sample meeting'),
    start_time_ms: capturedAtMs,
  });
  const annotation = {
    id: 'note-001',
    label: 'why?',
    kind: 'question',
    captured_at_ms: capturedAtMs + 15_000,
    source: 'annotation_device',
  };
  const speakerSignals = [{
    type: 'speaker_started',
    platform,
    speaker: {
      id: 'speaker-001',
      display_name: 'Alex',
    },
    occurred_at_ms: capturedAtMs + 20_000,
    confidence: 0.82,
    source: 'local_observer',
  }];
  const participantSignals = [{
    type: 'participant_joined',
    platform,
    participant: {
      id: 'participant-001',
      display_name: 'Alex',
    },
    occurred_at_ms: capturedAtMs + 5_000,
    source: 'local_observer',
  }];
  const providerPayload = {
    event_type: providerStartEventFor(platform),
    event_ts: capturedAtMs,
    occurred_at_ms: capturedAtMs,
    meeting: currentMeeting,
  };
  const candidateEnvironment = {
    windows: [{
      id: `${platform}-window-1`,
      focused: true,
      application: { name: firstNonEmpty(options.applicationName, options.application_name, platform) },
      tabs: [{
        id: `${platform}-tab-1`,
        active: true,
        url,
        title: currentMeeting.title,
        observed_at_ms: capturedAtMs,
      }],
    }],
  };

  return {
    observe_meeting_app: buildMeetingPlatformObserveRuntimeEvent(platform, {
      platform,
      url,
      title: currentMeeting.title,
      observed_at_ms: capturedAtMs,
      in_meeting: true,
      active_speaker: speakerSignals[0].speaker,
    }, options),
    observe_platform_candidates: buildMeetingPlatformCandidateObservationRuntimeEvent(candidateEnvironment, options),
    provider_event: buildMeetingPlatformProviderRuntimeEvent(platform, providerPayload, options),
    insert_annotation: buildMeetingPlatformAnnotationRuntimeEvent(platform, {
      annotation,
      current_meeting: currentMeeting,
    }, options),
    speaker_track: buildMeetingPlatformSpeakerTrackRuntimeEvent(platform, {
      current_meeting: currentMeeting,
      signals: speakerSignals,
    }, options),
    participant_track: buildMeetingPlatformParticipantTrackRuntimeEvent(platform, {
      current_meeting: currentMeeting,
      signals: participantSignals,
    }, options),
    timeline_view: buildMeetingPlatformTimelineViewRuntimeEvent(platform, {
      meeting: currentMeeting,
      annotations: [annotation],
      speakerTrack: {
        marks: speakerSignals,
      },
      participantTrack: {
        marks: participantSignals,
      },
    }, options),
    run_manifest: buildMeetingPlatformRunManifestRuntimeEvent({
      platforms: [platform],
      target: 'production',
      requireHandoffReady: false,
    }, options),
    run_handoff_readiness: buildMeetingPlatformRunHandoffReadinessRuntimeEvent({
      platforms: [platform],
      target: 'production',
    }, options),
  };
}

function runtimeEventActionRows(platform, endpoint) {
  return [
    {
      action: 'observe_meeting_app',
      client_method: 'observeMeetingApp',
      producer: 'browser_extension_or_native_observer',
      realtime_role: 'open_or_keep_current_axis',
      required_fields: ['platform', 'snapshot.observed_at_ms_or_sent_at_ms'],
      recommended_fields: ['snapshot.url', 'snapshot.title', 'snapshot.active_speaker'],
      provider_dependency: false,
      transcript_dependency: false,
      endpoint,
    },
    {
      action: 'observe_platform_candidates',
      client_method: 'observePlatformCandidates',
      producer: 'native_host_or_browser_window_detector',
      realtime_role: 'select_current_meeting_and_open_or_close_axis',
      required_fields: ['windows_or_tabs_or_applications_or_candidates'],
      recommended_fields: ['windows[].focused', 'tabs[].active', 'tabs[].url', 'applications[].name'],
      provider_dependency: false,
      transcript_dependency: false,
      endpoint,
    },
    {
      action: 'insert_annotation',
      client_method: 'insertAnnotation',
      producer: 'annotation_device_or_companion_app',
      realtime_role: 'primary_realtime_mark_write',
      required_fields: ['platform', 'annotation.captured_at_ms'],
      recommended_fields: ['annotation.id', 'annotation.label', 'current_meeting.meeting_id_or_url'],
      provider_dependency: false,
      transcript_dependency: false,
      endpoint,
    },
    {
      action: 'speaker_track',
      client_method: 'speakerTrack',
      producer: 'local_active_speaker_observer',
      realtime_role: 'speaker_position_marks_without_transcript',
      required_fields: ['platform', 'signals[].occurred_at_ms', 'signals[].speaker'],
      recommended_fields: ['signals[].confidence', 'current_meeting.meeting_id_or_url'],
      provider_dependency: false,
      transcript_dependency: false,
      endpoint,
    },
    {
      action: 'participant_track',
      client_method: 'participantTrack',
      producer: 'local_roster_observer_or_provider_reconcile',
      realtime_role: 'participant_join_leave_marks',
      required_fields: ['platform', 'signals[].occurred_at_ms', 'signals[].participant'],
      recommended_fields: ['signals[].source', 'current_meeting.meeting_id_or_url'],
      provider_dependency: false,
      transcript_dependency: false,
      endpoint,
    },
    {
      action: 'provider_event',
      client_method: 'ingestProvider',
      producer: 'official_webhook_or_event_subscription',
      realtime_role: 'reconcile_and_backfill_only',
      required_fields: ['platform', 'provider_event'],
      recommended_fields: ['provider_event.event_ts_or_occurred_at_ms', 'provider_event.meeting'],
      provider_dependency: true,
      transcript_dependency: false,
      endpoint,
    },
    {
      action: 'timeline_view',
      client_method: 'timelineView',
      producer: 'host_ui',
      realtime_role: 'render_view_model_request',
      required_fields: ['platform', 'meeting'],
      recommended_fields: ['annotations', 'speakerTrack', 'participantTrack'],
      provider_dependency: false,
      transcript_dependency: false,
      endpoint,
    },
    {
      action: 'runtime_bundles',
      client_method: 'runtimeBundles',
      producer: 'host_ci_or_admin_panel',
      realtime_role: 'runtime_bundle_inspection',
      required_fields: [],
      recommended_fields: ['platforms'],
      provider_dependency: false,
      transcript_dependency: false,
      endpoint,
    },
    {
      action: 'registry',
      client_method: 'registry',
      producer: 'host_ci_or_admin_panel',
      realtime_role: 'platform_registry_inspection',
      required_fields: [],
      recommended_fields: ['platforms'],
      provider_dependency: false,
      transcript_dependency: false,
      endpoint,
    },
    {
      action: 'manifest',
      client_method: 'manifest',
      producer: 'host_ci_or_admin_panel',
      realtime_role: 'static_integration_contract_gate',
      required_fields: [],
      recommended_fields: ['platforms'],
      provider_dependency: false,
      transcript_dependency: false,
      endpoint,
    },
    {
      action: 'run_manifest',
      client_method: 'runManifest',
      producer: 'host_ci_or_admin_panel',
      realtime_role: 'executable_integration_handoff_gate',
      required_fields: [],
      recommended_fields: ['platforms', 'requireHandoffReady', 'evidencePackage'],
      provider_dependency: false,
      transcript_dependency: false,
      endpoint,
    },
    {
      action: 'readiness',
      client_method: 'readiness',
      producer: 'host_ci_or_admin_panel',
      realtime_role: 'live_adapter_readiness_inspection',
      required_fields: [],
      recommended_fields: ['platforms'],
      provider_dependency: false,
      transcript_dependency: false,
      endpoint,
    },
    {
      action: 'handoff_readiness',
      client_method: 'handoffReadiness',
      producer: 'host_ci_or_admin_panel',
      realtime_role: 'static_handoff_readiness_inspection',
      required_fields: [],
      recommended_fields: ['platforms', 'target'],
      provider_dependency: false,
      transcript_dependency: false,
      endpoint,
    },
    {
      action: 'run_handoff_readiness',
      client_method: 'runHandoffReadiness',
      producer: 'host_ci_or_admin_panel',
      realtime_role: 'executable_handoff_readiness_gate',
      required_fields: [],
      recommended_fields: ['platforms', 'target', 'evidencePackage'],
      provider_dependency: false,
      transcript_dependency: false,
      endpoint,
    },
  ].map((row) => ({
    platform,
    ...row,
  }));
}

export function buildMeetingPlatformRuntimeEventPlan(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const endpoint = meetingPlatformRuntimeEventEndpoint(options);
  const examples = runtimeEventExamples(key, {
    ...options,
    source: firstNonEmpty(options.source, 'runtime_event_plan'),
  });
  const actionRows = runtimeEventActionRows(key, endpoint);
  return {
    type: 'meeting_platform_runtime_event_plan',
    schema: MEETING_PLATFORM_RUNTIME_EVENT_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA_VERSION,
    platform: key,
    endpoint,
    client_factory: 'createMeetingPlatformRuntimeEventClient',
    event_schema: MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA,
    supported_actions: MEETING_PLATFORM_RUNTIME_EVENT_ACTIONS,
    provider_start_event_example: providerStartEventFor(key),
    provider_end_event_example: providerEndEventFor(key),
    realtime_contract: {
      primary_clock_field: 'captured_at_ms',
      provider_events_required_for_realtime: false,
      transcript_required_for_realtime: false,
      local_observer_required_for_reliable_start_end: true,
      annotation_should_use_device_capture_time: true,
    },
    imports: {
      runtime_event: '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-event',
      integration_runtime: '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime',
      realtime_annotation: '@ai-annotation/meeting-timeline-sdk/adapters/platform-realtime-annotation',
      speaker_track: '@ai-annotation/meeting-timeline-sdk/adapters/platform-speaker-track',
      participant_track: '@ai-annotation/meeting-timeline-sdk/adapters/platform-participant-track',
      timeline_view: '@ai-annotation/meeting-timeline-sdk/adapters/platform-timeline-view',
    },
    sequence: [
      'create runtime event client with the host baseUrl',
      'send observe_meeting_app when the host detects an active meeting window',
      'send observe_platform_candidates when the host only has multi-window or native helper state',
      'send insert_annotation immediately when the annotation device finishes a mark',
      'send speaker_track or participant_track only after local filtering/debounce',
      'send provider_event as reconcile/backfill evidence when official events arrive',
      'request timeline_view from the host UI instead of duplicating SVG positioning logic',
      'send run_manifest or run_handoff_readiness from CI/admin tooling before host handoff',
    ],
    actions: actionRows,
    examples,
    next_actions: [
      'wire_runtime_event_endpoint_in_host_project',
      'forward_local_observer_snapshots_before_waiting_for_provider_events',
      'include_captured_at_ms_on_every_annotation',
      'keep_transcript_import_post_meeting_only',
    ],
  };
}

export function buildMeetingPlatformRuntimeEventPlanMatrix(options = {}) {
  const plans = selectedPlatforms(options).map((platform) => buildMeetingPlatformRuntimeEventPlan(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_runtime_event_plan_matrix',
    schema: MEETING_PLATFORM_RUNTIME_EVENT_PLAN_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA_VERSION,
    platform_count: plans.length,
    platforms: plans.map((plan) => plan.platform),
    realtime_provider_dependency_count: plans.filter((plan) => plan.realtime_contract.provider_events_required_for_realtime).length,
    transcript_realtime_dependency_count: plans.filter((plan) => plan.realtime_contract.transcript_required_for_realtime).length,
    rows: plans.flatMap((plan) => plan.actions.map((action) => ({
      platform: plan.platform,
      action: action.action,
      client_method: action.client_method,
      producer: action.producer,
      realtime_role: action.realtime_role,
      provider_dependency: action.provider_dependency,
      transcript_dependency: action.transcript_dependency,
    }))),
    plans,
    next_actions: unique(plans.flatMap((plan) => plan.next_actions ?? [])),
  };
}

export function meetingPlatformRuntimeEventEndpoint(options = {}) {
  const path = firstNonEmpty(options.path, options.endpoint, options.runtimeEventEndpoint, options.runtime_event_endpoint, MEETING_PLATFORM_RUNTIME_EVENT_ENDPOINT);
  const baseUrl = firstNonEmpty(options.baseUrl, options.base_url);
  if (!baseUrl) return path;
  return new URL(path, baseUrl).toString();
}

async function parseResponse(response) {
  const text = await response.text();
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

export function createMeetingPlatformRuntimeEventClient(options = {}) {
  const fetchImpl = options.fetch ?? options.fetchImpl ?? globalThis.fetch;
  if (typeof fetchImpl !== 'function') {
    throw new MeetingTimelineSdkError('fetch is required for meeting platform runtime event client');
  }
  const endpoint = meetingPlatformRuntimeEventEndpoint(options);
  const defaultHeaders = {
    'content-type': 'application/json',
    ...(options.headers ?? {}),
  };

  async function send(eventInput = {}, sendOptions = {}) {
    const event = eventInput?.schema === MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA
      ? assertMeetingPlatformRuntimeEvent(eventInput, sendOptions)
      : buildMeetingPlatformRuntimeEvent(eventInput, {
        ...options,
        ...sendOptions,
      });
    const response = await fetchImpl(sendOptions.endpoint ?? endpoint, {
      method: 'POST',
      ...(sendOptions.fetchOptions ?? sendOptions.fetch_options ?? {}),
      headers: {
        ...defaultHeaders,
        ...(sendOptions.headers ?? {}),
      },
      body: JSON.stringify(event),
      signal: sendOptions.signal,
    });
    const body = await parseResponse(response);
    if (!response.ok) {
      throw new MeetingTimelineApiError('Meeting platform runtime event request failed', {
        status: response.status,
        body,
        event,
      });
    }
    return body;
  }

  return {
    type: 'meeting_platform_runtime_event_client',
    schema: MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA,
    schema_version: MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA_VERSION,
    endpoint,
    send,
    observeMeetingApp(platform, snapshot = {}, observeOptions = {}) {
      return send(buildMeetingPlatformObserveRuntimeEvent(platform, snapshot, observeOptions), observeOptions);
    },
    observePlatformCandidates(input = {}, observeOptions = {}) {
      return send(buildMeetingPlatformCandidateObservationRuntimeEvent(input, observeOptions), observeOptions);
    },
    ingestProvider(platform, payload = {}, ingestOptions = {}) {
      return send(buildMeetingPlatformProviderRuntimeEvent(platform, payload, ingestOptions), ingestOptions);
    },
    insertAnnotation(platform, annotationInput = {}, markOptions = {}) {
      return send(buildMeetingPlatformAnnotationRuntimeEvent(platform, annotationInput, markOptions), markOptions);
    },
    insertMark(platform, annotationInput = {}, markOptions = {}) {
      return this.insertAnnotation(platform, annotationInput, markOptions);
    },
    speakerTrack(platform, input = {}, trackOptions = {}) {
      return send(buildMeetingPlatformSpeakerTrackRuntimeEvent(platform, input, trackOptions), trackOptions);
    },
    participantTrack(platform, input = {}, trackOptions = {}) {
      return send(buildMeetingPlatformParticipantTrackRuntimeEvent(platform, input, trackOptions), trackOptions);
    },
    timelineView(platform, input = {}, viewOptions = {}) {
      return send(buildMeetingPlatformTimelineViewRuntimeEvent(platform, input, viewOptions), viewOptions);
    },
    runtimeBundles(bundleOptions = {}) {
      return send({ action: 'runtime_bundles' }, bundleOptions);
    },
    registry(registryOptions = {}) {
      return send({ action: 'registry' }, registryOptions);
    },
    manifest(manifestOptions = {}) {
      return send({ action: 'manifest' }, manifestOptions);
    },
    readiness(readinessOptions = {}) {
      return send({ action: 'readiness' }, readinessOptions);
    },
    handoffReadiness(readinessOptions = {}) {
      return send({ action: 'handoff_readiness' }, readinessOptions);
    },
    runManifest(manifestOptions = {}) {
      return send(buildMeetingPlatformRunManifestRuntimeEvent(manifestOptions, manifestOptions), manifestOptions);
    },
    runHandoffReadiness(readinessOptions = {}) {
      return send(buildMeetingPlatformRunHandoffReadinessRuntimeEvent(readinessOptions, readinessOptions), readinessOptions);
    },
  };
}
