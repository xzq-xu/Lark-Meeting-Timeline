import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  buildMeetingPlatformAdapterStartupPlan,
} from './platform-adapter-startup.mjs';
import {
  captureMeetingAppDomSnapshot,
} from './meeting-app-capture.mjs';
import {
  MEETING_APP_FIXTURE_PLATFORMS,
} from './meeting-app-fixtures.mjs';
import {
  buildMeetingAppDomAdaptationDiagnosis,
} from './meeting-app-profile.mjs';
import {
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_ADAPTER_PREFLIGHT_SCHEMA = 'meeting_platform_adapter_preflight';
export const MEETING_PLATFORM_ADAPTER_PREFLIGHT_MATRIX_SCHEMA = 'meeting_platform_adapter_preflight_matrix';
export const MEETING_PLATFORM_ADAPTER_PREFLIGHT_SCHEMA_VERSION = 1;

const DEFAULT_PREFLIGHT_PLATFORMS = Object.freeze([
  'google_meet',
  'microsoft_teams',
  'zoom',
  'webex',
  'lark',
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

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function normalizeInput(input = {}) {
  if (typeof input === 'string' || input instanceof URL) return { url: String(input) };
  return input ?? {};
}

function normalizeMaybePlatform(value) {
  if (!value) return undefined;
  try {
    return normalizeMeetingPlatform(value);
  } catch {
    return String(value);
  }
}

function selectedPlatforms(input = {}, options = {}) {
  return unique(asArray(firstNonEmpty(
    options.platforms,
    options.platform_keys,
    input.platforms,
    input.platform_keys,
    DEFAULT_PREFLIGHT_PLATFORMS,
  )).map((platform) => normalizeMaybePlatform(platform)).filter(Boolean));
}

function platformAliases(platform) {
  return unique([
    platform,
    String(platform).replaceAll('_', '-'),
    platform === 'microsoft_teams' ? 'teams' : undefined,
    platform === 'google_meet' ? 'google-meet' : undefined,
  ]);
}

function inputForPlatform(platform, input = {}, options = {}) {
  const sources = [
    input.inputs,
    input.inputByPlatform,
    input.input_by_platform,
    options.inputs,
    options.inputByPlatform,
    options.input_by_platform,
  ];
  for (const source of sources) {
    if (!isPlainObject(source)) continue;
    for (const key of platformAliases(platform)) {
      if (source[key] != null) return source[key];
    }
  }
  return firstNonEmpty(input.input, input.snapshot, input.sample, options.input, options.snapshot, options.sample, {});
}

function evidenceInput(input = {}) {
  return compactObject({
    snapshots: input.snapshots,
    domSnapshots: input.domSnapshots,
    dom_snapshots: input.dom_snapshots,
    recordSet: input.recordSet,
    record_set: input.record_set,
    snapshotRecords: input.snapshotRecords,
    snapshot_records: input.snapshot_records,
    records: input.records,
    meetingAppRecordSet: input.meetingAppRecordSet,
    meeting_app_record_set: input.meeting_app_record_set,
    meetingAppRecords: input.meetingAppRecords,
    meeting_app_records: input.meeting_app_records,
  });
}

function evidenceSnapshotsForPlatform(platform, input = {}) {
  const snapshots = [];
  for (const source of [
    input.snapshots,
    input.domSnapshots,
    input.dom_snapshots,
  ]) {
    if (Array.isArray(source)) {
      snapshots.push(...source);
      continue;
    }
    if (!isPlainObject(source)) continue;
    for (const key of platformAliases(platform)) {
      if (source[key] != null) snapshots.push(...asArray(source[key]));
    }
  }
  return snapshots;
}

function preflightIssue(severity, code, message, details = {}) {
  return compactObject({
    severity,
    code,
    message,
    ...details,
  });
}

function preflightId(platform, options = {}) {
  return String(firstNonEmpty(
    options.preflightId,
    options.preflight_id,
    platform ? `${platform}-adapter-preflight` : 'meeting-platform-adapter-preflight',
  ));
}

function requireSpeakerTrack(options = {}) {
  return options.requireSpeakerTrack === true || options.require_speaker_track === true;
}

function requireCompleteLifecycle(options = {}) {
  return options.requireCompleteLifecycle === true || options.require_complete_lifecycle === true;
}

function buildDomDiagnosis(platform, input = {}, options = {}) {
  if (!MEETING_APP_FIXTURE_PLATFORMS.includes(platform)) {
    return {
      type: 'meeting_app_dom_adaptation_diagnosis',
      schema: 'meeting_app_dom_adaptation_diagnosis',
      platform,
      accepted: false,
      production_ready: false,
      evidence_count: 0,
      record_count: 0,
      selector_probe: { matched: {} },
      observer_probe: { coverage: {} },
      runtime_probe: {},
      issues: [
        preflightIssue('error', 'unsupported_live_dom_platform', 'This platform does not have a browser meeting-app DOM probe profile.', {
          supported_platforms: MEETING_APP_FIXTURE_PLATFORMS,
        }),
      ],
      next_actions: ['use_platform_adapter_startup_plan_or_provider_reconcile_for_this_platform'],
    };
  }
  return buildMeetingAppDomAdaptationDiagnosis({
    ...input,
    platform,
  }, {
    ...options,
    // A current live meeting should not need an already-ended snapshot before marks can be inserted.
    requireMeetingEnd: firstNonEmpty(options.requireMeetingEnd, options.require_meeting_end, false),
    require_meeting_end: firstNonEmpty(options.require_meeting_end, options.requireMeetingEnd, false),
  });
}

function diagnosisCoverage(diagnosis = {}) {
  return diagnosis.observer_probe?.coverage ?? {};
}

function buildReadiness(startup = {}, diagnosis = {}, options = {}) {
  const coverage = diagnosisCoverage(diagnosis);
  const speakerMatched = diagnosis.selector_probe?.matched?.active_speaker === true;
  const speakerStarted = coverage.speaker_started === true;
  const meetingStarted = coverage.meeting_started === true;
  const meetingEnded = coverage.meeting_ended === true;
  const hasLiveEvidence = (diagnosis.record_count ?? diagnosis.evidence_count ?? 0) > 0;
  const staticStartupReady = startup.realtime_startup_ready === true;
  const liveEvidenceReady = diagnosis.accepted === true && hasLiveEvidence;
  const speakerTrackReady = speakerMatched && speakerStarted;
  const speakerRequired = requireSpeakerTrack(options);
  const completeLifecycleRequired = requireCompleteLifecycle(options);
  const realtimeReady = staticStartupReady
    && liveEvidenceReady
    && meetingStarted
    && (!speakerRequired || speakerTrackReady)
    && (!completeLifecycleRequired || meetingEnded);
  return {
    static_startup_ready: staticStartupReady,
    live_evidence_ready: liveEvidenceReady,
    meeting_start_ready: meetingStarted,
    meeting_end_ready: meetingEnded,
    speaker_track_ready: speakerTrackReady,
    speaker_track_required: speakerRequired,
    complete_lifecycle_required: completeLifecycleRequired,
    provider_reconcile_required_for_realtime: startup.provider_reconcile?.required_for_realtime === true,
    transcript_blocks_realtime: startup.runtime_contract?.transcript_blocks_realtime === true,
    realtime_annotation_ready: realtimeReady,
    production_lifecycle_ready: staticStartupReady && liveEvidenceReady && meetingStarted && meetingEnded,
  };
}

function computedIssues(startup = {}, diagnosis = {}, readiness = {}) {
  const issues = [
    ...(startup.issues ?? []),
    ...(diagnosis.issues ?? []),
  ];
  if (startup.accepted !== true) {
    issues.push(preflightIssue('error', 'startup_not_accepted', 'The adapter startup plan is not accepted.'));
  }
  if (readiness.static_startup_ready !== true) {
    issues.push(preflightIssue('error', 'static_startup_not_ready', 'The platform is not ready to start a realtime local axis.'));
  }
  if ((diagnosis.record_count ?? diagnosis.evidence_count ?? 0) === 0) {
    issues.push(preflightIssue('error', 'missing_live_page_evidence', 'No live meeting page snapshots were supplied for this current-window preflight.'));
  }
  if (readiness.meeting_start_ready !== true) {
    issues.push(preflightIssue('error', 'missing_realtime_meeting_start', 'The live page evidence did not prove a meeting_started signal.'));
  }
  if (readiness.speaker_track_required && readiness.speaker_track_ready !== true) {
    issues.push(preflightIssue('error', 'missing_speaker_track_signal', 'Speaker track was required but active-speaker evidence was not stable.'));
  }
  if (readiness.complete_lifecycle_required && readiness.meeting_end_ready !== true) {
    issues.push(preflightIssue('error', 'missing_realtime_meeting_end', 'Complete lifecycle was required but live evidence did not prove a meeting_ended signal.'));
  }
  return uniqueIssues(issues);
}

function uniqueIssues(issues = []) {
  const seen = new Set();
  return issues.filter((issue) => {
    const key = `${issue.severity}:${issue.code}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function statusFor(readiness = {}, issues = []) {
  if (issues.some((issue) => issue.code === 'startup_not_accepted' || issue.code === 'static_startup_not_ready')) {
    return 'startup_not_ready';
  }
  if (issues.some((issue) => issue.code === 'missing_live_page_evidence')) {
    return 'needs_live_page_evidence';
  }
  if (readiness.realtime_annotation_ready !== true) {
    return 'live_page_partial';
  }
  if (readiness.production_lifecycle_ready === true) {
    return 'ready_for_realtime_annotations_and_lifecycle';
  }
  return 'ready_for_realtime_annotations';
}

function nextActions(startup = {}, diagnosis = {}, readiness = {}, issues = []) {
  const actions = [
    ...(startup.next_actions ?? []),
    ...(diagnosis.next_actions ?? []),
    ...issues.map((issue) => issue.code),
  ];
  if (readiness.static_startup_ready !== true) actions.push('resolve_supported_meeting_platform_and_install_local_surface');
  if (readiness.live_evidence_ready !== true) actions.push('collect_live_dom_snapshots_from_current_meeting_window');
  if (readiness.meeting_start_ready !== true) actions.push('keep_local_bridge_attached_during_join_until_meeting_started');
  if (readiness.speaker_track_required && readiness.speaker_track_ready !== true) actions.push('tune_active_speaker_selectors_or_audio_level_signal');
  if (readiness.meeting_end_ready !== true) actions.push('capture_leave_or_ended_state_to_validate_end_axis');
  if (readiness.realtime_annotation_ready === true) {
    actions.push('open_adapter_session_and_insert_marks_with_captured_at_ms');
    actions.push('keep_provider_and_transcript_reconcile_nonblocking');
  }
  return unique(actions);
}

function preflightSummary(startup = {}, diagnosis = {}, readiness = {}) {
  return {
    selected_surface: startup.selected_surface,
    install_target: startup.install_target,
    runtime_preset: startup.runtime?.preset,
    browser_match_count: startup.browser?.matches?.length ?? 0,
    content_script_match_count: startup.browser?.content_script?.matches?.length ?? 0,
    evidence_count: diagnosis.evidence_count ?? 0,
    record_count: diagnosis.record_count ?? 0,
    controls_matched: diagnosis.selector_probe?.matched?.controls === true,
    participants_matched: diagnosis.selector_probe?.matched?.participants === true,
    active_speaker_matched: diagnosis.selector_probe?.matched?.active_speaker === true,
    meeting_started: readiness.meeting_start_ready === true,
    meeting_ended: readiness.meeting_end_ready === true,
    speaker_track_ready: readiness.speaker_track_ready === true,
    transcript_blocks_realtime: readiness.transcript_blocks_realtime === true,
  };
}

function currentWindowCaptureOptions(options = {}) {
  return compactObject({
    ...options,
    ...(options.captureOptions ?? {}),
    ...(options.capture_options ?? {}),
    source: firstNonEmpty(
      options.captureSource,
      options.capture_source,
      options.source,
      'platform_adapter_current_window_preflight',
    ),
  });
}

function currentWindowInput(objectInput = {}, capturedSnapshot = {}, options = {}) {
  const platform = normalizeMaybePlatform(firstNonEmpty(
    objectInput.platform,
    objectInput.provider,
    options.platform,
    options.provider,
    capturedSnapshot.capture?.profile,
    capturedSnapshot.platform,
  ));
  const priorSnapshots = platform ? [
    ...evidenceSnapshotsForPlatform(platform, objectInput),
    ...evidenceSnapshotsForPlatform(platform, options),
  ] : [];
  return compactObject({
    platform,
    url: firstNonEmpty(objectInput.url, objectInput.href, capturedSnapshot.url, capturedSnapshot.page?.url, capturedSnapshot.dom?.url),
    title: firstNonEmpty(objectInput.title, capturedSnapshot.title, capturedSnapshot.page?.title, capturedSnapshot.dom?.title),
    snapshots: [
      ...priorSnapshots,
      capturedSnapshot,
    ],
  });
}

function capturedSnapshotSummary(capturedSnapshot = {}) {
  return {
    source: capturedSnapshot.source,
    observed_at_ms: capturedSnapshot.observedAtMs ?? capturedSnapshot.observed_at_ms,
    url: capturedSnapshot.url,
    title: capturedSnapshot.title,
    profile: capturedSnapshot.capture?.profile,
    profile_display_name: capturedSnapshot.capture?.profile_display_name,
    control_count: capturedSnapshot.capture?.control_count ?? capturedSnapshot.page?.controls?.length ?? 0,
    participant_count: capturedSnapshot.capture?.participant_count ?? capturedSnapshot.page?.participants?.length ?? 0,
    text_count: capturedSnapshot.capture?.text_count ?? capturedSnapshot.page?.texts?.length ?? 0,
    shadow_root_count: capturedSnapshot.capture?.shadow_root_count,
  };
}

function withoutEvidenceOptions(options = {}) {
  const {
    snapshots,
    domSnapshots,
    dom_snapshots,
    recordSet,
    record_set,
    snapshotRecords,
    snapshot_records,
    records,
    meetingAppRecordSet,
    meeting_app_record_set,
    meetingAppRecords,
    meeting_app_records,
    ...rest
  } = options;
  return rest;
}

export function buildMeetingPlatformAdapterPreflight(input = {}, options = {}) {
  const objectInput = normalizeInput(input);
  const startup = buildMeetingPlatformAdapterStartupPlan(objectInput, options);
  const platform = startup.platform ?? normalizeMaybePlatform(firstNonEmpty(
    objectInput.platform,
    objectInput.provider,
    options.platform,
  ));
  if (!platform) {
    const issues = computedIssues(startup, {}, { static_startup_ready: false });
    return {
      type: 'meeting_platform_adapter_preflight',
      schema: MEETING_PLATFORM_ADAPTER_PREFLIGHT_SCHEMA,
      schema_version: MEETING_PLATFORM_ADAPTER_PREFLIGHT_SCHEMA_VERSION,
      accepted: false,
      status: 'missing_platform',
      startup,
      readiness: buildReadiness(startup, {}, options),
      issues,
      next_actions: nextActions(startup, {}, {}, issues),
    };
  }
  const diagnosis = buildDomDiagnosis(platform, objectInput, options);
  const readiness = buildReadiness(startup, diagnosis, options);
  const issues = computedIssues(startup, diagnosis, readiness);
  const blockingIssues = issues.filter((issue) => issue.severity === 'error');
  return compactObject({
    type: 'meeting_platform_adapter_preflight',
    schema: MEETING_PLATFORM_ADAPTER_PREFLIGHT_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_PREFLIGHT_SCHEMA_VERSION,
    id: preflightId(platform, options),
    accepted: readiness.realtime_annotation_ready === true && blockingIssues.length === 0,
    status: statusFor(readiness, issues),
    platform,
    display_name: startup.display_name ?? diagnosis.display_name,
    input: startup.input,
    startup,
    dom_diagnosis: diagnosis,
    readiness,
    summary: preflightSummary(startup, diagnosis, readiness),
    issues,
    next_actions: nextActions(startup, diagnosis, readiness, issues),
  });
}

export function buildMeetingPlatformAdapterCurrentWindowPreflight(input = {}, options = {}) {
  const objectInput = normalizeInput(input);
  const capturedSnapshot = captureMeetingAppDomSnapshot(objectInput, currentWindowCaptureOptions(options));
  const preflight = buildMeetingPlatformAdapterPreflight(
    currentWindowInput(objectInput, capturedSnapshot, options),
    withoutEvidenceOptions(options),
  );
  return compactObject({
    ...preflight,
    current_window: {
      captured: true,
      capture_schema: capturedSnapshot.schema,
      capture_source: capturedSnapshot.source,
      capture_profile: capturedSnapshot.capture?.profile,
      observed_at_ms: capturedSnapshot.observedAtMs ?? capturedSnapshot.observed_at_ms,
      url: capturedSnapshot.url,
      title: capturedSnapshot.title,
    },
    capture: capturedSnapshotSummary(capturedSnapshot),
    captured_snapshot: options.includeCapturedSnapshot === true || options.include_captured_snapshot === true
      ? capturedSnapshot
      : undefined,
  });
}

export function buildMeetingPlatformAdapterPreflightMatrix(input = {}, options = {}) {
  const objectInput = normalizeInput(input);
  const platforms = selectedPlatforms(objectInput, options);
  const preflights = platforms.map((platform) => buildMeetingPlatformAdapterPreflight({
    ...evidenceInput(objectInput),
    ...inputForPlatform(platform, objectInput, options),
    platform,
  }, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_adapter_preflight_matrix',
    schema: MEETING_PLATFORM_ADAPTER_PREFLIGHT_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_PREFLIGHT_SCHEMA_VERSION,
    platform_count: preflights.length,
    accepted_count: preflights.filter((preflight) => preflight.accepted).length,
    realtime_ready_count: preflights.filter((preflight) => preflight.readiness?.realtime_annotation_ready === true).length,
    live_evidence_ready_count: preflights.filter((preflight) => preflight.readiness?.live_evidence_ready === true).length,
    meeting_start_ready_count: preflights.filter((preflight) => preflight.readiness?.meeting_start_ready === true).length,
    meeting_end_ready_count: preflights.filter((preflight) => preflight.readiness?.meeting_end_ready === true).length,
    speaker_track_ready_count: preflights.filter((preflight) => preflight.readiness?.speaker_track_ready === true).length,
    platforms: preflights.map((preflight) => preflight.platform).filter(Boolean),
    rows: preflights.map((preflight) => ({
      platform: preflight.platform,
      display_name: preflight.display_name,
      accepted: preflight.accepted,
      status: preflight.status,
      selected_surface: preflight.startup?.selected_surface,
      startup_ready: preflight.readiness?.static_startup_ready === true,
      live_evidence_ready: preflight.readiness?.live_evidence_ready === true,
      realtime_annotation_ready: preflight.readiness?.realtime_annotation_ready === true,
      production_lifecycle_ready: preflight.readiness?.production_lifecycle_ready === true,
      meeting_start_ready: preflight.readiness?.meeting_start_ready === true,
      meeting_end_ready: preflight.readiness?.meeting_end_ready === true,
      speaker_track_ready: preflight.readiness?.speaker_track_ready === true,
      issue_codes: (preflight.issues ?? []).map((issue) => issue.code),
      first_next_action: preflight.next_actions?.[0],
    })),
    preflights,
    next_actions: unique(preflights.flatMap((preflight) => preflight.next_actions ?? [])),
  };
}

export function assertMeetingPlatformAdapterCurrentWindowPreflight(input = {}, options = {}) {
  const preflight = buildMeetingPlatformAdapterCurrentWindowPreflight(input, options);
  if (preflight.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting platform adapter current-window preflight is not accepted', {
      platform: preflight.platform,
      status: preflight.status,
      issues: preflight.issues,
      next_actions: preflight.next_actions,
      preflight,
    });
  }
  return preflight;
}

export function assertMeetingPlatformAdapterPreflight(input = {}, options = {}) {
  const preflight = buildMeetingPlatformAdapterPreflight(input, options);
  if (preflight.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting platform adapter preflight is not accepted', {
      platform: preflight.platform,
      status: preflight.status,
      issues: preflight.issues,
      next_actions: preflight.next_actions,
      preflight,
    });
  }
  return preflight;
}

export function assertMeetingPlatformAdapterPreflightMatrix(input = {}, options = {}) {
  const matrix = buildMeetingPlatformAdapterPreflightMatrix(input, options);
  if (matrix.accepted_count !== matrix.platform_count) {
    throw new MeetingTimelineSdkError('Meeting platform adapter preflight matrix is not accepted', {
      platform_count: matrix.platform_count,
      accepted_count: matrix.accepted_count,
      failed_platforms: matrix.rows.filter((row) => row.accepted !== true).map((row) => row.platform),
      next_actions: matrix.next_actions,
      matrix,
    });
  }
  return matrix;
}
