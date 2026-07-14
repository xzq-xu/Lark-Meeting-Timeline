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
  selectNativeMeetingCandidate,
} from './native-meeting.mjs';
import {
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_ADAPTER_PREFLIGHT_SCHEMA = 'meeting_platform_adapter_preflight';
export const MEETING_PLATFORM_ADAPTER_PREFLIGHT_MATRIX_SCHEMA = 'meeting_platform_adapter_preflight_matrix';
export const MEETING_PLATFORM_ADAPTER_CANDIDATE_PREFLIGHT_SCHEMA = 'meeting_platform_adapter_candidate_preflight';
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

function getPath(raw, path) {
  const parts = path.split('.');
  let node = raw;
  for (const part of parts) node = node?.[part];
  return node;
}

function firstPath(raw, paths = []) {
  return firstNonEmpty(...paths.map((path) => getPath(raw, path)));
}

function firstBoolean(raw, paths = []) {
  for (const path of paths) {
    const value = getPath(raw, path);
    if (typeof value === 'boolean') return value;
  }
  return undefined;
}

function normalizeInput(input = {}) {
  if (typeof input === 'string' || input instanceof URL) return { url: String(input) };
  return input ?? {};
}

function explicitDocument(input = {}) {
  if (input?.querySelectorAll && input?.nodeType === 9) return input;
  if (input?.document?.querySelectorAll) return input.document;
  if (input?.window?.document?.querySelectorAll) return input.window.document;
  if (input?.tab?.document?.querySelectorAll) return input.tab.document;
  return null;
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

function signalGap(severity, code, signal, message, details = {}) {
  return compactObject({
    severity,
    code,
    signal,
    message,
    ...details,
  });
}

function controlSignalGapsFromSummary(summary, options = {}) {
  if (!looksLikeControlSignalSummary(summary)) return [];
  const speakerRequired = requireSpeakerTrack(options);
  const completeLifecycleRequired = requireCompleteLifecycle(options);
  const gaps = [
    summary.leave_available === true ? undefined : signalGap(
      'warning',
      'missing_leave_control_signal',
      'meeting_leave_available',
      'Current meeting evidence did not expose a leave/end control; end-axis capture may need provider or native fallback.',
    ),
    summary.participants_available === true && summary.participant_roster_observed === true ? undefined : signalGap(
      'info',
      'missing_participant_roster_signal',
      'participant_roster_observed',
      'Current meeting evidence did not expose a stable participant roster; participant markers may be partial.',
    ),
    summary.active_speaker_observed === true ? undefined : signalGap(
      speakerRequired ? 'error' : 'warning',
      'missing_active_speaker_signal',
      'active_speaker_candidate',
      speakerRequired
        ? 'Speaker tracking is required but current meeting evidence did not expose an active-speaker signal.'
        : 'Current meeting evidence did not expose an active-speaker signal; speaker markers may need audio/native fallback.',
    ),
    completeLifecycleRequired && summary.leave_available !== true ? signalGap(
      'error',
      'missing_lifecycle_end_control_signal',
      'meeting_leave_available',
      'Complete lifecycle validation is required but current evidence does not expose a leave/end control.',
    ) : undefined,
  ].filter(Boolean);
  return uniqueSignalGaps(gaps);
}

function controlSignalGapSummary(gaps = []) {
  const normalized = asArray(gaps).filter(Boolean);
  if (!normalized.length) {
    return {
      gap_count: 0,
      blocking_count: 0,
      warning_count: 0,
      info_count: 0,
      codes: [],
    };
  }
  return {
    gap_count: normalized.length,
    blocking_count: normalized.filter((gap) => gap.severity === 'error').length,
    warning_count: normalized.filter((gap) => gap.severity === 'warning').length,
    info_count: normalized.filter((gap) => gap.severity === 'info').length,
    codes: unique(normalized.map((gap) => gap.code)),
  };
}

function uniqueSignalGaps(gaps = []) {
  const seen = new Set();
  return gaps.filter((gap) => {
    const key = `${gap.severity}:${gap.code}:${gap.signal}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function looksLikeControlSignalSummary(value = {}) {
  if (!isPlainObject(value)) return false;
  return [
    'signal_types',
    'signal_count',
    'join_available',
    'waiting_room',
    'leave_available',
    'microphone_available',
    'camera_available',
    'screen_share_available',
    'screen_share_active',
    'captions_available',
    'recording_observed',
    'participants_available',
    'chat_available',
    'ai_summary_available',
    'active_speaker_observed',
    'participant_roster_observed',
  ].some((key) => value[key] != null);
}

export function buildMeetingPlatformAdapterControlSignalGaps(input = {}, options = {}) {
  const candidate = firstNonEmpty(
    input.control_signal_summary,
    input.controlSignalSummary,
    input.capture?.control_signal_summary,
    input.capture?.controlSignalSummary,
    input.current_window?.control_signal_summary,
    input.currentWindow?.controlSignalSummary,
    input.page?.control_signal_summary,
    input.page?.controlSignalSummary,
    input.dom?.control_signal_summary,
    input.dom?.controlSignalSummary,
    looksLikeControlSignalSummary(input) ? input : undefined,
  );
  const summary = looksLikeControlSignalSummary(candidate) ? candidate : undefined;
  const gaps = controlSignalGapsFromSummary(summary, options);
  return compactObject({
    type: 'meeting_platform_adapter_control_signal_gaps',
    schema: 'meeting_platform_adapter_control_signal_gaps',
    gap_count: gaps.length,
    summary: controlSignalGapSummary(gaps),
    gaps,
  });
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

function nativeInputForPlatform(platform, input = {}, options = {}) {
  return compactObject({
    ...input,
    platform,
    preferredPlatform: platform,
    preferred_platform: platform,
    observedAtMs: firstNonEmpty(input.observedAtMs, input.observed_at_ms, options.observedAtMs, options.observed_at_ms),
  });
}

function nativeMeetingActive(candidate = {}) {
  return firstBoolean(candidate, [
    'in_meeting',
    'inMeeting',
    'meeting.active',
    'window.in_meeting',
    'window.inMeeting',
    'window.active',
    'active',
    'focused',
    'selected',
    'accessibility.call_active',
    'accessibility.callActive',
    'audio.call_active',
    'audio.callActive',
    'call.active',
    'callActive',
    'call_active',
  ]) === true;
}

function nativeMeetingEnded(candidate = {}) {
  return firstBoolean(candidate, [
    'in_meeting',
    'inMeeting',
    'meeting.active',
    'window.in_meeting',
    'window.inMeeting',
    'accessibility.call_active',
    'accessibility.callActive',
    'audio.call_active',
    'audio.callActive',
    'call.active',
    'callActive',
    'call_active',
  ]) === false;
}

function nativeSpeakerReady(candidate = {}, input = {}) {
  return Boolean(firstPath(candidate, [
    'activeSpeaker',
    'active_speaker',
    'speaker',
    'window.activeSpeaker',
    'window.active_speaker',
    'accessibility.activeSpeaker',
    'accessibility.active_speaker',
    'audio.activeSpeaker',
    'audio.active_speaker',
  ]) ?? firstPath(input, [
    'activeSpeaker',
    'active_speaker',
    'speaker',
    'accessibility.activeSpeaker',
    'accessibility.active_speaker',
    'audio.activeSpeaker',
    'audio.active_speaker',
  ]));
}

function buildNativeDiagnosis(platform, input = {}, options = {}) {
  const nativeInput = nativeInputForPlatform(platform, input, options);
  const selection = selectNativeMeetingCandidate(nativeInput, {
    ...options,
    preferredPlatform: platform,
    preferred_platform: platform,
  });
  const selected = selection.selectedSnapshot;
  const candidateCount = selection.normalizedCandidates?.length ?? 0;
  const platformMatched = selected?.meeting?.platform === platform || selected?.platform === platform;
  const meetingStarted = Boolean(selected) && platformMatched && nativeMeetingActive(selected);
  const meetingEnded = Boolean(selected) && nativeMeetingEnded(selected);
  const speakerStarted = nativeSpeakerReady(selected ?? {}, nativeInput);
  const accepted = candidateCount > 0 && meetingStarted;
  const issues = [
    candidateCount > 0 ? undefined : preflightIssue('error', 'missing_native_candidate_evidence', 'No native meeting window/process candidate was supplied for this preflight.'),
    selected && !platformMatched ? preflightIssue('error', 'native_candidate_platform_mismatch', 'Native meeting candidate resolved to a different platform.', {
      expected_platform: platform,
      actual_platform: selected?.meeting?.platform ?? selected?.platform,
    }) : undefined,
    candidateCount > 0 && !meetingStarted ? preflightIssue('error', 'missing_native_meeting_start', 'Native evidence did not prove an active meeting window or call state.') : undefined,
  ].filter(Boolean);
  return compactObject({
    type: 'meeting_platform_native_evidence_diagnosis',
    schema: 'meeting_platform_native_evidence_diagnosis',
    platform,
    accepted,
    production_ready: false,
    evidence_count: candidateCount,
    record_count: candidateCount,
    selected_candidate: selected,
    selected_meeting: selection.detectedMeeting,
    normalized_candidate_count: candidateCount,
    selector_probe: {
      matched: {
        controls: meetingStarted,
        participants: Boolean(firstPath(selected ?? nativeInput, ['participants', 'meeting.participants', 'accessibility.participants'])),
        active_speaker: speakerStarted,
      },
    },
    observer_probe: {
      coverage: {
        meeting_started: meetingStarted,
        meeting_ended: meetingEnded,
        speaker_started: speakerStarted,
      },
    },
    runtime_probe: {
      native_detector: true,
      selected_process_name: selected?.processName ?? selected?.discovery?.process_name,
      selected_bundle_id: selected?.bundleId ?? selected?.discovery?.bundle_id,
    },
    issues,
    next_actions: accepted
      ? ['open_adapter_session_and_insert_marks_with_captured_at_ms']
      : ['collect_native_window_process_or_accessibility_sample'],
  });
}

function buildEvidenceDiagnosis(platform, input = {}, startup = {}, options = {}) {
  if (startup.selected_surface === 'native_detector') return buildNativeDiagnosis(platform, input, options);
  return buildDomDiagnosis(platform, input, options);
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
    issues.push(preflightIssue('error', 'missing_live_page_evidence', 'No live meeting page or native-window evidence was supplied for this preflight.'));
  }
  if (readiness.meeting_start_ready !== true) {
    issues.push(preflightIssue('error', 'missing_realtime_meeting_start', 'The live evidence did not prove a meeting_started signal.'));
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
  if (startup.selected_surface === 'native_detector' && readiness.live_evidence_ready !== true) actions.push('collect_native_window_process_or_accessibility_sample');
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
    evidence_kind: startup.selected_surface === 'native_detector' ? 'native_window' : 'dom_snapshot',
    browser_match_count: startup.browser?.matches?.length ?? 0,
    content_script_match_count: startup.browser?.content_script?.matches?.length ?? 0,
    evidence_count: diagnosis.evidence_count ?? 0,
    record_count: diagnosis.record_count ?? 0,
    native_candidate_count: diagnosis.normalized_candidate_count,
    native_selected_process: diagnosis.runtime_probe?.selected_process_name,
    native_selected_bundle_id: diagnosis.runtime_probe?.selected_bundle_id,
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
  const semanticSignals = [
    ...(capturedSnapshot.semanticSignals ?? []),
    ...(capturedSnapshot.page?.semanticSignals ?? []),
    ...(capturedSnapshot.dom?.semanticSignals ?? []),
  ];
  const semanticSignalTypes = unique([
    ...(capturedSnapshot.semanticSignalTypes ?? []),
    ...(capturedSnapshot.semantic_signal_types ?? []),
    ...(capturedSnapshot.page?.semanticSignalTypes ?? []),
    ...(capturedSnapshot.page?.semantic_signal_types ?? []),
    ...(capturedSnapshot.dom?.semanticSignalTypes ?? []),
    ...(capturedSnapshot.dom?.semantic_signal_types ?? []),
    ...(capturedSnapshot.capture?.semantic_signal_types ?? []),
    ...semanticSignals.map((signal) => signal?.type),
  ]);
  const interaction = firstNonEmpty(
    capturedSnapshot.interaction,
    capturedSnapshot.page?.interaction,
    capturedSnapshot.dom?.interaction,
  );
  const controlSignalSummary = firstNonEmpty(
    capturedSnapshot.controlSignalSummary,
    capturedSnapshot.control_signal_summary,
    capturedSnapshot.page?.controlSignalSummary,
    capturedSnapshot.page?.control_signal_summary,
    capturedSnapshot.dom?.controlSignalSummary,
    capturedSnapshot.dom?.control_signal_summary,
    capturedSnapshot.capture?.control_signal_summary,
  );
  const activeSpeaker = firstNonEmpty(
    capturedSnapshot.activeSpeaker,
    capturedSnapshot.active_speaker,
    interaction?.active_speaker_candidate,
    capturedSnapshot.page?.activeSpeaker,
    capturedSnapshot.dom?.activeSpeaker,
  );
  return {
    source: capturedSnapshot.source,
    observed_at_ms: capturedSnapshot.observedAtMs ?? capturedSnapshot.observed_at_ms,
    url: capturedSnapshot.url,
    title: capturedSnapshot.title,
    in_meeting: firstNonEmpty(capturedSnapshot.inMeeting, capturedSnapshot.in_meeting, capturedSnapshot.page?.inMeeting, capturedSnapshot.dom?.inMeeting),
    profile: capturedSnapshot.capture?.profile,
    profile_display_name: capturedSnapshot.capture?.profile_display_name,
    control_count: capturedSnapshot.capture?.control_count ?? capturedSnapshot.page?.controls?.length ?? 0,
    participant_count: capturedSnapshot.capture?.participant_count ?? capturedSnapshot.page?.participants?.length ?? 0,
    text_count: capturedSnapshot.capture?.text_count ?? capturedSnapshot.page?.texts?.length ?? 0,
    shadow_root_count: capturedSnapshot.capture?.shadow_root_count,
    semantic_signal_count: firstNonEmpty(capturedSnapshot.capture?.semantic_signal_count, semanticSignalTypes.length || undefined),
    semantic_signal_types: semanticSignalTypes.length ? semanticSignalTypes : undefined,
    control_signal_summary: controlSignalSummary,
    interaction,
    active_speaker_candidate: activeSpeaker ? compactObject({
      id: activeSpeaker.id,
      name: activeSpeaker.name ?? activeSpeaker.display_name,
      display_name: activeSpeaker.display_name ?? activeSpeaker.name,
      speaking: activeSpeaker.speaking ?? true,
      audioLevel: activeSpeaker.audioLevel ?? activeSpeaker.audio_level,
    }) : undefined,
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

function candidateArrays(input = {}, options = {}) {
  if (Array.isArray(input)) return [input];
  return [
    input.candidates,
    input.candidateTabs,
    input.candidate_tabs,
    input.tabs,
    input.items,
    input.rows,
    options.candidates,
    options.candidateTabs,
    options.candidate_tabs,
    options.tabs,
  ].filter((value) => value != null);
}

function candidateWindows(input = {}, options = {}) {
  if (Array.isArray(input)) return [];
  return [
    input.windows,
    input.browserWindows,
    input.browser_windows,
    input.window,
    options.windows,
    options.browserWindows,
    options.browser_windows,
  ].filter((value) => value != null).flatMap((value) => asArray(value));
}

function candidateUrl(candidate = {}) {
  if (typeof candidate === 'string' || candidate instanceof URL) return String(candidate);
  return firstNonEmpty(
    candidate.url,
    candidate.href,
    candidate.meeting_url,
    candidate.meetingUrl,
    candidate.join_url,
    candidate.joinUrl,
    candidate.location?.href,
    candidate.document?.location?.href,
    candidate.tab?.url,
    candidate.tab?.document?.location?.href,
    candidate.window?.url,
    candidate.window?.document?.location?.href,
    candidate.page?.url,
    candidate.dom?.url,
  );
}

function candidateTitle(candidate = {}) {
  if (typeof candidate === 'string' || candidate instanceof URL) return undefined;
  return firstNonEmpty(
    candidate.title,
    candidate.meeting_title,
    candidate.meetingTitle,
    candidate.document?.title,
    candidate.tab?.title,
    candidate.tab?.document?.title,
    candidate.window?.title,
    candidate.window?.document?.title,
    candidate.page?.title,
    candidate.dom?.title,
  );
}

function normalizeCandidate(candidate, meta = {}) {
  if (typeof candidate === 'string' || candidate instanceof URL) {
    return compactObject({
      candidate_index: meta.index,
      source: meta.source,
      url: String(candidate),
    });
  }
  const tab = candidate?.tab ?? candidate;
  return compactObject({
    ...candidate,
    candidate_index: meta.index,
    source: firstNonEmpty(candidate?.source, meta.source),
    window_id: firstNonEmpty(candidate?.window_id, candidate?.windowId, meta.window_id, meta.windowId),
    tab_id: firstNonEmpty(candidate?.tab_id, candidate?.tabId, candidate?.id, tab?.id),
    active: firstNonEmpty(candidate?.active, tab?.active),
    url: candidateUrl(candidate),
    title: candidateTitle(candidate),
  });
}

function flattenCandidateTabs(input = {}, options = {}) {
  const candidates = [];
  for (const source of candidateArrays(input, options)) {
    for (const candidate of asArray(source)) {
      candidates.push(normalizeCandidate(candidate, {
        index: candidates.length,
        source: 'candidate',
      }));
    }
  }
  const nativeContext = nativeCandidateContext(input);
  for (const [windowIndex, windowItem] of candidateWindows(input, options).entries()) {
    const windowCandidate = compactObject({
      ...nativeContext,
      ...windowItem,
    });
    const windowId = firstNonEmpty(windowItem?.id, windowItem?.window_id, windowItem?.windowId, windowIndex);
    const tabs = asArray(firstNonEmpty(windowItem?.tabs, windowItem?.candidateTabs, windowItem?.candidate_tabs));
    if (tabs.length === 0 && (candidateUrl(windowCandidate) || candidateHasNativeEvidence(windowCandidate))) {
      candidates.push(normalizeCandidate(windowCandidate, {
        index: candidates.length,
        source: 'window',
        window_id: windowId,
      }));
      continue;
    }
    for (const tab of tabs) {
      candidates.push(normalizeCandidate({
        ...nativeContext,
        ...tab,
        tab,
      }, {
        index: candidates.length,
        source: 'window_tab',
        window_id: windowId,
      }));
    }
  }
  return candidates;
}

function nativeCandidateContext(input = {}) {
  return compactObject({
    platform: input.platform,
    provider: input.provider,
    process: input.process,
    app: input.app,
    application: input.application,
    accessibility: input.accessibility,
    audio: input.audio,
    native: input.native,
    native_detector: input.native_detector,
  });
}

function candidateHasNativeEvidence(candidate = {}) {
  return Boolean(firstNonEmpty(
    candidate.platform,
    candidate.provider,
    candidate.process,
    candidate.app,
    candidate.application,
    candidate.accessibility,
    candidate.audio,
    candidate.native,
    candidate.native_detector,
    candidate.processName,
    candidate.process_name,
    candidate.bundleId,
    candidate.bundle_id,
    candidate.window?.ownerName,
    candidate.window?.appName,
    candidate.window?.app_name,
  ));
}

function candidatePreflightInput(candidate = {}) {
  const snapshot = firstNonEmpty(candidate.snapshot, candidate.captured_snapshot, candidate.capturedSnapshot);
  const snapshots = [
    ...asArray(candidate.snapshots),
    ...asArray(candidate.domSnapshots),
    ...asArray(candidate.dom_snapshots),
    ...asArray(snapshot),
  ];
  return compactObject({
    ...candidate,
    platform: normalizeMaybePlatform(firstNonEmpty(candidate.platform, candidate.provider)),
    url: candidateUrl(candidate),
    title: candidateTitle(candidate),
    snapshots: snapshots.length ? snapshots : undefined,
  });
}

function candidatePreflight(candidate = {}, options = {}) {
  const input = candidatePreflightInput(candidate);
  if (explicitDocument(candidate)) {
    return buildMeetingPlatformAdapterCurrentWindowPreflight(input, options);
  }
  return buildMeetingPlatformAdapterPreflight(input, options);
}

function candidateStatus(preflights = []) {
  if (preflights.length === 0) return 'empty_candidates';
  if (preflights.some((preflight) => preflight.accepted === true)) return 'ready_for_realtime_annotations';
  if (preflights.some((preflight) => preflight.readiness?.static_startup_ready === true)) return 'needs_live_page_evidence';
  if (preflights.some((preflight) => preflight.platform)) return 'startup_not_ready';
  return 'no_supported_candidates';
}

function candidateActivityScore(candidate = {}) {
  return [
    firstBoolean(candidate, ['active', 'tab.active', 'window.active']) === true ? 35 : 0,
    firstBoolean(candidate, ['focused', 'selected', 'current', 'tab.highlighted', 'tab.selected', 'window.focused']) === true ? 25 : 0,
    firstBoolean(candidate, ['audible', 'tab.audible', 'audio.call_active', 'audio.callActive']) === true ? 10 : 0,
  ].reduce((sum, value) => sum + value, 0);
}

function preflightInteractionScore(preflight = {}) {
  const interaction = preflight.capture?.interaction ?? preflight.current_window?.interaction;
  const semanticTypes = new Set(preflight.capture?.semantic_signal_types ?? preflight.current_window?.semantic_signal_types ?? []);
  return [
    interaction?.in_call === true ? 90 : 0,
    interaction?.can_leave === true ? 70 : 0,
    interaction?.active_speaker_candidate ? 55 : 0,
    interaction?.participant_roster_observed === true ? 20 : 0,
    semanticTypes.has('meeting_leave_available') ? 35 : 0,
    semanticTypes.has('active_speaker_candidate') ? 35 : 0,
    semanticTypes.has('meeting_join_available') && interaction?.in_call !== true ? -30 : 0,
  ].reduce((sum, value) => sum + value, 0);
}

function candidatePreflightScore(candidate = {}, preflight = {}) {
  if (!preflight.platform) return -1_000 + candidateActivityScore(candidate);
  return [
    preflight.accepted === true ? 1_000 : 0,
    preflight.readiness?.realtime_annotation_ready === true ? 700 : 0,
    preflight.readiness?.live_evidence_ready === true ? 280 : 0,
    preflight.readiness?.meeting_start_ready === true ? 220 : 0,
    preflight.current_window?.captured === true ? 120 : 0,
    preflight.readiness?.speaker_track_ready === true ? 70 : 0,
    preflight.readiness?.static_startup_ready === true ? 45 : 0,
    preflightInteractionScore(preflight),
    20,
    candidateActivityScore(candidate),
  ].reduce((sum, value) => sum + value, 0);
}

function candidateSelectionReason(candidate = {}, preflight = {}) {
  if (!preflight.platform) return 'unsupported_candidate';
  if (preflight.accepted === true && preflight.current_window?.captured === true) return 'accepted_current_window_live_candidate';
  if (preflight.accepted === true) return 'accepted_live_candidate';
  if (preflight.readiness?.live_evidence_ready === true) return 'live_candidate_needs_more_signals';
  if (preflight.readiness?.static_startup_ready === true && candidateActivityScore(candidate) > 0) return 'active_supported_candidate_needs_live_evidence';
  if (preflight.readiness?.static_startup_ready === true) return 'supported_candidate_needs_live_evidence';
  return 'startup_not_ready';
}

function rankedCandidateEntries(candidates = [], preflights = []) {
  return candidates
    .map((candidate, index) => ({
      index,
      candidate,
      preflight: preflights[index] ?? {},
      score: candidatePreflightScore(candidate, preflights[index] ?? {}),
      reason: candidateSelectionReason(candidate, preflights[index] ?? {}),
    }))
    .filter((entry) => entry.preflight.platform)
    .sort((a, b) => b.score - a.score || a.index - b.index)
    .map((entry, rankIndex) => ({
      ...entry,
      rank: rankIndex + 1,
    }));
}

function candidateRow(candidate = {}, preflight = {}, selected = false, selection = {}) {
  const controlSignalGaps = preflight.capture?.control_signal_gaps ?? preflight.current_window?.control_signal_gaps;
  const controlSignalGapSummary = preflight.capture?.control_signal_gap_summary ?? preflight.current_window?.control_signal_gap_summary;
  return compactObject({
    candidate_index: candidate.candidate_index,
    selection_rank: selection.rank,
    selection_score: selection.score,
    selection_reason: selection.reason,
    selected,
    source: candidate.source,
    window_id: candidate.window_id,
    tab_id: candidate.tab_id,
    active: candidate.active,
    url: candidate.url,
    title: candidate.title,
    platform: preflight.platform,
    display_name: preflight.display_name,
    selected_surface: preflight.summary?.selected_surface ?? preflight.startup?.selected_surface,
    accepted: preflight.accepted,
    status: preflight.status,
    startup_ready: preflight.readiness?.static_startup_ready === true,
    live_evidence_ready: preflight.readiness?.live_evidence_ready === true,
    realtime_annotation_ready: preflight.readiness?.realtime_annotation_ready === true,
    production_lifecycle_ready: preflight.readiness?.production_lifecycle_ready === true,
    meeting_start_ready: preflight.readiness?.meeting_start_ready === true,
    meeting_end_ready: preflight.readiness?.meeting_end_ready === true,
    speaker_track_ready: preflight.readiness?.speaker_track_ready === true,
    current_window_captured: preflight.current_window?.captured === true,
    interaction_in_call: preflight.capture?.interaction?.in_call ?? preflight.current_window?.interaction?.in_call,
    interaction_can_leave: preflight.capture?.interaction?.can_leave ?? preflight.current_window?.interaction?.can_leave,
    interaction_pre_join: preflight.capture?.interaction?.pre_join ?? preflight.current_window?.interaction?.pre_join,
    semantic_signal_types: preflight.capture?.semantic_signal_types ?? preflight.current_window?.semantic_signal_types,
    control_signal_summary: preflight.capture?.control_signal_summary ?? preflight.current_window?.control_signal_summary,
    control_signal_gaps: controlSignalGaps,
    control_signal_gap_summary: controlSignalGapSummary,
    control_signal_gap_count: controlSignalGapSummary?.gap_count,
    control_signal_gap_codes: controlSignalGapSummary?.codes,
    active_speaker_candidate_id: preflight.capture?.active_speaker_candidate?.id ?? preflight.current_window?.active_speaker_candidate?.id,
    active_speaker_candidate_name: preflight.capture?.active_speaker_candidate?.name ?? preflight.current_window?.active_speaker_candidate?.name,
    issue_codes: (preflight.issues ?? []).map((issue) => issue.code),
    first_next_action: preflight.next_actions?.[0],
  });
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
  const diagnosis = buildEvidenceDiagnosis(platform, objectInput, startup, options);
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
    dom_diagnosis: startup.selected_surface === 'native_detector' ? undefined : diagnosis,
    native_diagnosis: startup.selected_surface === 'native_detector' ? diagnosis : undefined,
    readiness,
    summary: preflightSummary(startup, diagnosis, readiness),
    issues,
    next_actions: nextActions(startup, diagnosis, readiness, issues),
  });
}

export function buildMeetingPlatformAdapterCurrentWindowPreflight(input = {}, options = {}) {
  const objectInput = normalizeInput(input);
  const capturedSnapshot = captureMeetingAppDomSnapshot(objectInput, currentWindowCaptureOptions(options));
  const captureSummary = capturedSnapshotSummary(capturedSnapshot);
  const controlSignalGaps = controlSignalGapsFromSummary(captureSummary.control_signal_summary, options);
  const controlSignalGapSummaryValue = controlSignalGapSummary(controlSignalGaps);
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
      in_meeting: captureSummary.in_meeting,
      interaction: captureSummary.interaction,
      semantic_signal_types: captureSummary.semantic_signal_types,
      control_signal_summary: captureSummary.control_signal_summary,
      control_signal_gaps: controlSignalGaps.length ? controlSignalGaps : undefined,
      control_signal_gap_summary: controlSignalGapSummaryValue,
      active_speaker_candidate: captureSummary.active_speaker_candidate,
    },
    summary: {
      ...(preflight.summary ?? {}),
      current_window_in_meeting: captureSummary.in_meeting,
      current_window_interaction: captureSummary.interaction,
      current_window_semantic_signal_types: captureSummary.semantic_signal_types,
      current_window_control_signal_summary: captureSummary.control_signal_summary,
      current_window_control_signal_gaps: controlSignalGaps.length ? controlSignalGaps : undefined,
      current_window_control_signal_gap_summary: controlSignalGapSummaryValue,
      current_window_active_speaker_candidate: captureSummary.active_speaker_candidate,
    },
    capture: {
      ...captureSummary,
      control_signal_gaps: controlSignalGaps.length ? controlSignalGaps : undefined,
      control_signal_gap_summary: controlSignalGapSummaryValue,
    },
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
      control_signal_gap_count: preflight.summary?.current_window_control_signal_gap_summary?.gap_count,
      control_signal_gap_codes: preflight.summary?.current_window_control_signal_gap_summary?.codes,
      issue_codes: (preflight.issues ?? []).map((issue) => issue.code),
      first_next_action: preflight.next_actions?.[0],
    })),
    preflights,
    next_actions: unique(preflights.flatMap((preflight) => preflight.next_actions ?? [])),
  };
}

export function buildMeetingPlatformAdapterCandidatePreflight(input = {}, options = {}) {
  const objectInput = normalizeInput(input);
  const candidates = flattenCandidateTabs(objectInput, options);
  const preflights = candidates.map((candidate) => candidatePreflight(candidate, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  const ranked = rankedCandidateEntries(candidates, preflights);
  const selectedIndex = ranked[0]?.index ?? -1;
  const selectedPreflight = selectedIndex >= 0 ? preflights[selectedIndex] : undefined;
  const rankByIndex = new Map(ranked.map((entry) => [entry.index, entry]));
  const rows = preflights.map((preflight, index) => candidateRow(
    candidates[index],
    preflight,
    index === selectedIndex,
    rankByIndex.get(index),
  ));
  return compactObject({
    type: 'meeting_platform_adapter_candidate_preflight',
    schema: MEETING_PLATFORM_ADAPTER_CANDIDATE_PREFLIGHT_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_PREFLIGHT_SCHEMA_VERSION,
    selection_strategy: 'score_accepted_live_current_window_then_active_candidate',
    accepted: preflights.some((preflight) => preflight.accepted === true),
    status: candidateStatus(preflights),
    candidate_count: candidates.length,
    supported_candidate_count: preflights.filter((preflight) => preflight.platform).length,
    accepted_count: preflights.filter((preflight) => preflight.accepted).length,
    realtime_ready_count: preflights.filter((preflight) => preflight.readiness?.realtime_annotation_ready === true).length,
    live_evidence_ready_count: preflights.filter((preflight) => preflight.readiness?.live_evidence_ready === true).length,
    meeting_start_ready_count: preflights.filter((preflight) => preflight.readiness?.meeting_start_ready === true).length,
    meeting_end_ready_count: preflights.filter((preflight) => preflight.readiness?.meeting_end_ready === true).length,
    speaker_track_ready_count: preflights.filter((preflight) => preflight.readiness?.speaker_track_ready === true).length,
    selected_candidate_index: selectedIndex >= 0 ? selectedIndex : undefined,
    selected_candidate_score: ranked[0]?.score,
    selected_platform: selectedPreflight?.platform,
    selected_status: selectedPreflight?.status,
    platforms: unique(preflights.map((preflight) => preflight.platform).filter(Boolean)),
    rows,
    preflights,
    next_actions: unique(preflights.flatMap((preflight) => preflight.next_actions ?? [])),
  });
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

export function assertMeetingPlatformAdapterCandidatePreflight(input = {}, options = {}) {
  const preflight = buildMeetingPlatformAdapterCandidatePreflight(input, options);
  if (preflight.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting platform adapter candidate preflight has no accepted realtime candidate', {
      status: preflight.status,
      candidate_count: preflight.candidate_count,
      rows: preflight.rows,
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
