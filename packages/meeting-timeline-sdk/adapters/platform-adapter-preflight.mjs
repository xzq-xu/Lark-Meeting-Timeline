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

function selectedCandidateIndex(preflights = []) {
  const accepted = preflights.findIndex((preflight) => preflight.accepted === true);
  if (accepted >= 0) return accepted;
  const staticReady = preflights.findIndex((preflight) => preflight.readiness?.static_startup_ready === true);
  if (staticReady >= 0) return staticReady;
  return preflights.findIndex((preflight) => preflight.platform);
}

function candidateRow(candidate = {}, preflight = {}, selected = false) {
  return compactObject({
    candidate_index: candidate.candidate_index,
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

export function buildMeetingPlatformAdapterCandidatePreflight(input = {}, options = {}) {
  const objectInput = normalizeInput(input);
  const candidates = flattenCandidateTabs(objectInput, options);
  const preflights = candidates.map((candidate) => candidatePreflight(candidate, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  const selectedIndex = selectedCandidateIndex(preflights);
  const selectedPreflight = selectedIndex >= 0 ? preflights[selectedIndex] : undefined;
  const rows = preflights.map((preflight, index) => candidateRow(candidates[index], preflight, index === selectedIndex));
  return compactObject({
    type: 'meeting_platform_adapter_candidate_preflight',
    schema: MEETING_PLATFORM_ADAPTER_CANDIDATE_PREFLIGHT_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_PREFLIGHT_SCHEMA_VERSION,
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
