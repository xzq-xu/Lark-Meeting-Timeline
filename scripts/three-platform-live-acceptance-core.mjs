import { normalizeMeetingPlatform } from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';

export const THREE_PLATFORM_LIVE_ACCEPTANCE_PLATFORMS = Object.freeze([
  'google_meet',
  'microsoft_teams',
  'zoom',
]);

function asArray(value) {
  if (Array.isArray(value)) return value;
  return value == null ? [] : [value];
}

function getPath(value, path) {
  let node = value;
  for (const part of path.split('.')) node = node?.[part];
  return node;
}

function firstBoolean(value, paths) {
  for (const path of paths) {
    const candidate = getPath(value, path);
    if (typeof candidate === 'boolean') return candidate;
  }
  return undefined;
}

function numeric(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function absoluteMs(value) {
  const direct = numeric(value);
  if (direct != null) return direct < 10_000_000_000 ? direct * 1000 : direct;
  const parsed = Date.parse(String(value ?? ''));
  return Number.isFinite(parsed) ? parsed : null;
}

function recordTime(record = {}) {
  return absoluteMs(record.captured_at_ms ?? record.capturedAtMs ?? record.snapshot?.observedAtMs ?? record.snapshot?.observed_at_ms) ?? 0;
}

function recordUrl(record = {}) {
  return String(
    record.snapshot?.url
      ?? record.snapshot?.page?.url
      ?? record.snapshot?.dom?.url
      ?? '',
  ).trim();
}

function officialMeetingUrl(record = {}, platform) {
  let url;
  try { url = new URL(recordUrl(record)); } catch { return false; }
  const host = url.hostname.toLowerCase();
  const path = url.pathname.toLowerCase();
  if (platform === 'google_meet') {
    return host === 'meet.google.com' && /^\/[a-z]{3}-[a-z]{4}-[a-z]{3}\/?$/.test(path);
  }
  if (platform === 'microsoft_teams') {
    return host === 'teams.microsoft.com'
      || host.endsWith('.teams.microsoft.com')
      || host === 'teams.live.com'
      || host.endsWith('.teams.live.com')
      || host === 'teams.cloud.microsoft'
      || host.endsWith('.teams.cloud.microsoft');
  }
  if (platform === 'zoom') {
    const zoomHost = host === 'zoom.us'
      || host.endsWith('.zoom.us')
      || host === 'zoom.com'
      || host.endsWith('.zoom.com');
    return zoomHost && (path.includes('/wc/') || path.includes('/j/') || path.includes('/join/'));
  }
  return false;
}

function captureProfileMatches(record = {}, platform) {
  const profile = record.snapshot?.capture?.profile;
  if (!profile) return false;
  try { return normalizeMeetingPlatform(profile) === platform; } catch { return false; }
}

function trustedCaptureProvenance(record = {}, platform) {
  const snapshot = record.snapshot ?? {};
  const source = `${record.source ?? ''} ${snapshot.source ?? ''}`.toLowerCase();
  if (!source.includes('meeting_app_extension_')) return false;
  if (snapshot.schema !== 'meeting_app_dom_capture') return false;
  const phase = String(record.phase ?? '').toLowerCase();
  if (phase === 'ended' && snapshot.lifecycle?.meeting_was_active === true) return true;
  return captureProfileMatches(record, platform);
}

function realRecord(record = {}, platform) {
  const source = `${record.source ?? ''} ${record.snapshot?.source ?? ''}`.toLowerCase();
  const title = String(record.snapshot?.title ?? '').toLowerCase();
  return Boolean(record.snapshot)
    && !/fixture|synthetic|simulation|sdk fixture/.test(`${source} ${title}`)
    && officialMeetingUrl(record, platform)
    && trustedCaptureProvenance(record, platform);
}

function semanticTypes(snapshot = {}) {
  const rows = [
    ...asArray(snapshot.semanticSignalTypes),
    ...asArray(snapshot.semantic_signal_types),
    ...asArray(snapshot.capture?.semantic_signal_types),
    ...asArray(snapshot.page?.semanticSignalTypes),
    ...asArray(snapshot.page?.semantic_signal_types),
    ...asArray(snapshot.dom?.semanticSignalTypes),
    ...asArray(snapshot.dom?.semantic_signal_types),
    ...asArray(snapshot.semanticSignals).map((item) => item?.type),
    ...asArray(snapshot.page?.semanticSignals).map((item) => item?.type),
    ...asArray(snapshot.dom?.semanticSignals).map((item) => item?.type),
  ];
  return new Set(rows.map((item) => String(item ?? '').trim()).filter(Boolean));
}

function activeParticipant(snapshot = {}) {
  const participants = [
    ...asArray(snapshot.participants),
    ...asArray(snapshot.page?.participants),
    ...asArray(snapshot.dom?.participants),
  ];
  return participants.some((item) => (
    item?.speaking === true
    || item?.isSpeaking === true
    || item?.active_speaker === true
    || item?.activeSpeaker === true
  ));
}

function desktopControlEvidence(snapshot = {}) {
  const matched = snapshot.matched_controls ?? snapshot.matchedControls ?? {};
  const hasLeave = asArray(matched.leave).length > 0;
  const hasSupport = ['microphone', 'camera', 'participants', 'share']
    .some((key) => asArray(matched[key]).length > 0);
  return hasLeave && hasSupport;
}

export function hasStrongInCallEvidence(snapshot = {}) {
  const explicit = firstBoolean(snapshot, [
    'inMeeting',
    'in_meeting',
    'page.inMeeting',
    'page.in_meeting',
    'dom.inMeeting',
    'dom.in_meeting',
    'interaction.in_call',
    'page.interaction.in_call',
    'dom.interaction.in_call',
  ]);
  const types = semanticTypes(snapshot);
  const strongSemantic = types.has('meeting_leave_available')
    || types.has('active_speaker_candidate')
    || types.has('screen_share_active')
    || activeParticipant(snapshot)
    || desktopControlEvidence(snapshot);
  return explicit === true && strongSemantic;
}

function hasStrongEndedEvidence(record = {}, activeRecord = {}) {
  const snapshot = record.snapshot ?? {};
  const explicit = firstBoolean(snapshot, [
    'inMeeting',
    'in_meeting',
    'page.inMeeting',
    'page.in_meeting',
    'dom.inMeeting',
    'dom.in_meeting',
  ]);
  const preJoin = firstBoolean(snapshot, [
    'interaction.pre_join',
    'page.interaction.pre_join',
    'dom.interaction.pre_join',
  ]) === true;
  const types = semanticTypes(snapshot);
  const joinAvailable = types.has('meeting_join_available') || types.has('meeting_waiting_room');
  const activeUrl = String(activeRecord.snapshot?.url ?? activeRecord.snapshot?.page?.url ?? '');
  const endedUrl = String(snapshot.url ?? snapshot.page?.url ?? '');
  const navigatedAway = Boolean(activeUrl && endedUrl && activeUrl !== endedUrl);
  return explicit === false
    || preJoin
    || joinAvailable
    || (navigatedAway && !hasStrongInCallEvidence(snapshot));
}

function speakerIdentity(row = {}) {
  return String(
    row.speaker_id
      ?? row.speaker_name
      ?? row.participant_id
      ?? row.participant_name
      ?? row.payload?.speaker_id
      ?? row.payload?.speaker_name
      ?? row.label
      ?? '',
  ).trim();
}

function check(id, passed, details = {}) {
  return { id, passed: Boolean(passed), ...details };
}

function rowsWithin(rows = [], startMs, endMs) {
  return rows.filter((row) => {
    const atMs = absoluteMs(row.captured_at_ms ?? row.capturedAtMs);
    return atMs != null && atMs >= startMs - 1_500 && atMs <= endMs + 1_500;
  });
}

function selectRun(records = [], meeting = {}, platform) {
  const sorted = records
    .filter((record) => realRecord(record, platform))
    .sort((left, right) => recordTime(left) - recordTime(right));
  const meetingStartMs = absoluteMs(meeting.start_time ?? meeting.start_time_ms);
  const activeCandidates = sorted.filter((record) => (
    String(record.phase ?? '').toLowerCase() === 'active'
    && hasStrongInCallEvidence(record.snapshot)
  ));
  const active = meetingStartMs == null
    ? activeCandidates.at(-1)
    : activeCandidates
      .filter((record) => Math.abs(recordTime(record) - meetingStartMs) <= 60_000)
      .sort((left, right) => Math.abs(recordTime(left) - meetingStartMs) - Math.abs(recordTime(right) - meetingStartMs))
      .at(0);
  if (!active) return { active: null, ended: null, records: sorted };
  const ended = sorted.find((record) => (
    recordTime(record) >= recordTime(active)
    && String(record.phase ?? '').toLowerCase() === 'ended'
    && hasStrongEndedEvidence(record, active)
  )) ?? null;
  return { active, ended, records: sorted };
}

export function evaluateThreePlatformLiveEvidence(input = {}, options = {}) {
  let platform = null;
  try {
    platform = normalizeMeetingPlatform(options.platform ?? input.platform);
  } catch {}
  const meeting = input.meeting ?? {};
  const run = selectRun(input.meetingAppRecords ?? input.meeting_app_records ?? [], meeting, platform);
  const activeMs = run.active ? recordTime(run.active) : null;
  const endedMs = run.ended ? recordTime(run.ended) : null;
  const rangeEndMs = endedMs ?? Number.MAX_SAFE_INTEGER;
  const annotations = activeMs == null ? [] : rowsWithin(input.annotations ?? [], activeMs, rangeEndMs);
  const speakerMarkers = activeMs == null ? [] : rowsWithin(input.speaker_markers ?? [], activeMs, rangeEndMs);
  const acceptanceAnnotations = annotations.filter((row) => (
    row.source === (options.annotationSource ?? 'three_platform_live_acceptance')
    || String(row.id ?? '').startsWith('three-platform-live-acceptance-')
  ));
  const speakerStarts = speakerMarkers.filter((row) => (
    row.kind === 'speaker_started'
    || row.intent === 'speaker_track'
  ) && speakerIdentity(row));
  const annotation = acceptanceAnnotations[0] ?? null;
  const annotationCapturedMs = absoluteMs(annotation?.captured_at_ms);
  const speakerVisibleLatencyMs = numeric(speakerStarts[0]?.visible_latency_ms);
  const eventOrder = activeMs != null
    && endedMs != null
    && activeMs <= endedMs
    && annotationCapturedMs != null
    && annotationCapturedMs >= activeMs - 1_500
    && annotationCapturedMs <= endedMs + 1_500;
  const visibleLatencyMs = numeric(annotation?.visible_latency_ms);
  const timelineErrorMs = numeric(annotation?.timeline_error_ms);
  const startObserverLatencyMs = numeric(input.measurements?.active_observer_to_axis_latency_ms);
  const endObserverLatencyMs = numeric(input.measurements?.ended_observer_to_axis_latency_ms);
  const previousMeetingAnnotationCount = numeric(input.measurements?.previous_meeting_annotation_count_on_new_axis);
  const duplicateVisible = annotation?.duplicate_visible === true || Number(annotation?.visible_instance_count ?? 1) > 1;
  const requestedSinceMs = absoluteMs(options.sinceMs ?? options.since_ms);
  const checks = [
    check('supported_platform', THREE_PLATFORM_LIVE_ACCEPTANCE_PLATFORMS.includes(platform), { platform }),
    check('meeting_identity', Boolean(input.meeting_id ?? meeting.meeting_id)),
    check('real_active_snapshot', Boolean(run.active)),
    check('real_ended_snapshot', Boolean(run.ended)),
    check('stable_speaker_marker', speakerStarts.length > 0, { speaker_count: speakerStarts.length }),
    check('speaker_marker_visible_latency', speakerVisibleLatencyMs != null && speakerVisibleLatencyMs >= 0 && speakerVisibleLatencyMs <= Number(options.maxSpeakerLatencyMs ?? 3_000), { value_ms: speakerVisibleLatencyMs }),
    check('realtime_acceptance_annotation', Boolean(annotation), { annotation_id: annotation?.id }),
    check('event_order', eventOrder, { active_at_ms: activeMs, annotation_at_ms: annotationCapturedMs, ended_at_ms: endedMs }),
    check('fresh_run', requestedSinceMs == null || (activeMs != null && activeMs >= requestedSinceMs - 10_000), { required_since_ms: requestedSinceMs }),
    check('annotation_visible_latency', visibleLatencyMs != null && visibleLatencyMs >= 0 && visibleLatencyMs <= Number(options.maxAnnotationLatencyMs ?? 3_000), { value_ms: visibleLatencyMs }),
    check('annotation_timeline_alignment', timelineErrorMs != null && Math.abs(timelineErrorMs) <= Number(options.maxTimelineErrorMs ?? 1_500), { value_ms: timelineErrorMs }),
    check('start_observer_latency', startObserverLatencyMs != null && startObserverLatencyMs >= 0 && startObserverLatencyMs <= Number(options.maxLifecycleLatencyMs ?? 3_000), { value_ms: startObserverLatencyMs }),
    check('end_observer_latency', endObserverLatencyMs != null && endObserverLatencyMs >= 0 && endObserverLatencyMs <= Number(options.maxLifecycleLatencyMs ?? 3_000), { value_ms: endObserverLatencyMs }),
    check('cross_meeting_isolation', previousMeetingAnnotationCount === 0, { previous_annotation_count: previousMeetingAnnotationCount }),
    check('no_duplicate_annotation', !duplicateVisible),
  ];
  const accepted = checks.every((item) => item.passed);
  return {
    type: 'three_platform_live_evidence_evaluation',
    schema: 'three_platform_live_evidence_evaluation',
    schema_version: 1,
    platform,
    file: options.file,
    meeting_id: input.meeting_id ?? meeting.meeting_id ?? null,
    run_id: input.run_id ?? input.runId ?? null,
    accepted,
    production_ready: accepted,
    active_at_ms: activeMs,
    annotation_at_ms: annotationCapturedMs,
    ended_at_ms: endedMs,
    speaker_identity: speakerIdentity(speakerStarts[0]),
    annotation_id: annotation?.id ?? null,
    measurements: {
      annotation_visible_latency_ms: visibleLatencyMs,
      annotation_timeline_error_ms: timelineErrorMs,
      speaker_marker_visible_latency_ms: speakerVisibleLatencyMs,
      start_observer_to_axis_latency_ms: startObserverLatencyMs,
      end_observer_to_axis_latency_ms: endObserverLatencyMs,
      previous_meeting_annotation_count_on_new_axis: previousMeetingAnnotationCount,
    },
    checks,
    failed_check_ids: checks.filter((item) => !item.passed).map((item) => item.id),
  };
}

export function buildThreePlatformLiveAcceptanceReport(evaluations = [], options = {}) {
  const requiredPlatforms = (options.platforms ?? THREE_PLATFORM_LIVE_ACCEPTANCE_PLATFORMS)
    .map((platform) => normalizeMeetingPlatform(platform));
  const rows = requiredPlatforms.map((platform) => {
    const candidates = evaluations
      .filter((row) => row.platform === platform)
      .sort((left, right) => Number(left.active_at_ms ?? 0) - Number(right.active_at_ms ?? 0));
    return candidates.at(-1) ?? {
      platform,
      accepted: false,
      production_ready: false,
      failed_check_ids: ['missing_real_meeting_evidence'],
    };
  });
  const accepted = rows.length > 0 && rows.every((row) => row.accepted === true);
  return {
    type: 'three_platform_live_acceptance_report',
    schema: 'three_platform_live_acceptance_report',
    schema_version: 1,
    generated_at: new Date().toISOString(),
    required_platforms: requiredPlatforms,
    platform_count: rows.length,
    accepted_platform_count: rows.filter((row) => row.accepted).length,
    accepted,
    production_ready: accepted,
    rows,
    remaining_platforms: rows.filter((row) => !row.accepted).map((row) => row.platform),
  };
}
