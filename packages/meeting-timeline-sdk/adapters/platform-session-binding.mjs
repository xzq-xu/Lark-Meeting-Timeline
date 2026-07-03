import {
  buildMeetingStartPayload,
  compactObject,
  normalizeAbsoluteMs,
} from '../index.mjs';
import { detectMeetingFromUrl, stableMeetingIdFromUrl } from './meeting-url.mjs';
import {
  MEETING_PLATFORM_KEYS,
  buildPlatformIntegrationPlan,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_SESSION_BINDING_PLAN_SCHEMA = 'meeting_platform_session_binding_plan';
export const MEETING_PLATFORM_SESSION_BINDING_MATRIX_SCHEMA = 'meeting_platform_session_binding_matrix';
export const MEETING_PLATFORM_SESSION_BINDING_SCHEMA = 'meeting_platform_session_binding';
export const MEETING_PLATFORM_SESSION_BINDING_SCHEMA_VERSION = 1;

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

function numberOption(options, keys, fallback) {
  const value = firstNonEmpty(...keys.map((key) => options[key]));
  const parsed = Number(value ?? fallback);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function selectedPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_PLATFORM_KEYS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

function getPath(raw, path) {
  const parts = path.split('.');
  let node = raw;
  for (const part of parts) node = node?.[part];
  return node;
}

function firstPath(raw, paths) {
  return firstNonEmpty(...paths.map((path) => getPath(raw, path)));
}

function maybeAbsoluteMs(value, fieldName) {
  if (value == null || value === '') return undefined;
  try {
    return normalizeAbsoluteMs(value, fieldName);
  } catch {
    return undefined;
  }
}

function cleanId(value) {
  return String(value ?? '')
    .trim()
    .replace(/\s+/g, '-')
    .replace(/[^a-zA-Z0-9._:@~-]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .toLowerCase();
}

function normalizedText(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function urlValue(input = {}) {
  return firstPath(input, [
    'meeting.meeting_url',
    'meeting.meetingUrl',
    'meeting.url',
    'meeting.join_url',
    'meeting.joinUrl',
    'meeting_url',
    'meetingUrl',
    'join_url',
    'joinUrl',
    'url',
    'href',
    'window.url',
    'browser.url',
    'tab.url',
  ]);
}

function titleValue(input = {}) {
  return firstPath(input, [
    'meeting.title',
    'meeting.topic',
    'meeting.name',
    'title',
    'topic',
    'name',
    'window.title',
    'browser.title',
    'tab.title',
  ]);
}

function platformValue(input = {}, fallback) {
  const value = firstPath(input, [
    'meeting.platform',
    'platform',
    'provider',
    'adapter',
  ]);
  try {
    return normalizeMeetingPlatform(firstNonEmpty(value, fallback));
  } catch {
    return normalizeMeetingPlatform(fallback);
  }
}

function meetingIdValue(input = {}) {
  return firstPath(input, [
    'meeting.meeting_id',
    'meeting.meetingId',
    'meeting.id',
    'meeting_id',
    'meetingId',
    'session_id',
    'sessionId',
  ]);
}

function externalMeetingIdValue(input = {}) {
  return firstPath(input, [
    'meeting.external_meeting_id',
    'meeting.externalMeetingId',
    'meeting.meeting_no',
    'meeting.meetingNo',
    'external_meeting_id',
    'externalMeetingId',
    'meeting_no',
    'meetingNo',
    'conference_record',
    'conferenceRecord',
  ]);
}

function startTimeValue(input = {}) {
  return maybeAbsoluteMs(firstPath(input, [
    'meeting.start_time_ms',
    'meeting.startTimeMs',
    'meeting.started_at_ms',
    'meeting.startedAtMs',
    'meeting.start_time',
    'meeting.startTime',
    'meeting.started_at',
    'meeting.startedAt',
    'start_time_ms',
    'startTimeMs',
    'started_at_ms',
    'startedAtMs',
    'start_time',
    'startTime',
    'started_at',
    'startedAt',
    'occurred_at_ms',
    'occurredAtMs',
    'observed_at_ms',
    'observedAtMs',
    'captured_at_ms',
    'capturedAtMs',
  ]), 'meeting.start_time_ms');
}

function endTimeValue(input = {}) {
  return maybeAbsoluteMs(firstPath(input, [
    'meeting.end_time_ms',
    'meeting.endTimeMs',
    'meeting.ended_at_ms',
    'meeting.endedAtMs',
    'meeting.end_time',
    'meeting.endTime',
    'meeting.ended_at',
    'meeting.endedAt',
    'end_time_ms',
    'endTimeMs',
    'ended_at_ms',
    'endedAtMs',
    'end_time',
    'endTime',
    'ended_at',
    'endedAt',
  ]), 'meeting.end_time_ms');
}

function sourceRank(source) {
  return {
    current_axis: 100,
    local_observer: 90,
    browser_observer: 90,
    native_observer: 88,
    provider_signal: 70,
    provider_event: 70,
    annotation: 40,
    unknown: 10,
  }[source] ?? 10;
}

function normalizeCandidate(source, input = {}, platform, options = {}) {
  if (!input || typeof input !== 'object') return undefined;
  const detected = detectMeetingFromUrl(input) ?? {};
  const rawUrl = firstNonEmpty(urlValue(input), detected.meeting_url);
  const key = platformValue({
    ...input,
    platform: firstNonEmpty(input.platform, input.provider, detected.platform),
  }, platform);
  const meetingId = cleanId(firstNonEmpty(meetingIdValue(input), detected.meeting_id));
  const externalMeetingId = cleanId(firstNonEmpty(externalMeetingIdValue(input), detected.external_meeting_id));
  const stableUrlId = rawUrl ? stableMeetingIdFromUrl(rawUrl) : undefined;
  const title = firstNonEmpty(titleValue(input), detected.title);
  const startTimeMs = startTimeValue(input);
  const endTimeMs = endTimeValue(input);
  const confidence = firstNonEmpty(input.confidence, input.discovery?.confidence, detected.confidence, meetingId || externalMeetingId || stableUrlId ? 'medium' : 'low');
  const candidate = compactObject({
    id: `${source}-${cleanId(firstNonEmpty(meetingId, externalMeetingId, stableUrlId, title, 'unknown'))}`,
    source,
    source_rank: sourceRank(source),
    platform: key,
    meeting_id: meetingId || undefined,
    external_meeting_id: externalMeetingId || undefined,
    meeting_url: rawUrl ? String(rawUrl) : undefined,
    stable_url_id: stableUrlId || undefined,
    title,
    normalized_title: title ? normalizedText(title) : undefined,
    start_time_ms: startTimeMs,
    end_time_ms: endTimeMs,
    confidence,
    raw_type: input.type,
  });
  if (!candidate.meeting_id && !candidate.external_meeting_id && !candidate.stable_url_id && !candidate.title && !options.includeEmptyCandidates) {
    return undefined;
  }
  return candidate;
}

function normalizeSignalCandidate(signal = {}, platform) {
  if (!signal || typeof signal !== 'object') return undefined;
  const meeting = signal.meeting ?? {};
  return normalizeCandidate('provider_signal', {
    ...meeting,
    platform: firstNonEmpty(meeting.platform, signal.platform, platform),
    occurred_at_ms: signal.occurred_at_ms ?? signal.occurredAtMs,
    type: signal.type,
  }, platform);
}

function currentMeetingInput(input = {}) {
  return firstNonEmpty(input.current_meeting, input.currentMeeting, input.current_axis, input.currentAxis, input.meeting);
}

function localInputs(input = {}) {
  return [
    ...asArray(firstNonEmpty(input.local_observers, input.localObservers)),
    ...asArray(firstNonEmpty(input.local_meetings, input.localMeetings)),
    ...asArray(firstNonEmpty(input.browser_observers, input.browserObservers)),
    ...asArray(firstNonEmpty(input.native_observers, input.nativeObservers)),
    firstNonEmpty(input.local_observer, input.localObserver, input.local_meeting, input.localMeeting, input.browser, input.window, input.tab),
  ].filter(Boolean);
}

function providerSignalInputs(input = {}) {
  return [
    ...asArray(firstNonEmpty(input.provider_signals, input.providerSignals, input.signals)),
    ...asArray(firstNonEmpty(input.provider_events, input.providerEvents, input.events)),
    firstNonEmpty(input.provider_signal, input.providerSignal, input.signal, input.provider_event, input.providerEvent),
  ].filter(Boolean);
}

function annotationInput(input = {}) {
  return firstNonEmpty(input.annotation, input.mark, input.item);
}

function collectCandidates(platform, input = {}, options = {}) {
  const rows = [];
  const current = normalizeCandidate('current_axis', currentMeetingInput(input), platform, options);
  if (current) rows.push(current);
  for (const local of localInputs(input)) {
    const source = local?.source?.includes?.('native') ? 'native_observer' : 'local_observer';
    const candidate = normalizeCandidate(source, local, platform, options);
    if (candidate) rows.push(candidate);
  }
  for (const signal of providerSignalInputs(input)) {
    const candidate = normalizeSignalCandidate(signal, platform);
    if (candidate) rows.push(candidate);
  }
  const annotation = annotationInput(input);
  if (annotation) {
    const candidate = normalizeCandidate('annotation', annotation, platform, {
      ...options,
      includeEmptyCandidates: true,
    });
    if (candidate) rows.push(candidate);
  }
  return rows;
}

function exactMatch(a, b, field) {
  return a[field] && b[field] && cleanId(a[field]) === cleanId(b[field]);
}

function titleMatch(a, b) {
  if (!a.normalized_title || !b.normalized_title) return false;
  if (a.normalized_title === b.normalized_title) return true;
  return a.normalized_title.length > 6
    && b.normalized_title.length > 6
    && (a.normalized_title.includes(b.normalized_title) || b.normalized_title.includes(a.normalized_title));
}

function timeDeltaMs(a, b) {
  const at = firstNonEmpty(a.start_time_ms, a.end_time_ms);
  const bt = firstNonEmpty(b.start_time_ms, b.end_time_ms);
  return at != null && bt != null ? Math.abs(at - bt) : undefined;
}

function scorePair(a, b, options = {}) {
  const maxStartDeltaMs = numberOption(options, ['maxStartDeltaMs', 'max_start_delta_ms'], 10 * 60_000);
  const evidence = [];
  let score = 0;
  if (a.platform && b.platform && a.platform === b.platform) {
    score += 20;
    evidence.push('platform_match');
  }
  if (exactMatch(a, b, 'meeting_id')) {
    score += 60;
    evidence.push('meeting_id_match');
  }
  if (exactMatch(a, b, 'external_meeting_id')) {
    score += 55;
    evidence.push('external_meeting_id_match');
  }
  if (exactMatch(a, b, 'stable_url_id')) {
    score += 55;
    evidence.push('meeting_url_match');
  }
  if (titleMatch(a, b)) {
    score += 10;
    evidence.push('title_match');
  }
  const delta = timeDeltaMs(a, b);
  if (delta != null && delta <= maxStartDeltaMs) {
    score += 15;
    evidence.push('time_window_match');
  }
  if (a.platform && b.platform && a.platform !== b.platform) {
    score -= 30;
    evidence.push('platform_conflict');
  }
  if (a.meeting_id && b.meeting_id && a.meeting_id !== b.meeting_id) {
    score -= 40;
    evidence.push('meeting_id_conflict');
  }
  if (a.external_meeting_id && b.external_meeting_id && a.external_meeting_id !== b.external_meeting_id) {
    score -= 30;
    evidence.push('external_meeting_id_conflict');
  }
  return compactObject({
    a: a.id,
    b: b.id,
    score,
    evidence,
    time_delta_ms: delta,
  });
}

function bestReference(candidates = []) {
  return candidates.slice().sort((a, b) => {
    if (b.source_rank !== a.source_rank) return b.source_rank - a.source_rank;
    const ai = [a.meeting_id, a.external_meeting_id, a.stable_url_id].filter(Boolean).length;
    const bi = [b.meeting_id, b.external_meeting_id, b.stable_url_id].filter(Boolean).length;
    return bi - ai;
  })[0];
}

function mergeMeeting(platform, candidates = [], reference = {}) {
  const sorted = candidates.slice().sort((a, b) => b.source_rank - a.source_rank);
  return compactObject({
    platform,
    meeting_id: firstNonEmpty(reference.meeting_id, ...sorted.map((item) => item.meeting_id)),
    external_meeting_id: firstNonEmpty(reference.external_meeting_id, ...sorted.map((item) => item.external_meeting_id)),
    meeting_url: firstNonEmpty(reference.meeting_url, ...sorted.map((item) => item.meeting_url)),
    title: firstNonEmpty(reference.title, ...sorted.map((item) => item.title)),
    start_time_ms: firstNonEmpty(...sorted.map((item) => item.start_time_ms)),
    end_time_ms: firstNonEmpty(...sorted.map((item) => item.end_time_ms)),
  });
}

function statusFor({ current, local, provider, annotation, accepted, conflicts }) {
  if (conflicts.length > 0) return 'binding_conflict';
  if (current && accepted) return 'bound_to_current_axis';
  if (!current && local && accepted) return 'open_axis_from_local_observer';
  if (!current && provider && accepted) return 'open_axis_from_provider_start';
  if (annotation) return 'pending_binding';
  return 'insufficient_identity';
}

function nextActionsFor(status) {
  return {
    bound_to_current_axis: ['insert_realtime_annotation_on_current_axis', 'keep_provider_events_for_reconcile'],
    open_axis_from_local_observer: ['start_meeting_session_from_local_observer', 'insert_annotation_after_axis_start'],
    open_axis_from_provider_start: ['start_meeting_session_from_provider_event', 'prefer_local_observer_when_available'],
    pending_binding: ['store_annotation_as_pending_until_local_or_provider_meeting_identity_arrives'],
    binding_conflict: ['do_not_switch_current_axis_automatically', 'ask_host_to_resolve_meeting_identity_conflict'],
    insufficient_identity: ['collect_meeting_url_or_provider_meeting_id_before_binding'],
  }[status] ?? ['inspect_session_binding_status'];
}

export function buildMeetingPlatformSessionBindingPlan(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const integration = buildPlatformIntegrationPlan(key, options);
  return {
    type: 'meeting_platform_session_binding_plan',
    schema: MEETING_PLATFORM_SESSION_BINDING_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_SESSION_BINDING_SCHEMA_VERSION,
    platform: key,
    display_name: integration.display_name,
    status: 'session_binding_contract_ready',
    identity_fields: ['platform', 'meeting_id', 'external_meeting_id', 'meeting_url', 'stable_url_id', 'title', 'start_time_ms'],
    match_policy: {
      accept_score: numberOption(options, ['acceptScore', 'accept_score'], 70),
      conflict_score: numberOption(options, ['conflictScore', 'conflict_score'], 20),
      max_start_delta_ms: numberOption(options, ['maxStartDeltaMs', 'max_start_delta_ms'], 10 * 60_000),
    },
    realtime_policy: {
      local_observer_preferred: true,
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      annotation_requires_captured_at_ms: true,
    },
    output_contract: {
      status_values: [
        'bound_to_current_axis',
        'open_axis_from_local_observer',
        'open_axis_from_provider_start',
        'pending_binding',
        'binding_conflict',
        'insufficient_identity',
      ],
      start_payload_field: 'start_payload',
      selected_meeting_field: 'selected_meeting',
    },
    next_actions: [
      'collect_local_observer_identity',
      'collect_provider_meeting_identity_for_reconcile',
      'bind_annotation_to_selected_meeting_identity',
    ],
  };
}

export function buildMeetingPlatformSessionBindingMatrix(options = {}) {
  const plans = selectedPlatforms(options).map((platform) => buildMeetingPlatformSessionBindingPlan(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_session_binding_matrix',
    schema: MEETING_PLATFORM_SESSION_BINDING_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_SESSION_BINDING_SCHEMA_VERSION,
    platform_count: plans.length,
    provider_blocking_count: plans.filter((plan) => plan.realtime_policy.provider_events_block_realtime).length,
    transcript_blocking_count: plans.filter((plan) => plan.realtime_policy.transcript_blocks_realtime).length,
    platforms: plans.map((plan) => plan.platform),
    rows: plans.map((plan) => ({
      platform: plan.platform,
      display_name: plan.display_name,
      status: plan.status,
      local_observer_preferred: plan.realtime_policy.local_observer_preferred,
      provider_events_block_realtime: plan.realtime_policy.provider_events_block_realtime,
      transcript_blocks_realtime: plan.realtime_policy.transcript_blocks_realtime,
      accept_score: plan.match_policy.accept_score,
    })),
    plans,
    next_actions: unique(plans.flatMap((plan) => plan.next_actions ?? [])),
  };
}

export function buildMeetingPlatformSessionBinding(platform, input = {}, options = {}) {
  const key = normalizeMeetingPlatform(firstNonEmpty(input.platform, platform));
  const acceptScore = numberOption(options, ['acceptScore', 'accept_score'], 70);
  const conflictScore = numberOption(options, ['conflictScore', 'conflict_score'], 20);
  const candidates = collectCandidates(key, input, options);
  const current = candidates.find((candidate) => candidate.source === 'current_axis');
  const local = candidates.find((candidate) => candidate.source === 'local_observer' || candidate.source === 'browser_observer' || candidate.source === 'native_observer');
  const provider = candidates.find((candidate) => candidate.source === 'provider_signal');
  const annotation = candidates.find((candidate) => candidate.source === 'annotation');
  const reference = current ?? local ?? provider ?? bestReference(candidates);
  const matches = reference
    ? candidates
      .filter((candidate) => candidate.id !== reference.id)
      .map((candidate) => scorePair(reference, candidate, options))
    : [];
  const conflicts = matches.filter((match) => match.score < conflictScore && (
    match.evidence?.includes('platform_conflict') || match.evidence?.includes('meeting_id_conflict')
  ));
  const positiveMatches = matches.filter((match) => match.score >= acceptScore);
  const enoughIdentity = Boolean(reference?.meeting_id || reference?.external_meeting_id || reference?.stable_url_id || reference?.title);
  const accepted = Boolean(enoughIdentity && (matches.length === 0 || positiveMatches.length > 0 || reference.source !== 'current_axis'));
  const status = statusFor({ current, local, provider, annotation, accepted, conflicts });
  const selectedMeeting = accepted && conflicts.length === 0 ? mergeMeeting(key, candidates, reference) : undefined;
  const startPayload = ['open_axis_from_local_observer', 'open_axis_from_provider_start'].includes(status)
    ? buildMeetingStartPayload({
      ...selectedMeeting,
      detector_source: status === 'open_axis_from_local_observer' ? 'local_observer_session_binding' : 'provider_event_session_binding',
      suppress_auto_annotations: true,
    })
    : undefined;
  return compactObject({
    type: 'meeting_platform_session_binding',
    schema: MEETING_PLATFORM_SESSION_BINDING_SCHEMA,
    schema_version: MEETING_PLATFORM_SESSION_BINDING_SCHEMA_VERSION,
    platform: key,
    status,
    accepted_for_realtime: ['bound_to_current_axis', 'open_axis_from_local_observer', 'open_axis_from_provider_start', 'pending_binding'].includes(status),
    bind_to_current_axis: status === 'bound_to_current_axis',
    should_start_axis: status === 'open_axis_from_local_observer' || status === 'open_axis_from_provider_start',
    should_store_pending: status === 'pending_binding',
    reference_candidate_id: reference?.id,
    selected_meeting: selectedMeeting,
    start_payload: startPayload,
    candidate_count: candidates.length,
    match_count: matches.length,
    positive_match_count: positiveMatches.length,
    conflict_count: conflicts.length,
    candidates,
    matches,
    conflicts,
    policy: {
      accept_score: acceptScore,
      conflict_score: conflictScore,
      max_start_delta_ms: numberOption(options, ['maxStartDeltaMs', 'max_start_delta_ms'], 10 * 60_000),
    },
    next_actions: nextActionsFor(status),
  });
}
