import { compactObject } from '../index.mjs';
import { meetingAppSnapshotRecords } from './meeting-app-snapshot-recorder.mjs';
import { normalizePlatformEvent } from './platform-ingest.mjs';
import { normalizeMeetingPlatform } from './platform-setup.mjs';

export const MEETING_PLATFORM_EVIDENCE_CORRELATION_SCHEMA = 'meeting_platform_evidence_correlation';
export const MEETING_PLATFORM_EVIDENCE_CORRELATION_SCHEMA_VERSION = 1;

const DEFAULT_MAX_CLOCK_SKEW_MS = 10 * 60 * 1000;

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  return Array.isArray(value) ? value : value == null ? [] : [value];
}

function uniqueList(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function issue(severity, code, message, details = {}) {
  return compactObject({
    severity,
    code,
    message,
    ...details,
  });
}

function normalizeUrl(value) {
  if (!value) return undefined;
  try {
    const parsed = new URL(String(value));
    parsed.hash = '';
    parsed.search = '';
    return parsed.toString().replace(/\/+$/, '').toLowerCase();
  } catch {
    return String(value).trim().replace(/\/+$/, '').toLowerCase();
  }
}

function platformOfRecord(record = {}) {
  return firstNonEmpty(record.platform, record.provider, record.adapter, record.snapshot?.platform, record.snapshot?.provider);
}

function platformMatches(record = {}, platform) {
  const raw = platformOfRecord(record);
  if (!raw) return true;
  try {
    return normalizeMeetingPlatform(raw) === platform;
  } catch {
    return false;
  }
}

function recordTimestamp(record = {}) {
  const value = firstNonEmpty(record.captured_at_ms, record.capturedAtMs, record.observed_at_ms, record.observedAtMs);
  const number = Number(value);
  return Number.isFinite(number) ? Math.round(number) : undefined;
}

function providerPayload(record = {}) {
  return firstNonEmpty(record.body, record.payload, record.raw_event, record.rawEvent, record.raw);
}

function providerReceivedAtMs(record = {}) {
  return firstNonEmpty(record.received_at_ms, record.receivedAtMs, record.captured_at_ms, record.capturedAtMs);
}

function normalizeProviderSignals(platform, providerRecords = []) {
  const signals = [];
  const errors = [];
  for (const record of asArray(providerRecords)) {
    if (!record || !platformMatches(record, platform)) continue;
    try {
      const normalized = normalizePlatformEvent(platform, providerPayload(record), {
        receivedAtMs: providerReceivedAtMs(record),
      });
      for (const signal of normalized.signals ?? []) {
        signals.push({
          ...signal,
          capture_id: record.id,
          captured_at_ms: recordTimestamp(record),
        });
      }
    } catch (error) {
      errors.push(issue(
        'warning',
        'provider_record_not_normalized',
        'Provider capture record could not be normalized for evidence correlation.',
        { capture_id: record.id, error: String(error?.message ?? error) },
      ));
    }
  }
  return { signals, errors };
}

function meetingIdentityFromProvider(signal = {}) {
  const meeting = signal.meeting ?? {};
  return compactObject({
    platform: meeting.platform ?? signal.platform,
    meeting_id: meeting.meeting_id,
    external_meeting_id: meeting.external_meeting_id,
    meeting_url: meeting.meeting_url,
    title: meeting.title,
  });
}

function meetingIdentityFromDom(record = {}) {
  return compactObject({
    platform: record.platform ?? record.provider ?? record.snapshot?.platform ?? record.snapshot?.provider,
    meeting_id: firstNonEmpty(record.meeting_id, record.meetingId, record.snapshot?.meeting_id, record.snapshot?.meetingId),
    external_meeting_id: firstNonEmpty(
      record.external_meeting_id,
      record.externalMeetingId,
      record.snapshot?.external_meeting_id,
      record.snapshot?.externalMeetingId,
    ),
    meeting_url: firstNonEmpty(record.meeting_url, record.meetingUrl, record.url, record.snapshot?.meeting_url, record.snapshot?.meetingUrl, record.snapshot?.url),
    title: firstNonEmpty(record.title, record.snapshot?.title),
  });
}

function identityTokens(identity = {}) {
  return uniqueList([
    identity.meeting_id ? `id:${identity.meeting_id}` : undefined,
    identity.external_meeting_id ? `external:${identity.external_meeting_id}` : undefined,
    identity.meeting_url ? `url:${normalizeUrl(identity.meeting_url)}` : undefined,
  ]);
}

function identitySummary(identities = []) {
  const tokens = uniqueList(identities.flatMap((identity) => identityTokens(identity)));
  return {
    count: identities.length,
    tokens,
    meetings: identities,
  };
}

function minMax(values = []) {
  const numbers = values.map(Number).filter(Number.isFinite);
  if (numbers.length === 0) return {};
  return {
    min_ms: Math.min(...numbers),
    max_ms: Math.max(...numbers),
  };
}

function providerWindow(signals = []) {
  const startTimes = signals.filter((signal) => signal.type === 'meeting_started').map((signal) => signal.occurred_at_ms);
  const endTimes = signals.filter((signal) => signal.type === 'meeting_ended').map((signal) => signal.occurred_at_ms);
  const allTimes = signals.map((signal) => firstNonEmpty(signal.occurred_at_ms, signal.captured_at_ms));
  const fallback = minMax(allTimes);
  return compactObject({
    start_ms: minMax(startTimes).min_ms ?? fallback.min_ms,
    end_ms: minMax(endTimes).max_ms ?? fallback.max_ms,
    observed_min_ms: fallback.min_ms,
    observed_max_ms: fallback.max_ms,
  });
}

function domWindow(records = []) {
  const activeTimes = records.filter((record) => record.phase === 'active').map(recordTimestamp);
  const endedTimes = records.filter((record) => record.phase === 'ended').map(recordTimestamp);
  const allTimes = records.map(recordTimestamp);
  const fallback = minMax(allTimes);
  return compactObject({
    start_ms: minMax(activeTimes).min_ms ?? fallback.min_ms,
    end_ms: minMax(endedTimes).max_ms ?? fallback.max_ms,
    observed_min_ms: fallback.min_ms,
    observed_max_ms: fallback.max_ms,
  });
}

function overlapMs(left = {}, right = {}) {
  const leftStart = firstNonEmpty(left.start_ms, left.observed_min_ms);
  const leftEnd = firstNonEmpty(left.end_ms, left.observed_max_ms, leftStart);
  const rightStart = firstNonEmpty(right.start_ms, right.observed_min_ms);
  const rightEnd = firstNonEmpty(right.end_ms, right.observed_max_ms, rightStart);
  if ([leftStart, leftEnd, rightStart, rightEnd].some((value) => !Number.isFinite(Number(value)))) return undefined;
  return Math.max(0, Math.min(Number(leftEnd), Number(rightEnd)) - Math.max(Number(leftStart), Number(rightStart)));
}

function deltaMs(left, right) {
  if (!Number.isFinite(Number(left)) || !Number.isFinite(Number(right))) return undefined;
  return Math.abs(Number(left) - Number(right));
}

function timeAlignment(provider = {}, dom = {}, maxClockSkewMs = DEFAULT_MAX_CLOCK_SKEW_MS) {
  const startDelta = deltaMs(provider.start_ms, dom.start_ms);
  const endDelta = deltaMs(provider.end_ms, dom.end_ms);
  const overlap = overlapMs(provider, dom);
  const hasComparableTimes = startDelta != null || endDelta != null || overlap != null;
  const passed = !hasComparableTimes
    ? true
    : [
      startDelta == null || startDelta <= maxClockSkewMs,
      endDelta == null || endDelta <= maxClockSkewMs,
      overlap == null || overlap > 0 || startDelta <= maxClockSkewMs || endDelta <= maxClockSkewMs,
    ].every(Boolean);
  return compactObject({
    passed,
    max_clock_skew_ms: maxClockSkewMs,
    provider_start_ms: provider.start_ms,
    provider_end_ms: provider.end_ms,
    dom_start_ms: dom.start_ms,
    dom_end_ms: dom.end_ms,
    start_delta_ms: startDelta,
    end_delta_ms: endDelta,
    overlap_ms: overlap,
  });
}

function resolveCorrelationStatus({ providerSignals, domRecords, identityMatched, providerTokens, domTokens, timePassed, issues }) {
  const blocking = issues.filter((item) => item.severity === 'error');
  if (blocking.length > 0) return { status: 'failed', passed: false, confidence: 'none' };
  if (providerSignals.length === 0 || domRecords.length === 0) {
    return { status: 'single_source', passed: true, confidence: 'none' };
  }
  if (identityMatched && timePassed) return { status: 'matched', passed: true, confidence: 'high' };
  if (timePassed) {
    return {
      status: providerTokens.length > 0 && domTokens.length > 0 ? 'time_matched_identity_unconfirmed' : 'time_matched',
      passed: true,
      confidence: 'medium',
    };
  }
  return { status: 'failed', passed: false, confidence: 'none' };
}

export function buildMeetingPlatformEvidenceCorrelation(platform, input = {}, options = {}) {
  const key = normalizeMeetingPlatform(platform ?? input.platform ?? input.provider ?? input.adapter);
  const providerRecords = asArray(firstNonEmpty(
    input.providerRecords,
    input.provider_records,
    input.providerCaptureRecords,
    input.provider_capture_records,
    input.captureRecords,
    input.capture_records,
    input.provider_records,
  ));
  const domRecords = meetingAppSnapshotRecords(firstNonEmpty(
    input.meetingAppRecordSet,
    input.meeting_app_record_set,
    input.recordSet,
    input.record_set,
    input.meetingAppRecords,
    input.meeting_app_records,
    [],
  )).filter((record) => platformMatches(record, key));
  const maxClockSkewMs = Number(firstNonEmpty(
    options.maxClockSkewMs,
    options.max_clock_skew_ms,
    input.maxClockSkewMs,
    input.max_clock_skew_ms,
    DEFAULT_MAX_CLOCK_SKEW_MS,
  ));
  const { signals: providerSignals, errors } = normalizeProviderSignals(key, providerRecords);
  const providerMeetings = providerSignals.map(meetingIdentityFromProvider).filter((identity) => identity.meeting_id || identity.external_meeting_id || identity.meeting_url);
  const domMeetings = domRecords.map(meetingIdentityFromDom).filter((identity) => identity.meeting_id || identity.external_meeting_id || identity.meeting_url);
  const providerSummary = identitySummary(providerMeetings);
  const domSummary = identitySummary(domMeetings);
  const sharedTokens = providerSummary.tokens.filter((token) => domSummary.tokens.includes(token));
  const identityMatched = sharedTokens.length > 0;
  const providerTime = providerWindow(providerSignals);
  const domTime = domWindow(domRecords);
  const time = timeAlignment(providerTime, domTime, maxClockSkewMs);
  const issues = [...errors];

  if (providerRecords.length === 0 || domRecords.length === 0) {
    issues.push(issue(
      'warning',
      'single_source_correlation',
      'Evidence correlation has only one evidence source; use rollout status for pilot, but production should include provider and local observer evidence.',
      { provider_record_count: providerRecords.length, dom_record_count: domRecords.length },
    ));
  }
  if (providerRecords.length > 0 && providerSignals.length === 0) {
    issues.push(issue('error', 'provider_records_without_signals', 'Provider records did not normalize into meeting signals.'));
  }
  if (providerRecords.length > 0 && domRecords.length > 0 && providerSummary.tokens.length > 0 && domSummary.tokens.length > 0 && !identityMatched) {
    issues.push(issue(
      'warning',
      'identity_not_confirmed',
      'Provider and local observer evidence did not share an explicit meeting id or URL; time alignment is used as fallback.',
      { provider_tokens: providerSummary.tokens, dom_tokens: domSummary.tokens },
    ));
  }
  if (time.passed === false) {
    issues.push(issue(
      'error',
      'time_window_mismatch',
      'Provider and local observer evidence are outside the allowed same-meeting time window.',
      { time_alignment: time },
    ));
  }

  const status = resolveCorrelationStatus({
    providerSignals,
    domRecords,
    identityMatched,
    providerTokens: providerSummary.tokens,
    domTokens: domSummary.tokens,
    timePassed: time.passed,
    issues,
  });

  return compactObject({
    schema: MEETING_PLATFORM_EVIDENCE_CORRELATION_SCHEMA,
    schema_version: MEETING_PLATFORM_EVIDENCE_CORRELATION_SCHEMA_VERSION,
    type: 'meeting_platform_evidence_correlation',
    platform: key,
    ...status,
    provider_record_count: providerRecords.length,
    provider_signal_count: providerSignals.length,
    meeting_app_record_count: domRecords.length,
    identity_match: {
      passed: identityMatched,
      shared_tokens: sharedTokens,
      provider: providerSummary,
      local_observer: domSummary,
    },
    time_alignment: time,
    coverage: {
      provider_has_start: providerSignals.some((signal) => signal.type === 'meeting_started'),
      provider_has_end: providerSignals.some((signal) => signal.type === 'meeting_ended'),
      local_has_active: domRecords.some((record) => record.phase === 'active'),
      local_has_ended: domRecords.some((record) => record.phase === 'ended'),
    },
    issues,
  });
}
