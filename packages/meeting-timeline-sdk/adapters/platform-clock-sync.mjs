import {
  compactObject,
  normalizeAbsoluteMs,
} from '../index.mjs';
import {
  MEETING_PLATFORM_KEYS,
  buildPlatformIntegrationPlan,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_CLOCK_SYNC_PLAN_SCHEMA = 'meeting_platform_clock_sync_plan';
export const MEETING_PLATFORM_CLOCK_SYNC_MATRIX_SCHEMA = 'meeting_platform_clock_sync_matrix';
export const MEETING_PLATFORM_CLOCK_SYNC_REPORT_SCHEMA = 'meeting_platform_clock_sync_report';
export const MEETING_PLATFORM_CLOCK_SYNC_SCHEMA_VERSION = 1;

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

function boolOption(options, keys, fallback) {
  const value = firstNonEmpty(...keys.map((key) => options[key]));
  return value == null ? fallback : value !== false && value !== 'false';
}

function selectedPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_PLATFORM_KEYS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

function maybeAbsoluteMs(value, fieldName) {
  if (value == null || value === '') return undefined;
  try {
    return normalizeAbsoluteMs(value, fieldName);
  } catch {
    return undefined;
  }
}

function maybeNumber(value) {
  if (value == null || value === '') return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function round(value) {
  return Number.isFinite(value) ? Math.round(value) : undefined;
}

function median(values = []) {
  const rows = values.filter(Number.isFinite).slice().sort((a, b) => a - b);
  if (rows.length === 0) return undefined;
  const middle = Math.floor(rows.length / 2);
  return rows.length % 2 ? rows[middle] : (rows[middle - 1] + rows[middle]) / 2;
}

function policy(options = {}) {
  return {
    endpoint: String(firstNonEmpty(options.endpoint, options.timeSyncEndpoint, options.time_sync_endpoint, '/api/time')),
    sample_count: numberOption(options, ['sampleCount', 'sample_count'], 3),
    max_recommended_skew_ms: numberOption(options, ['maxRecommendedSkewMs', 'max_recommended_skew_ms'], 500),
    max_recommended_rtt_ms: numberOption(options, ['maxRecommendedRttMs', 'max_recommended_rtt_ms'], 1000),
    max_uncertainty_ms: numberOption(options, ['maxUncertaintyMs', 'max_uncertainty_ms'], 500),
    apply_offset_to_annotation: boolOption(options, ['applyOffsetToAnnotation', 'apply_offset_to_annotation'], true),
  };
}

function sampleRows(input = {}) {
  const directRows = firstNonEmpty(
    input.samples,
    input.clock_samples,
    input.clockSamples,
    input.time_sync_samples,
    input.timeSyncSamples,
  );
  const rows = asArray(directRows);
  const directSync = firstNonEmpty(input.clock_sync, input.clockSync, input.time_sync, input.timeSync);
  if (directSync) rows.push(directSync);
  return rows;
}

function normalizeClockSyncSample(sample = {}, index = 0, options = {}) {
  const sendAtMs = maybeAbsoluteMs(firstNonEmpty(
    sample.client_send_at_ms,
    sample.clientSendAtMs,
    sample.send_at_ms,
    sample.sendAtMs,
    sample.sent_at_ms,
    sample.sentAtMs,
    sample.request_started_at_ms,
    sample.requestStartedAtMs,
  ), 'client_send_at_ms');
  const receiveAtMs = maybeAbsoluteMs(firstNonEmpty(
    sample.client_receive_at_ms,
    sample.clientReceiveAtMs,
    sample.receive_at_ms,
    sample.receiveAtMs,
    sample.received_at_ms,
    sample.receivedAtMs,
    sample.response_received_at_ms,
    sample.responseReceivedAtMs,
  ), 'client_receive_at_ms');
  const serverTimeMs = maybeAbsoluteMs(firstNonEmpty(
    sample.server_time_ms,
    sample.serverTimeMs,
    sample.server_time,
    sample.serverTime,
  ), 'server_time_ms');
  const directOffsetMs = maybeNumber(firstNonEmpty(
    sample.offset_ms,
    sample.offsetMs,
    sample.clock_offset_ms,
    sample.clockOffsetMs,
    sample.estimated_offset_ms,
    sample.estimatedOffsetMs,
  ));
  const directRttMs = maybeNumber(firstNonEmpty(sample.rtt_ms, sample.rttMs, sample.round_trip_ms, sample.roundTripMs));
  const rttMs = receiveAtMs != null && sendAtMs != null ? receiveAtMs - sendAtMs : directRttMs;
  const midpointMs = receiveAtMs != null && sendAtMs != null ? (sendAtMs + receiveAtMs) / 2 : undefined;
  const offsetMs = serverTimeMs != null && midpointMs != null ? serverTimeMs - midpointMs : directOffsetMs;
  const uncertaintyMs = rttMs == null ? undefined : Math.ceil(Math.max(0, rttMs) / 2);
  const warnings = [];
  if (offsetMs == null) warnings.push('missing_offset_ms');
  if (rttMs == null) warnings.push('missing_rtt_ms');
  if (rttMs != null && rttMs < 0) warnings.push('negative_rtt_ms');
  if (rttMs != null && rttMs > options.max_recommended_rtt_ms) warnings.push('rtt_above_recommended');
  if (offsetMs != null && Math.abs(offsetMs) > options.max_recommended_skew_ms) warnings.push('clock_skew_above_recommended');
  if (uncertaintyMs != null && uncertaintyMs > options.max_uncertainty_ms) warnings.push('uncertainty_above_recommended');
  const usable = offsetMs != null && (rttMs == null || rttMs >= 0);
  return compactObject({
    id: firstNonEmpty(sample.id, sample.sample_id, sample.sampleId, `clock-sample-${index + 1}`),
    source: firstNonEmpty(sample.source, sample.method, midpointMs == null ? 'direct_offset' : 'midpoint_ntp'),
    client_send_at_ms: sendAtMs,
    client_receive_at_ms: receiveAtMs,
    server_time_ms: serverTimeMs,
    midpoint_ms: midpointMs == null ? undefined : round(midpointMs),
    offset_ms: round(offsetMs),
    rtt_ms: round(rttMs),
    uncertainty_ms: uncertaintyMs,
    usable,
    warnings,
    raw: sample.raw,
  });
}

function chooseSample(samples = []) {
  const usable = samples.filter((sample) => sample.usable && Number.isFinite(sample.offset_ms));
  if (usable.length === 0) return undefined;
  return usable.slice().sort((a, b) => {
    const ar = Number.isFinite(a.rtt_ms) ? a.rtt_ms : Number.POSITIVE_INFINITY;
    const br = Number.isFinite(b.rtt_ms) ? b.rtt_ms : Number.POSITIVE_INFINITY;
    if (ar !== br) return ar - br;
    return Math.abs(a.offset_ms) - Math.abs(b.offset_ms);
  })[0];
}

function annotationInput(input = {}) {
  return firstNonEmpty(input.annotation, input.mark, input.item, input.payload?.annotation);
}

function capturedAtMs(input = {}) {
  return maybeAbsoluteMs(firstNonEmpty(
    input.captured_at_ms,
    input.capturedAtMs,
    input.captured_at,
    input.capturedAt,
    input.ink_end_at_ms,
    input.inkEndAtMs,
    input.timestamp_ms,
    input.timestampMs,
    input.timestamp,
    input.ts,
    input.payload?.timing?.captured_at_ms,
    input.payload?.timing?.capturedAtMs,
  ), 'annotation.captured_at_ms');
}

function calibratedAnnotation(annotation, offsetMs, enabled) {
  if (!annotation || offsetMs == null) return undefined;
  const rawCapturedAtMs = capturedAtMs(annotation);
  if (rawCapturedAtMs == null) return undefined;
  const calibratedCapturedAtMs = enabled ? rawCapturedAtMs + offsetMs : rawCapturedAtMs;
  return {
    annotation: compactObject({
      ...annotation,
      captured_at_ms: round(calibratedCapturedAtMs),
      clock_sync: {
        applied: enabled,
        offset_ms: round(offsetMs),
        raw_captured_at_ms: rawCapturedAtMs,
      },
    }),
    timing: {
      raw_captured_at_ms: rawCapturedAtMs,
      calibrated_captured_at_ms: round(calibratedCapturedAtMs),
      applied_offset_ms: enabled ? round(offsetMs) : 0,
    },
  };
}

function statusFor({ samples = [], selectedSample, recommendedOffsetMs, syncPolicy }) {
  if (samples.length === 0) return 'clock_sync_missing';
  if (!selectedSample) return 'clock_sync_unusable';
  if ((selectedSample.uncertainty_ms ?? 0) > syncPolicy.max_uncertainty_ms) return 'clock_sync_unstable';
  if (Math.abs(recommendedOffsetMs ?? 0) > syncPolicy.max_recommended_skew_ms) return 'clock_sync_requires_offset';
  return 'clock_sync_ready';
}

function nextActionsFor(status) {
  return {
    clock_sync_ready: ['apply_clock_offset_before_sending_captured_at_ms'],
    clock_sync_requires_offset: ['apply_clock_offset_to_device_captured_at_ms', 'alert_operator_if_device_clock_remains_skewed'],
    clock_sync_unstable: ['retry_time_sync_with_lower_rtt', 'avoid_counting_marks_as_realtime_until_clock_is_stable'],
    clock_sync_unusable: ['collect_client_send_server_time_and_client_receive_samples'],
    clock_sync_missing: ['call_time_sync_endpoint_before_sending_annotations'],
  }[status] ?? ['inspect_clock_sync_report'];
}

export function buildMeetingPlatformClockSyncPlan(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const integration = buildPlatformIntegrationPlan(key, options);
  const syncPolicy = policy(options);
  return {
    type: 'meeting_platform_clock_sync_plan',
    schema: MEETING_PLATFORM_CLOCK_SYNC_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_CLOCK_SYNC_SCHEMA_VERSION,
    platform: key,
    display_name: integration.display_name,
    status: 'clock_sync_contract_ready',
    required: true,
    endpoint: syncPolicy.endpoint,
    algorithm: {
      name: 'midpoint_offset',
      formula: 'clock_offset_ms = server_time_ms - ((client_send_at_ms + client_receive_at_ms) / 2)',
      captured_at_formula: 'captured_at_ms = device_mark_end_ms + clock_offset_ms',
      recommended_sample_count: syncPolicy.sample_count,
    },
    thresholds: {
      max_recommended_skew_ms: syncPolicy.max_recommended_skew_ms,
      max_recommended_rtt_ms: syncPolicy.max_recommended_rtt_ms,
      max_uncertainty_ms: syncPolicy.max_uncertainty_ms,
    },
    input_contract: {
      sample_fields: ['client_send_at_ms', 'server_time_ms', 'client_receive_at_ms'],
      direct_offset_fields: ['offset_ms', 'rtt_ms'],
      annotation_time_fields: ['captured_at_ms', 'ink_end_at_ms', 'timestamp_ms'],
    },
    realtime_policy: {
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      device_clock_sync_required: true,
    },
    next_actions: [
      'sample_time_sync_endpoint_before_annotation',
      'apply_offset_to_device_mark_end_time',
      'send_calibrated_captured_at_ms_to_annotation_intake',
    ],
  };
}

export function buildMeetingPlatformClockSyncMatrix(options = {}) {
  const plans = selectedPlatforms(options).map((platform) => buildMeetingPlatformClockSyncPlan(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_clock_sync_matrix',
    schema: MEETING_PLATFORM_CLOCK_SYNC_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_CLOCK_SYNC_SCHEMA_VERSION,
    platform_count: plans.length,
    required_count: plans.filter((plan) => plan.required).length,
    provider_blocking_count: plans.filter((plan) => plan.realtime_policy.provider_events_block_realtime).length,
    transcript_blocking_count: plans.filter((plan) => plan.realtime_policy.transcript_blocks_realtime).length,
    platforms: plans.map((plan) => plan.platform),
    rows: plans.map((plan) => ({
      platform: plan.platform,
      display_name: plan.display_name,
      endpoint: plan.endpoint,
      required: plan.required,
      max_recommended_skew_ms: plan.thresholds.max_recommended_skew_ms,
      provider_events_block_realtime: plan.realtime_policy.provider_events_block_realtime,
      transcript_blocks_realtime: plan.realtime_policy.transcript_blocks_realtime,
    })),
    plans,
    next_actions: unique(plans.flatMap((plan) => plan.next_actions ?? [])),
  };
}

export function buildMeetingPlatformClockSyncReport(platform, input = {}, options = {}) {
  const key = normalizeMeetingPlatform(firstNonEmpty(input.platform, platform));
  const syncPolicy = policy(options);
  const samples = sampleRows(input).map((sample, index) => normalizeClockSyncSample(sample, index, syncPolicy));
  const selectedSample = chooseSample(samples);
  const offsets = samples.filter((sample) => sample.usable && Number.isFinite(sample.offset_ms)).map((sample) => sample.offset_ms);
  const recommendedOffsetMs = selectedSample?.offset_ms;
  const medianOffsetMs = round(median(offsets));
  const offsetJitterMs = offsets.length > 1 ? Math.max(...offsets) - Math.min(...offsets) : 0;
  const status = statusFor({ samples, selectedSample, recommendedOffsetMs, syncPolicy });
  const acceptedForRealtime = Boolean(
    selectedSample
      && (selectedSample.uncertainty_ms ?? 0) <= syncPolicy.max_uncertainty_ms
      && !selectedSample.warnings.includes('negative_rtt_ms'),
  );
  const warnings = unique(samples.flatMap((sample) => sample.warnings ?? []));
  const annotation = annotationInput(input);
  const calibrated = calibratedAnnotation(annotation, recommendedOffsetMs, syncPolicy.apply_offset_to_annotation);
  return compactObject({
    type: 'meeting_platform_clock_sync_report',
    schema: MEETING_PLATFORM_CLOCK_SYNC_REPORT_SCHEMA,
    schema_version: MEETING_PLATFORM_CLOCK_SYNC_SCHEMA_VERSION,
    platform: key,
    status,
    accepted_for_realtime: acceptedForRealtime,
    endpoint: syncPolicy.endpoint,
    sample_count: samples.length,
    usable_sample_count: offsets.length,
    selected_sample_id: selectedSample?.id,
    selected_rtt_ms: selectedSample?.rtt_ms,
    selected_uncertainty_ms: selectedSample?.uncertainty_ms,
    recommended_offset_ms: recommendedOffsetMs,
    median_offset_ms: medianOffsetMs,
    offset_jitter_ms: round(offsetJitterMs),
    max_recommended_skew_ms: syncPolicy.max_recommended_skew_ms,
    max_recommended_rtt_ms: syncPolicy.max_recommended_rtt_ms,
    max_uncertainty_ms: syncPolicy.max_uncertainty_ms,
    operator_attention_required: warnings.includes('clock_skew_above_recommended') || warnings.includes('rtt_above_recommended'),
    warnings,
    samples,
    annotation_time: calibrated?.timing,
    calibrated_annotation: calibrated?.annotation,
    next_actions: nextActionsFor(status),
  });
}
