import { compactObject } from '../index.mjs';
import { diagnosePlatformEvent } from './platform-ingest.mjs';
import {
  MEETING_PLATFORM_KEYS,
  buildPlatformIntegrationPlan,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const PLATFORM_ACCEPTANCE_COVERAGE_KEYS = Object.freeze([
  'meeting_start',
  'meeting_end',
  'participant_track',
  'speaker_activity',
  'artifact_ready',
  'subscription_lifecycle',
]);

function asArray(value) {
  return Array.isArray(value) ? value : value == null ? [] : [value];
}

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function issue(severity, code, message, details = {}) {
  return compactObject({
    severity,
    code,
    message,
    ...details,
  });
}

function normalizeCoverageKeys(value, fallback = ['meeting_start']) {
  const rows = asArray(value ?? fallback)
    .map((item) => String(item || '').trim())
    .filter(Boolean);
  return [...new Set(rows)];
}

function defaultRequiredCoverage(platform, options = {}) {
  if (options.requiredCoverage || options.required_coverage) {
    return normalizeCoverageKeys(options.requiredCoverage ?? options.required_coverage);
  }
  const keys = ['meeting_start'];
  if (options.requireEndEvent === true || options.require_end_event === true) keys.push('meeting_end');
  if (platform === 'local_detector' && (options.requireEndEvent == null && options.require_end_event == null)) {
    return ['meeting_start'];
  }
  return keys;
}

function emptyCoverage() {
  return {
    realtime_axis: false,
    meeting_start: false,
    meeting_end: false,
    participant_track: false,
    speaker_activity: false,
    artifact_ready: false,
    subscription_lifecycle: false,
  };
}

function mergeCoverage(diagnostics = []) {
  const coverage = emptyCoverage();
  for (const diagnostic of diagnostics) {
    for (const [key, value] of Object.entries(diagnostic.coverage ?? {})) {
      coverage[key] = Boolean(coverage[key] || value);
    }
  }
  coverage.realtime_axis = Boolean(coverage.meeting_start || coverage.meeting_end);
  return coverage;
}

function mergedSignalTypes(diagnostics = []) {
  return [...new Set(diagnostics.flatMap((item) => item.signal_types ?? []))];
}

function sampleCollectionFor(platform, samplesInput = undefined) {
  if (samplesInput == null) return [];
  if (Array.isArray(samplesInput)) {
    return samplesInput.filter((sample) => {
      const samplePlatform = sample?.platform ?? sample?.provider ?? sample?.adapter;
      if (!samplePlatform) return true;
      try {
        return normalizeMeetingPlatform(samplePlatform) === platform;
      } catch {
        return false;
      }
    });
  }
  if (typeof samplesInput !== 'object') return [];
  for (const [key, value] of Object.entries(samplesInput)) {
    try {
      if (normalizeMeetingPlatform(key) === platform) return asArray(value);
    } catch {
      // Ignore non-platform keys.
    }
  }
  return [];
}

function samplePayload(sample = {}) {
  if (sample && typeof sample === 'object' && !Array.isArray(sample)) {
    if (sample.body !== undefined) return sample.body;
    if (sample.payload !== undefined) return sample.payload;
    if (sample.raw_event !== undefined) return sample.raw_event;
    if (sample.rawEvent !== undefined) return sample.rawEvent;
    if (sample.raw !== undefined) return sample.raw;
  }
  return sample;
}

function sampleOptions(sample = {}, options = {}) {
  if (!sample || typeof sample !== 'object' || Array.isArray(sample)) return options.diagnosticOptions ?? options.diagnostic_options ?? {};
  return {
    ...(options.diagnosticOptions ?? options.diagnostic_options ?? {}),
    ...(sample.options ?? {}),
    receivedAtMs: firstNonEmpty(
      sample.receivedAtMs,
      sample.received_at_ms,
      sample.received_at,
      sample.timestamp,
      sample.ts,
      sample.options?.receivedAtMs,
      sample.options?.received_at_ms,
    ),
  };
}

function sampleLabel(sample = {}, index = 0) {
  if (!sample || typeof sample !== 'object' || Array.isArray(sample)) return `sample_${index + 1}`;
  return firstNonEmpty(sample.label, sample.name, sample.id, sample.event_id, sample.eventId, `sample_${index + 1}`);
}

function diagnoseSamples(platform, samples = [], options = {}) {
  return samples.map((sample, index) => {
    const diagnostic = diagnosePlatformEvent(
      platform,
      samplePayload(sample),
      sampleOptions(sample, options),
    );
    return compactObject({
      label: sampleLabel(sample, index),
      diagnostic,
    });
  });
}

function sampleIssues(sampleReports = []) {
  return sampleReports.flatMap((sample) => (
    (sample.diagnostic.issues ?? []).map((item) => ({
      ...item,
      sample: sample.label,
    }))
  ));
}

function reportIssues({ readiness, sampleReports, missingRequiredCoverage, sampleCount }) {
  const issues = [];
  if (readiness?.ready === false) {
    issues.push(issue(
      readiness.blocking_count > 0 ? 'error' : 'warning',
      'platform_setup_not_ready',
      'Platform setup readiness is not fully ready.',
      { blocking_count: readiness.blocking_count, warning_count: readiness.warning_count },
    ));
  }
  if (sampleCount === 0) {
    issues.push(issue(
      'warning',
      'no_acceptance_samples',
      'No real or fixture platform events were provided for sample acceptance.',
    ));
  }
  if (missingRequiredCoverage.length > 0) {
    issues.push(issue(
      'warning',
      'missing_required_coverage',
      'Sample events do not cover all required timeline capabilities.',
      { missing_coverage: missingRequiredCoverage },
    ));
  }
  return [...issues, ...sampleIssues(sampleReports)];
}

function reportStatus({ readiness, sampleCount, missingRequiredCoverage, sampleReports }) {
  if (readiness?.blocking_count > 0) return 'blocked';
  if (sampleCount === 0) return 'pending_samples';
  if (sampleReports.every((sample) => sample.diagnostic.ok === false)) return 'samples_failed';
  if (missingRequiredCoverage.length > 0) return 'missing_required_coverage';
  return 'accepted';
}

export function buildPlatformAcceptanceReport(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const plan = buildPlatformIntegrationPlan(key, options);
  const samples = sampleCollectionFor(key, options.samples ?? options.sampleEvents ?? options.sample_events);
  const sampleReports = diagnoseSamples(key, samples, options);
  const diagnostics = sampleReports.map((sample) => sample.diagnostic);
  const coverage = mergeCoverage(diagnostics);
  const requiredCoverage = defaultRequiredCoverage(key, options);
  const missingRequiredCoverage = requiredCoverage.filter((item) => coverage[item] !== true);
  const issues = reportIssues({
    readiness: plan.readiness,
    sampleReports,
    missingRequiredCoverage,
    sampleCount: samples.length,
  });
  const status = reportStatus({
    readiness: plan.readiness,
    sampleCount: samples.length,
    missingRequiredCoverage,
    sampleReports,
  });
  return compactObject({
    platform: key,
    display_name: plan.display_name,
    status,
    accepted: status === 'accepted',
    recommended_mode: plan.recommended_mode,
    source_priority: plan.source_priority,
    readiness: plan.readiness,
    subscription_maintenance: plan.subscription_maintenance,
    required_coverage: requiredCoverage,
    missing_required_coverage: missingRequiredCoverage,
    coverage,
    signal_types: mergedSignalTypes(diagnostics),
    sample_count: samples.length,
    actionable_sample_count: diagnostics.filter((item) => item.actionable).length,
    samples: sampleReports,
    issues,
    next_actions: issues.map((item) => item.code),
  });
}

export function buildAllPlatformAcceptanceReports(options = {}) {
  return MEETING_PLATFORM_KEYS.map((platform) => buildPlatformAcceptanceReport(platform, options));
}

export function buildMeetingPlatformAcceptanceSummary(options = {}) {
  const reports = buildAllPlatformAcceptanceReports(options);
  return {
    ok: reports.every((item) => item.accepted),
    accepted_count: reports.filter((item) => item.accepted).length,
    blocked_count: reports.filter((item) => item.status === 'blocked').length,
    pending_samples_count: reports.filter((item) => item.status === 'pending_samples').length,
    missing_required_coverage_count: reports.filter((item) => item.status === 'missing_required_coverage').length,
    reports,
  };
}
