import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  buildMeetingPlatformAcceptanceSummary,
  buildPlatformAcceptanceReport,
} from './platform-acceptance.mjs';
import { buildPlatformCaptureSamples } from './platform-capture.mjs';
import {
  buildPlatformFixtureSamples,
} from './platform-fixtures.mjs';
import {
  buildMeetingPlatformOnboardingReport,
  buildMeetingPlatformOnboardingSummary,
} from './platform-onboarding.mjs';
import { MEETING_PLATFORM_KEYS, normalizeMeetingPlatform } from './platform-setup.mjs';

export const PLATFORM_LAUNCH_GATE_EVIDENCE_LEVELS = Object.freeze([
  'captured_events',
  'sample_events',
  'fixture_events',
  'none',
]);

function asArray(value) {
  return Array.isArray(value) ? value : value == null ? [] : [value];
}

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function uniqueList(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
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

function defaultCoverage(platform, options = {}) {
  if (options.requiredCoverage || options.required_coverage) {
    return uniqueList(options.requiredCoverage ?? options.required_coverage);
  }
  const key = normalizeMeetingPlatform(platform);
  const coverage = ['meeting_start'];
  if (options.requireEndEvent !== false && options.require_end_event !== false && key !== 'local_detector') {
    coverage.push('meeting_end');
  }
  if (options.requireParticipants === true || options.require_participants === true) coverage.push('participant_track');
  if (options.requireSpeakerActivity === true || options.require_speaker_activity === true) coverage.push('speaker_activity');
  if (options.requireArtifact === true || options.require_artifact === true) coverage.push('artifact_ready');
  if (options.requireSubscriptionLifecycle === true || options.require_subscription_lifecycle === true) coverage.push('subscription_lifecycle');
  return uniqueList(coverage);
}

function recordsForPlatform(platform, records = []) {
  return asArray(records).filter((record) => {
    try {
      return normalizeMeetingPlatform(record?.platform ?? record?.provider ?? record?.adapter) === platform;
    } catch {
      return false;
    }
  });
}

function resolveGateSamples(platform, options = {}) {
  const records = recordsForPlatform(platform, options.records ?? options.captureRecords ?? options.capture_records);
  if (records.length > 0) {
    return {
      evidence_level: 'captured_events',
      samples: buildPlatformCaptureSamples(records),
      evidence_count: records.length,
    };
  }
  const sampleEvents = options.samples ?? options.sampleEvents ?? options.sample_events;
  const selectedSamples = sampleCollectionFor(platform, sampleEvents);
  if (selectedSamples.length > 0) {
    return {
      evidence_level: 'sample_events',
      samples: { [platform]: selectedSamples },
      evidence_count: selectedSamples.length,
    };
  }
  if (options.allowFixtureEvidence === true || options.allow_fixture_evidence === true) {
    const fixtureSamples = buildPlatformFixtureSamples(platform, options);
    return {
      evidence_level: 'fixture_events',
      samples: { [platform]: fixtureSamples },
      evidence_count: fixtureSamples.length,
    };
  }
  return {
    evidence_level: 'none',
    samples: {},
    evidence_count: 0,
  };
}

function issue(severity, code, message, details = {}) {
  return compactObject({
    severity,
    code,
    message,
    ...details,
  });
}

function gateIssues({ options, evidenceLevel, evidenceCount, onboarding, acceptance, missingCoverage }) {
  const issues = [];
  if (onboarding.permission_plan?.readiness?.blocking_count > 0) {
    issues.push(issue(
      'error',
      'setup_not_ready',
      'Platform setup readiness has blocking checks.',
      { readiness: onboarding.permission_plan.readiness },
    ));
  }
  if (evidenceLevel === 'none') {
    issues.push(issue(
      'error',
      'missing_real_event_evidence',
      'No captured or sample platform events were provided for launch gate.',
    ));
  }
  if (evidenceLevel === 'fixture_events') {
    const severity = options.allowFixtureProduction === true || options.allow_fixture_production === true ? 'warning' : 'error';
    issues.push(issue(
      severity,
      'fixture_evidence_only',
      'Fixture events prove SDK wiring but do not prove provider delivery in this environment.',
      { evidence_count: evidenceCount },
    ));
  }
  if (acceptance.status === 'pending_samples') {
    issues.push(issue('error', 'acceptance_pending_samples', 'Acceptance has no actionable samples.'));
  }
  if (acceptance.status === 'samples_failed') {
    issues.push(issue('error', 'acceptance_samples_failed', 'Acceptance samples failed to normalize into timeline signals.'));
  }
  if (missingCoverage.length > 0) {
    issues.push(issue(
      'error',
      'missing_required_coverage',
      'Launch gate required coverage is missing.',
      { missing_coverage: missingCoverage },
    ));
  }
  if (acceptance.issues?.length) {
    for (const item of acceptance.issues) {
      const severity = item.severity === 'error' ? 'error' : 'warning';
      issues.push(issue(severity, `acceptance:${item.code}`, item.message, { sample: item.sample }));
    }
  }
  return issues;
}

function gateStatus(blockingIssues = [], warnings = []) {
  if (blockingIssues.length > 0) return 'failed';
  if (warnings.length > 0) return 'warning';
  return 'passed';
}

export function buildPlatformLaunchGate(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const requiredCoverage = defaultCoverage(key, options);
  const resolved = resolveGateSamples(key, options);
  const acceptance = buildPlatformAcceptanceReport(key, {
    ...options,
    samples: resolved.samples,
    requireEndEvent: requiredCoverage.includes('meeting_end'),
    requiredCoverage,
  });
  const onboarding = buildMeetingPlatformOnboardingReport(key, {
    ...options,
    samples: resolved.samples,
    requireEndEvent: requiredCoverage.includes('meeting_end'),
    requiredCoverage,
  });
  const missingCoverage = requiredCoverage.filter((coverage) => acceptance.coverage?.[coverage] !== true);
  const issues = gateIssues({
    options,
    evidenceLevel: resolved.evidence_level,
    evidenceCount: resolved.evidence_count,
    onboarding,
    acceptance,
    missingCoverage,
  });
  const blockingIssues = issues.filter((item) => item.severity === 'error');
  const warnings = issues.filter((item) => item.severity !== 'error');
  const status = gateStatus(blockingIssues, warnings);
  return compactObject({
    platform: key,
    status,
    passed: status !== 'failed',
    production_ready: status === 'passed' && resolved.evidence_level !== 'fixture_events',
    evidence_level: resolved.evidence_level,
    evidence_count: resolved.evidence_count,
    required_coverage: requiredCoverage,
    missing_required_coverage: missingCoverage,
    coverage: acceptance.coverage,
    readiness: onboarding.permission_plan?.readiness,
    onboarding_status: onboarding.status,
    acceptance_status: acceptance.status,
    sample_count: acceptance.sample_count,
    actionable_sample_count: acceptance.actionable_sample_count,
    blocking_issues: blockingIssues,
    warnings,
    next_actions: uniqueList([
      ...blockingIssues.map((item) => item.code),
      ...warnings.map((item) => item.code),
      ...(onboarding.next_actions ?? []),
    ]),
    reports: {
      acceptance,
      onboarding,
    },
  });
}

export function buildAllPlatformLaunchGates(options = {}) {
  return MEETING_PLATFORM_KEYS.map((platform) => buildPlatformLaunchGate(platform, options));
}

export function buildMeetingPlatformLaunchGateSummary(options = {}) {
  const gates = buildAllPlatformLaunchGates(options);
  return {
    ok: gates.every((gate) => gate.passed),
    production_ready: gates.every((gate) => gate.production_ready),
    passed_count: gates.filter((gate) => gate.passed).length,
    production_ready_count: gates.filter((gate) => gate.production_ready).length,
    failed_count: gates.filter((gate) => gate.status === 'failed').length,
    warning_count: gates.filter((gate) => gate.status === 'warning').length,
    gates,
  };
}

export function assertPlatformLaunchGate(platform, options = {}) {
  const gate = buildPlatformLaunchGate(platform, options);
  const requireProductionReady = options.requireProductionReady !== false && options.require_production_ready !== false;
  if (gate.status === 'failed' || (requireProductionReady && gate.production_ready !== true)) {
    throw new MeetingTimelineSdkError(`Meeting platform launch gate failed for ${gate.platform}`, {
      platform: gate.platform,
      status: gate.status,
      production_ready: gate.production_ready,
      evidence_level: gate.evidence_level,
      blocking_issues: gate.blocking_issues,
      warnings: gate.warnings,
      next_actions: gate.next_actions,
    });
  }
  return gate;
}

export function assertAllPlatformLaunchGates(options = {}) {
  const summary = buildMeetingPlatformLaunchGateSummary(options);
  const requireProductionReady = options.requireProductionReady !== false && options.require_production_ready !== false;
  if (summary.ok !== true || (requireProductionReady && summary.production_ready !== true)) {
    throw new MeetingTimelineSdkError('Meeting platform launch gates failed', {
      ok: summary.ok,
      production_ready: summary.production_ready,
      failed_count: summary.failed_count,
      warning_count: summary.warning_count,
      gates: summary.gates.map((gate) => ({
        platform: gate.platform,
        status: gate.status,
        production_ready: gate.production_ready,
        evidence_level: gate.evidence_level,
        blocking_issues: gate.blocking_issues,
        warnings: gate.warnings,
        next_actions: gate.next_actions,
      })),
    });
  }
  return summary;
}

export function buildFixtureLaunchGateInput(options = {}) {
  return {
    ...options,
    allowFixtureEvidence: true,
  };
}
