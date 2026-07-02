import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import { meetingAppBrowserRuntimePreset } from './meeting-app-browser-runtime.mjs';
import { meetingAppDomCaptureProfile } from './meeting-app-capture.mjs';
import {
  MEETING_APP_FIXTURE_PLATFORMS,
  diagnoseMeetingAppFixtureLifecycle,
} from './meeting-app-fixtures.mjs';
import { normalizeMeetingAppSnapshot, observeMeetingAppSample } from './meeting-apps.mjs';
import { meetingAppSnapshotsFromRecords } from './meeting-app-snapshot-recorder.mjs';
import { normalizeMeetingPlatform } from './platform-setup.mjs';

export const MEETING_APP_LAUNCH_GATE_EVIDENCE_LEVELS = Object.freeze([
  'captured_dom',
  'fixture_dom',
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

function normalizeAppPlatform(platform) {
  const key = normalizeMeetingPlatform(platform);
  if (!MEETING_APP_FIXTURE_PLATFORMS.includes(key)) {
    throw new MeetingTimelineSdkError(`Unsupported meeting app launch gate platform: ${String(platform || '(empty)')}`, {
      platform,
      supported_platforms: MEETING_APP_FIXTURE_PLATFORMS,
    });
  }
  return key;
}

function sampleCollectionFor(platform, snapshotsInput = undefined) {
  if (snapshotsInput == null) return [];
  if (Array.isArray(snapshotsInput)) {
    return snapshotsInput.filter((snapshot) => {
      const candidate = snapshot?.platform ?? snapshot?.provider ?? snapshot?.capture?.profile;
      if (!candidate) return true;
      try {
        return normalizeMeetingPlatform(candidate) === platform;
      } catch {
        return false;
      }
    });
  }
  if (typeof snapshotsInput !== 'object') return [];
  for (const [key, value] of Object.entries(snapshotsInput)) {
    try {
      if (normalizeMeetingPlatform(key) === platform) return asArray(value);
    } catch {
      // Ignore non-platform keys.
    }
  }
  return [];
}

function defaultRequiredCoverage(options = {}) {
  if (options.requiredCoverage || options.required_coverage) {
    return uniqueList(options.requiredCoverage ?? options.required_coverage);
  }
  const required = [
    'runtime_preset',
    'capture_profile',
    'mutation_observer',
    'mutation_track_selectors',
    'mutation_ignore_selectors',
    'platform_detected',
    'meeting_id',
    'in_meeting',
    'meeting_started',
    'active_speaker',
    'speaker_started',
  ];
  if (options.requireMeetingEnd !== false && options.require_meeting_end !== false) {
    required.push('meeting_ended');
  }
  return required;
}

function observeSnapshots(platform, snapshots = [], options = {}) {
  let state = null;
  const reports = [];
  const signals = [];
  for (const snapshot of snapshots) {
    const normalized = normalizeMeetingAppSnapshot(snapshot, options);
    const observedAtMs = firstNonEmpty(
      snapshot?.observedAtMs,
      snapshot?.observed_at_ms,
      options.observedAtMs,
      options.observed_at_ms,
    );
    const observation = observeMeetingAppSample(state, snapshot, {
      source: options.source ?? 'meeting_app_launch_gate',
      speakerOptions: {
        minStableMs: 0,
        ...(options.speakerOptions ?? {}),
        ...(options.speaker_options ?? {}),
      },
      observedAtMs,
      ...options,
    });
    state = observation.state;
    reports.push({ snapshot, normalized, observation });
    signals.push(...observation.signals);
  }
  const signalTypes = signals.map((signal) => signal.type);
  const normalizedItems = reports.map((report) => report.normalized).filter(Boolean);
  const coverage = {
    platform_detected: normalizedItems.some((item) => item.platform === platform),
    meeting_id: normalizedItems.some((item) => Boolean(item.meeting_id)),
    in_meeting: normalizedItems.some((item) => item.inMeeting === true),
    active_speaker: normalizedItems.some((item) => Boolean(item.activeSpeaker?.id || item.activeSpeaker?.name)),
    meeting_started: signalTypes.includes('meeting_started'),
    speaker_started: signalTypes.includes('speaker_started'),
    meeting_ended: signalTypes.includes('meeting_ended'),
    ended_in_meeting_false: normalizedItems.some((item) => item.inMeeting === false),
  };
  return {
    reports,
    signals,
    signal_types: signalTypes,
    coverage,
  };
}

function runtimeCoverage(platform, options = {}) {
  const runtimePreset = meetingAppBrowserRuntimePreset(platform, options);
  const captureProfile = meetingAppDomCaptureProfile(platform, options);
  return {
    preset: runtimePreset,
    capture_profile: captureProfile,
    coverage: {
      runtime_preset: Boolean(runtimePreset),
      capture_profile: Boolean(captureProfile),
      mutation_observer: runtimePreset?.observeMutations === true || runtimePreset?.observe_mutations === true,
      mutation_track_selectors: (runtimePreset?.mutationTrackSelectors?.length ?? 0) > 0,
      mutation_ignore_selectors: (runtimePreset?.mutationIgnoreSelectors?.length ?? 0) > 0,
    },
    recommended_runtime: runtimePreset ? {
      runtimePreset: platform,
      captureOptions: runtimePreset.captureOptions,
      observeMutations: runtimePreset.observeMutations,
      mutationDebounceMs: runtimePreset.mutationDebounceMs,
      speakerStableFollowupMs: runtimePreset.speakerStableFollowupMs,
      sampleIntervalMs: runtimePreset.sampleIntervalMs,
      unchangedObserveEveryMs: runtimePreset.unchangedObserveEveryMs,
      mutation_track_selector_count: runtimePreset.mutationTrackSelectors?.length ?? 0,
      mutation_ignore_selector_count: runtimePreset.mutationIgnoreSelectors?.length ?? 0,
    } : undefined,
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

function issuesForGate({
  evidenceLevel,
  evidenceCount,
  missingCoverage,
  options,
}) {
  const issues = [];
  if (evidenceLevel === 'none') {
    issues.push(issue('error', 'missing_meeting_app_evidence', 'No captured DOM snapshots or fixture evidence were available.'));
  }
  if (evidenceLevel === 'fixture_dom') {
    const severity = options.allowFixtureProduction === true || options.allow_fixture_production === true ? 'warning' : 'error';
    issues.push(issue(
      severity,
      'fixture_dom_only',
      'Fixture DOM proves SDK wiring but does not prove live meeting app DOM compatibility.',
      { evidence_count: evidenceCount },
    ));
  }
  if (missingCoverage.length > 0) {
    issues.push(issue(
      'error',
      'missing_required_coverage',
      'Meeting app launch gate required coverage is missing.',
      { missing_coverage: missingCoverage },
    ));
  }
  return issues;
}

function gateStatus(blockingIssues = [], warnings = []) {
  if (blockingIssues.length > 0) return 'failed';
  if (warnings.length > 0) return 'warning';
  return 'passed';
}

export function buildMeetingAppLaunchGate(platform, options = {}) {
  const key = normalizeAppPlatform(platform);
  const runtime = runtimeCoverage(key, options);
  const snapshotRecords = firstNonEmpty(
    options.snapshotRecords,
    options.snapshot_records,
    options.recordSet,
    options.record_set,
    options.records,
  );
  const capturedSnapshots = [
    ...sampleCollectionFor(key, options.snapshots ?? options.domSnapshots ?? options.dom_snapshots),
    ...sampleCollectionFor(key, snapshotRecords ? meetingAppSnapshotsFromRecords(snapshotRecords, options) : []),
  ];
  const requiredCoverage = defaultRequiredCoverage(options);
  const evidenceLevel = capturedSnapshots.length > 0
    ? 'captured_dom'
    : (options.allowFixtureEvidence === false || options.allow_fixture_evidence === false ? 'none' : 'fixture_dom');
  const fixtureLifecycle = evidenceLevel === 'fixture_dom'
    ? diagnoseMeetingAppFixtureLifecycle(key, {
      ...options,
      speakerOptions: {
        minStableMs: 0,
        ...(options.speakerOptions ?? {}),
        ...(options.speaker_options ?? {}),
      },
    })
    : null;
  const captured = evidenceLevel === 'captured_dom'
    ? observeSnapshots(key, capturedSnapshots, options)
    : null;
  const appCoverage = evidenceLevel === 'captured_dom'
    ? captured.coverage
    : evidenceLevel === 'fixture_dom' ? {
      platform_detected: true,
      meeting_id: true,
      in_meeting: fixtureLifecycle?.coverage.meeting_started === true,
      active_speaker: fixtureLifecycle?.coverage.speaker_started === true,
      meeting_started: fixtureLifecycle?.coverage.meeting_started === true,
      speaker_started: fixtureLifecycle?.coverage.speaker_started === true,
      meeting_ended: fixtureLifecycle?.coverage.meeting_ended === true,
      ended_in_meeting_false: fixtureLifecycle?.coverage.ended_in_meeting_false === true,
    } : {
      platform_detected: false,
      meeting_id: false,
      in_meeting: false,
      active_speaker: false,
      meeting_started: false,
      speaker_started: false,
      meeting_ended: false,
      ended_in_meeting_false: false,
    };
  const coverage = {
    ...runtime.coverage,
    ...appCoverage,
  };
  const missingCoverage = requiredCoverage.filter((item) => coverage[item] !== true);
  const issues = issuesForGate({
    evidenceLevel,
    evidenceCount: capturedSnapshots.length || (fixtureLifecycle ? 2 : 0),
    missingCoverage,
    options,
  });
  const blockingIssues = issues.filter((item) => item.severity === 'error');
  const warnings = issues.filter((item) => item.severity !== 'error');
  const status = gateStatus(blockingIssues, warnings);
  return compactObject({
    type: 'meeting_app_launch_gate',
    platform: key,
    status,
    passed: status !== 'failed',
    production_ready: status === 'passed' && evidenceLevel === 'captured_dom',
    evidence_level: evidenceLevel,
    evidence_count: capturedSnapshots.length || (fixtureLifecycle ? 2 : 0),
    required_coverage: requiredCoverage,
    missing_required_coverage: missingCoverage,
    coverage,
    runtime_ready: runtime.coverage.runtime_preset
      && runtime.coverage.capture_profile
      && runtime.coverage.mutation_observer
      && runtime.coverage.mutation_track_selectors
      && runtime.coverage.mutation_ignore_selectors,
    recommended_runtime: runtime.recommended_runtime,
    blocking_issues: blockingIssues,
    warnings,
    next_actions: uniqueList([
      ...blockingIssues.map((item) => item.code),
      ...warnings.map((item) => item.code),
    ]),
    reports: {
      runtime_preset: runtime.preset,
      capture_profile: runtime.capture_profile,
      fixture_lifecycle: fixtureLifecycle,
      captured,
    },
  });
}

export function buildAllMeetingAppLaunchGates(options = {}) {
  const platforms = options.platforms ?? options.platform_keys ?? MEETING_APP_FIXTURE_PLATFORMS;
  return platforms.map((platform) => buildMeetingAppLaunchGate(platform, options));
}

export function buildMeetingAppLaunchGateSummary(options = {}) {
  const gates = buildAllMeetingAppLaunchGates(options);
  return {
    type: 'meeting_app_launch_gate_summary',
    ok: gates.every((gate) => gate.passed),
    production_ready: gates.every((gate) => gate.production_ready),
    passed_count: gates.filter((gate) => gate.passed).length,
    production_ready_count: gates.filter((gate) => gate.production_ready).length,
    failed_count: gates.filter((gate) => gate.status === 'failed').length,
    warning_count: gates.filter((gate) => gate.status === 'warning').length,
    gates,
  };
}

export function assertMeetingAppLaunchGate(platform, options = {}) {
  const gate = buildMeetingAppLaunchGate(platform, options);
  const requireProductionReady = options.requireProductionReady !== false && options.require_production_ready !== false;
  if (gate.status === 'failed' || (requireProductionReady && gate.production_ready !== true)) {
    throw new MeetingTimelineSdkError(`Meeting app launch gate failed for ${gate.platform}`, {
      gate,
    });
  }
  return gate;
}

export function assertAllMeetingAppLaunchGates(options = {}) {
  const summary = buildMeetingAppLaunchGateSummary(options);
  const requireProductionReady = options.requireProductionReady !== false && options.require_production_ready !== false;
  if (!summary.ok || (requireProductionReady && summary.production_ready !== true)) {
    throw new MeetingTimelineSdkError('Meeting app launch gate summary failed', {
      summary,
    });
  }
  return summary;
}
