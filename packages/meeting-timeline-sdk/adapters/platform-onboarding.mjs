import { compactObject } from '../index.mjs';
import { buildArtifactImportPlans } from './artifact-plan.mjs';
import {
  buildMeetingPlatformAcceptanceSummary,
  buildPlatformAcceptanceReport,
} from './platform-acceptance.mjs';
import {
  MEETING_PLATFORM_KEYS,
  buildPlatformIntegrationPlan,
  buildPlatformPermissionPlan,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

function asArray(value) {
  return Array.isArray(value) ? value : value == null ? [] : [value];
}

function uniqueList(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function acceptanceOptions(options = {}) {
  return {
    ...options,
    diagnosticOptions: {
      ...(options.diagnosticOptions ?? options.diagnostic_options ?? {}),
      includeRawSignals: true,
    },
  };
}

function rawSignalsFromAcceptance(report = {}) {
  return (report.samples ?? []).flatMap((sample) => sample.diagnostic?.raw_signals ?? []);
}

function selectedArtifactSignals(acceptance = {}, options = {}) {
  return [
    ...asArray(options.artifactSignals ?? options.artifact_signals),
    ...rawSignalsFromAcceptance(acceptance),
  ].filter((signal) => signal?.type === 'artifact_ready');
}

function setupNextActions(permissionPlan = {}) {
  const missingEnv = permissionPlan.missing_security_env ?? [];
  const readinessChecks = permissionPlan.readiness?.checks ?? [];
  return [
    ...missingEnv.map((name) => `configure_env:${name}`),
    ...readinessChecks
      .filter((check) => check.ok !== true && check.severity === 'error')
      .map((check) => `fix_setup:${check.id}`),
  ];
}

function acceptanceNextActions(acceptance = {}) {
  if (acceptance.sample_count === 0) return ['provide_real_event_samples'];
  return acceptance.next_actions ?? [];
}

function artifactNextActions(artifactPlans = []) {
  return artifactPlans.map((plan) => {
    if (plan.action === 'fetch_and_import_transcript') return `artifact:${plan.platform}:fetch_transcript`;
    if (plan.action === 'store_recording_artifact') return `artifact:${plan.platform}:store_recording`;
    if (plan.action === 'fetch_and_store_smart_notes') return `artifact:${plan.platform}:fetch_smart_notes`;
    return `artifact:${plan.platform}:store_metadata`;
  });
}

function issue(severity, code, message, details = {}) {
  return compactObject({ severity, code, message, ...details });
}

function candidateObservationContract(platform, options = {}) {
  const override = options.candidateObservation ?? options.candidate_observation ?? {};
  return compactObject({
    platform,
    ready: override.ready ?? override.candidate_observation_ready ?? true,
    message_type: override.message_type ?? 'meeting_timeline.observe_candidates',
    required_permission: override.required_permission ?? 'tabs',
    runtime_event_action: override.runtime_event_action ?? 'observe_platform_candidates',
    runtime_event_client_method: override.runtime_event_client_method ?? 'observePlatformCandidates',
    endpoint: override.endpoint ?? '/api/meeting-platform/observe-candidates',
    producer: override.producer ?? 'browser_extension_background_or_native_host',
  });
}

function candidateObservationGate(candidateObservation = {}) {
  const issues = [];
  if (candidateObservation.ready !== true) {
    issues.push(issue('error', 'candidate_observation_not_ready', 'Candidate observation must be ready for host-level realtime axis binding.'));
  }
  if (candidateObservation.message_type !== 'meeting_timeline.observe_candidates') {
    issues.push(issue('error', 'candidate_observation_invalid_message_type', 'Candidate observation must use meeting_timeline.observe_candidates.'));
  }
  if (candidateObservation.required_permission !== 'tabs') {
    issues.push(issue('error', 'candidate_observation_missing_tabs_permission', 'Candidate observation must declare the tabs permission.'));
  }
  if (candidateObservation.runtime_event_action !== 'observe_platform_candidates') {
    issues.push(issue('error', 'candidate_observation_invalid_runtime_action', 'Candidate observation must route to observe_platform_candidates.'));
  }
  if (!candidateObservation.endpoint) {
    issues.push(issue('error', 'candidate_observation_missing_endpoint', 'Candidate observation must expose a host endpoint.'));
  }
  const blocking = issues.filter((item) => item.severity === 'error');
  return {
    accepted: blocking.length === 0,
    blocking_count: blocking.length,
    warning_count: issues.length - blocking.length,
    issues,
    candidate_observation: candidateObservation,
  };
}

function candidateObservationNextActions(gate = {}) {
  if (gate.accepted === true) return [];
  return (gate.issues ?? [])
    .filter((issue) => issue.severity === 'error')
    .map((issue) => `fix_candidate_observation:${issue.code}`);
}

function onboardingStatus(permissionPlan = {}, acceptance = {}, candidateGate = {}) {
  if (permissionPlan.readiness?.blocking_count > 0) return 'blocked_by_setup';
  if (candidateGate.accepted !== true) return 'blocked_by_runtime_contract';
  if (acceptance.accepted) return 'ready';
  if (acceptance.status === 'pending_samples') return 'needs_real_samples';
  if (acceptance.status === 'missing_required_coverage') return 'needs_more_coverage';
  if (acceptance.status === 'samples_failed') return 'samples_failed';
  return acceptance.status ?? 'unknown';
}

function runtimeContract(integrationPlan = {}, candidateObservation = {}) {
  return compactObject({
    mode: integrationPlan.recommended_mode,
    source_priority: integrationPlan.source_priority,
    annotation_time_field: integrationPlan.realtime_annotations?.required_field ?? 'captured_at_ms',
    primary_axis_source: integrationPlan.realtime_axis?.primary,
    candidate_observation_required_for_host_axis_binding: true,
    candidate_observation_ready: candidateObservation.ready === true,
    candidate_observation: candidateObservation,
    provider_events_enabled: integrationPlan.provider_events?.enabled,
    provider_event_endpoint: integrationPlan.provider_events?.endpoint,
    transcript_import_strategy: integrationPlan.post_meeting_transcript?.strategy,
    speaker_activity_strategy: integrationPlan.speaker_activity?.strategy,
    modules: integrationPlan.modules,
  });
}

export function buildMeetingPlatformOnboardingReport(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const permissionPlan = buildPlatformPermissionPlan(key, options);
  const integrationPlan = buildPlatformIntegrationPlan(key, options);
  const candidateObservation = candidateObservationContract(key, options);
  const candidateGate = candidateObservationGate(candidateObservation);
  const acceptance = buildPlatformAcceptanceReport(key, acceptanceOptions(options));
  const artifactPlans = buildArtifactImportPlans(selectedArtifactSignals(acceptance, options), {
    importEndpoint: options.importEndpoint ?? options.import_endpoint,
  });
  const nextActions = uniqueList([
    ...setupNextActions(permissionPlan),
    ...candidateObservationNextActions(candidateGate),
    ...acceptanceNextActions(acceptance),
    ...artifactNextActions(artifactPlans),
  ]);
  return compactObject({
    platform: key,
    display_name: integrationPlan.display_name,
    status: onboardingStatus(permissionPlan, acceptance, candidateGate),
    ready: permissionPlan.readiness?.ready === true
      && candidateGate.accepted === true
      && acceptance.accepted === true,
    recommended_mode: integrationPlan.recommended_mode,
    runtime_contract: runtimeContract(integrationPlan, candidateObservation),
    permission_plan: permissionPlan,
    integration_plan: integrationPlan,
    candidate_observation: candidateObservation,
    candidate_observation_gate: candidateGate,
    acceptance,
    artifact_import_plans: artifactPlans,
    next_actions: nextActions,
  });
}

export function buildAllMeetingPlatformOnboardingReports(options = {}) {
  return MEETING_PLATFORM_KEYS.map((platform) => buildMeetingPlatformOnboardingReport(platform, options));
}

export function buildMeetingPlatformOnboardingSummary(options = {}) {
  const reports = buildAllMeetingPlatformOnboardingReports(options);
  const acceptanceSummary = buildMeetingPlatformAcceptanceSummary(acceptanceOptions(options));
  return {
    ok: reports.every((item) => item.ready),
    ready_count: reports.filter((item) => item.ready).length,
    blocked_count: reports.filter((item) => item.status === 'blocked_by_setup').length,
    blocked_by_runtime_contract_count: reports.filter((item) => item.status === 'blocked_by_runtime_contract').length,
    needs_real_samples_count: reports.filter((item) => item.status === 'needs_real_samples').length,
    needs_more_coverage_count: reports.filter((item) => item.status === 'needs_more_coverage').length,
    samples_failed_count: reports.filter((item) => item.status === 'samples_failed').length,
    acceptance_summary: acceptanceSummary,
    reports,
  };
}
