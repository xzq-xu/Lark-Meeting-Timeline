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

function onboardingStatus(permissionPlan = {}, acceptance = {}) {
  if (permissionPlan.readiness?.blocking_count > 0) return 'blocked_by_setup';
  if (acceptance.accepted) return 'ready';
  if (acceptance.status === 'pending_samples') return 'needs_real_samples';
  if (acceptance.status === 'missing_required_coverage') return 'needs_more_coverage';
  if (acceptance.status === 'samples_failed') return 'samples_failed';
  return acceptance.status ?? 'unknown';
}

function runtimeContract(integrationPlan = {}) {
  return compactObject({
    mode: integrationPlan.recommended_mode,
    source_priority: integrationPlan.source_priority,
    annotation_time_field: integrationPlan.realtime_annotations?.required_field ?? 'captured_at_ms',
    primary_axis_source: integrationPlan.realtime_axis?.primary,
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
  const acceptance = buildPlatformAcceptanceReport(key, acceptanceOptions(options));
  const artifactPlans = buildArtifactImportPlans(selectedArtifactSignals(acceptance, options), {
    importEndpoint: options.importEndpoint ?? options.import_endpoint,
  });
  const nextActions = uniqueList([
    ...setupNextActions(permissionPlan),
    ...acceptanceNextActions(acceptance),
    ...artifactNextActions(artifactPlans),
  ]);
  return compactObject({
    platform: key,
    display_name: integrationPlan.display_name,
    status: onboardingStatus(permissionPlan, acceptance),
    ready: permissionPlan.readiness?.ready === true && acceptance.accepted === true,
    recommended_mode: integrationPlan.recommended_mode,
    runtime_contract: runtimeContract(integrationPlan),
    permission_plan: permissionPlan,
    integration_plan: integrationPlan,
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
    needs_real_samples_count: reports.filter((item) => item.status === 'needs_real_samples').length,
    needs_more_coverage_count: reports.filter((item) => item.status === 'needs_more_coverage').length,
    samples_failed_count: reports.filter((item) => item.status === 'samples_failed').length,
    acceptance_summary: acceptanceSummary,
    reports,
  };
}
