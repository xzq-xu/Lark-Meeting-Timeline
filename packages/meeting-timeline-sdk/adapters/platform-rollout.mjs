import { compactObject } from '../index.mjs';
import { MEETING_APP_FIXTURE_PLATFORMS } from './meeting-app-fixtures.mjs';
import { buildMeetingAppLaunchGate } from './meeting-app-gate.mjs';
import { buildMeetingAppLiveSnapshotCapturePlan } from './meeting-app-profile.mjs';
import { buildPlatformLaunchGate } from './platform-gate.mjs';
import {
  MEETING_PLATFORM_KEYS,
  buildPlatformIntegrationPlan,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_ROLLOUT_STATUSES = Object.freeze([
  'production_ready',
  'realtime_ready_provider_pending',
  'provider_ready_collect_local_evidence',
  'provider_ready_local_blocked',
  'needs_live_dom_and_provider_evidence',
  'needs_live_dom_evidence',
  'needs_provider_event_evidence',
  'needs_detector_event_evidence',
  'blocked',
]);

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  return Array.isArray(value) ? value : value == null ? [] : [value];
}

function uniqueList(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function issueCodes(gate = {}) {
  const item = gate ?? {};
  return uniqueList([
    ...(item.blocking_issues ?? []).map((issue) => issue.code),
    ...(item.warnings ?? []).map((issue) => issue.code),
    ...(item.next_actions ?? []),
  ]);
}

function providerGateOptions(options = {}) {
  return compactObject({
    ...options,
    records: firstNonEmpty(
      options.providerRecords,
      options.provider_records,
      options.providerCaptureRecords,
      options.provider_capture_records,
      options.captureRecords,
      options.capture_records,
      options.records,
    ),
    samples: firstNonEmpty(
      options.providerSamples,
      options.provider_samples,
      options.sampleEvents,
      options.sample_events,
      options.samples,
    ),
    allowFixtureEvidence: firstNonEmpty(
      options.allowProviderFixtureEvidence,
      options.allow_provider_fixture_evidence,
      options.allowFixtureEvidence,
      options.allow_fixture_evidence,
      false,
    ),
    allowFixtureProduction: firstNonEmpty(
      options.allowProviderFixtureProduction,
      options.allow_provider_fixture_production,
      options.allowFixtureProduction,
      options.allow_fixture_production,
      false,
    ),
  });
}

function meetingAppGateOptions(options = {}) {
  return compactObject({
    ...options,
    snapshots: firstNonEmpty(
      options.meetingAppSnapshots,
      options.meeting_app_snapshots,
      options.domSnapshots,
      options.dom_snapshots,
      options.snapshots,
    ),
    snapshotRecords: firstNonEmpty(
      options.meetingAppSnapshotRecords,
      options.meeting_app_snapshot_records,
      options.meetingAppRecords,
      options.meeting_app_records,
      options.snapshotRecords,
      options.snapshot_records,
    ),
    recordSet: firstNonEmpty(
      options.meetingAppRecordSet,
      options.meeting_app_record_set,
      options.recordSet,
      options.record_set,
    ),
    allowFixtureEvidence: firstNonEmpty(
      options.allowMeetingAppFixtureEvidence,
      options.allow_meeting_app_fixture_evidence,
      options.allowFixtureEvidence,
      options.allow_fixture_evidence,
      false,
    ),
    allowFixtureProduction: firstNonEmpty(
      options.allowMeetingAppFixtureProduction,
      options.allow_meeting_app_fixture_production,
      options.allowFixtureProduction,
      options.allow_fixture_production,
      false,
    ),
  });
}

function localObserverGate(platform, options = {}) {
  if (!MEETING_APP_FIXTURE_PLATFORMS.includes(platform)) return null;
  return buildMeetingAppLaunchGate(platform, meetingAppGateOptions(options));
}

function providerStatus(gate = {}) {
  if (gate.production_ready) return 'ready';
  if (gate.readiness?.ready === true) return 'setup_ready_missing_evidence';
  return 'not_ready';
}

function rolloutStatus(platform, localGate, providerGate) {
  const providerReady = providerGate?.production_ready === true;
  const providerSetupReady = providerGate?.readiness?.ready === true;
  if (platform === 'local_detector') {
    if (providerReady) return 'production_ready';
    if (providerSetupReady) return 'needs_detector_event_evidence';
    return 'blocked';
  }
  const localReady = localGate?.production_ready === true;
  const localRuntimeReady = localGate?.runtime_ready === true;
  if (localReady && providerReady) return 'production_ready';
  if (localReady) return 'realtime_ready_provider_pending';
  if (providerReady && localRuntimeReady) return 'provider_ready_collect_local_evidence';
  if (providerReady) return 'provider_ready_local_blocked';
  if (localRuntimeReady && providerSetupReady) return 'needs_live_dom_and_provider_evidence';
  if (localRuntimeReady) return 'needs_live_dom_evidence';
  if (providerSetupReady) return 'needs_provider_event_evidence';
  return 'blocked';
}

function recommendedMode(status) {
  if (status === 'production_ready' || status === 'realtime_ready_provider_pending') {
    return 'local_observer_realtime_with_provider_reconcile';
  }
  if (status === 'provider_ready_collect_local_evidence' || status === 'provider_ready_local_blocked') {
    return 'provider_events_realtime_until_local_observer_is_verified';
  }
  if (status === 'needs_live_dom_and_provider_evidence' || status === 'needs_live_dom_evidence') {
    return 'capture_live_dom_snapshots_first';
  }
  if (status === 'needs_provider_event_evidence') return 'capture_provider_events';
  if (status === 'needs_detector_event_evidence') return 'capture_local_detector_events';
  return 'fix_setup_before_rollout';
}

function highLevelActions(platform, status, localGate, providerGate) {
  const actions = [];
  if (platform !== 'local_detector' && localGate?.production_ready !== true) {
    actions.push('capture_live_dom_snapshots_for_local_observer');
    if (localGate?.runtime_ready !== true) actions.push('fix_meeting_app_runtime_or_capture_profile');
  }
  if (providerGate?.production_ready !== true) {
    actions.push(platform === 'local_detector'
      ? 'capture_local_detector_start_end_events'
      : 'capture_real_provider_start_end_events');
  }
  if (providerGate?.readiness?.ready === false) actions.push('complete_provider_setup_readiness');
  if (status === 'production_ready') actions.push('enable_pilot_rollout_with_monitoring');
  return actions;
}

function evidenceFileName(platform, suffix) {
  return `data/${suffix}/${platform}.json`;
}

function captureCommands(platform) {
  return {
    build_extension: 'npm run meeting-app:extension:build',
    validate_dom_evidence: `npm run meeting-app:evidence-gate -- --input=${evidenceFileName(platform, 'meeting-app-evidence')} --report-file=data/meeting-app-live-gate-report.json`,
    validate_dom_matrix: 'npm run meeting-app:evidence-matrix',
    validate_provider_events: 'buildPlatformLaunchGate(platform, { records: providerRecords, env, baseUrl })',
    validate_rollout: 'buildMeetingPlatformRolloutPlan(platform, { providerRecords, meetingAppRecordSet, env, baseUrl })',
  };
}

function localDomRunbook(platform, options = {}) {
  if (!MEETING_APP_FIXTURE_PLATFORMS.includes(platform)) return undefined;
  const capturePlan = buildMeetingAppLiveSnapshotCapturePlan(platform, options);
  return {
    objective: 'prove_low_latency_local_observer_for_realtime_annotation_axis',
    evidence_file: evidenceFileName(platform, 'meeting-app-evidence'),
    capture_plan: capturePlan,
    required_snapshots: capturePlan.required_snapshots,
    recommended_snapshots: capturePlan.recommended_snapshots,
    minimum_record_count: capturePlan.minimum_record_count,
    capture_api: [
      'window.__meetingTimelineLiveCapture.captureActive()',
      'window.__meetingTimelineLiveCapture.captureEnded()',
      'window.__meetingTimelineLiveCapture.exportRecords()',
      'window.__meetingTimelineLiveCapture.evidencePackage()',
      'window.__meetingTimelineLiveCapture.diagnose()',
    ],
    validation: capturePlan.validation,
  };
}

function providerRunbook(platform, integration = {}, providerGate = {}) {
  const providerEvents = integration.provider_events ?? {};
  const requiredCoverage = providerGate.required_coverage ?? ['meeting_start', 'meeting_end'];
  return {
    objective: platform === 'local_detector'
      ? 'prove_host_detector_can_emit_start_end_with_absolute_time'
      : 'prove_provider_events_can_reconcile_or_backfill_the_local_axis',
    evidence_file: platform === 'local_detector'
      ? evidenceFileName(platform, 'local-detector-evidence')
      : evidenceFileName(platform, 'provider-evidence'),
    transport: providerEvents.transport,
    endpoint: providerEvents.endpoint,
    status_endpoint: providerEvents.status_endpoint,
    event_types: providerEvents.event_types ?? [],
    lifecycle_event_types: providerEvents.lifecycle_event_types ?? [],
    required_coverage: requiredCoverage,
    current_missing_coverage: providerGate.missing_required_coverage ?? requiredCoverage,
    validation: {
      adapter: '@ai-annotation/meeting-timeline-sdk/adapters/platform-gate',
      method: 'buildPlatformLaunchGate',
      production_ready_requires: 'captured_events',
      default_options: {
        allowFixtureEvidence: false,
      },
    },
  };
}

function runbookSteps(platform, plan, localDom, provider) {
  const steps = [];
  if (localDom) {
    steps.push({
      id: 'install_capture_runtime',
      title: 'Install the browser or host capture runtime',
      output: 'content_script_or_host_observer_running_on_real_meeting_surface',
      command: captureCommands(platform).build_extension,
    });
    steps.push({
      id: 'capture_local_dom_active_and_ended',
      title: 'Capture active-speaker and ended DOM snapshots from the same real meeting',
      output: localDom.evidence_file,
      required_snapshot_ids: localDom.required_snapshots.map((item) => item.id),
    });
    steps.push({
      id: 'validate_local_dom_gate',
      title: 'Validate local DOM evidence for realtime annotation readiness',
      output: 'meeting_app_launch_gate.production_ready',
      command: captureCommands(platform).validate_dom_evidence,
    });
  }
  steps.push({
    id: platform === 'local_detector' ? 'capture_detector_events' : 'capture_provider_events',
    title: platform === 'local_detector'
      ? 'Capture host detector start/end events with absolute timestamps'
      : 'Capture real provider start/end events from the configured webhook path',
    output: provider.evidence_file,
    required_coverage: provider.required_coverage,
  });
  steps.push({
    id: 'validate_provider_gate',
    title: 'Validate provider or detector event evidence',
    output: 'platform_launch_gate.production_ready',
    command: captureCommands(platform).validate_provider_events,
  });
  steps.push({
    id: 'validate_rollout',
    title: 'Combine local DOM and provider evidence into a rollout decision',
    output: 'meeting_platform_rollout_plan.status',
    command: captureCommands(platform).validate_rollout,
    success_condition: plan.production_ready
      ? 'status === "production_ready"'
      : 'ready_for_realtime_annotations === true for pilot, production_ready === true for full rollout',
  });
  return steps;
}

export function buildMeetingPlatformRolloutPlan(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const integration = buildPlatformIntegrationPlan(key, options);
  const providerGate = buildPlatformLaunchGate(key, providerGateOptions(options));
  const localGate = localObserverGate(key, options);
  const status = rolloutStatus(key, localGate, providerGate);
  const localReady = localGate?.production_ready === true;
  const providerReady = providerGate.production_ready === true;
  const readyForRealtime = key === 'local_detector'
    ? providerReady
    : localReady || providerReady;
  return compactObject({
    type: 'meeting_platform_rollout_plan',
    platform: key,
    display_name: integration.display_name,
    status,
    recommended_mode: recommendedMode(status),
    production_ready: status === 'production_ready',
    ready_for_realtime_annotations: readyForRealtime,
    realtime_axis_primary: integration.realtime_axis?.primary,
    source_priority: integration.source_priority,
    local_observer: key === 'local_detector' ? undefined : {
      supported: Boolean(localGate),
      status: localGate?.status ?? 'unsupported',
      runtime_ready: localGate?.runtime_ready ?? false,
      production_ready: localReady,
      evidence_level: localGate?.evidence_level ?? 'none',
      evidence_count: localGate?.evidence_count ?? 0,
      missing_required_coverage: localGate?.missing_required_coverage ?? ['captured_dom_evidence'],
      issue_codes: issueCodes(localGate),
      gate: localGate,
    },
    provider_events: {
      status: providerStatus(providerGate),
      gate_status: providerGate.status,
      setup_ready: providerGate.readiness?.ready ?? false,
      production_ready: providerReady,
      evidence_level: providerGate.evidence_level,
      evidence_count: providerGate.evidence_count,
      missing_required_coverage: providerGate.missing_required_coverage ?? [],
      issue_codes: issueCodes(providerGate),
      event_types: integration.provider_events?.event_types ?? [],
      endpoint: integration.provider_events?.endpoint,
      gate: providerGate,
    },
    post_meeting_transcript: integration.post_meeting_transcript,
    required_timestamp_invariant: integration.realtime_axis?.invariant,
    next_actions: uniqueList([
      ...highLevelActions(key, status, localGate, providerGate),
      ...issueCodes(localGate),
      ...issueCodes(providerGate),
    ]),
    integration_plan: integration,
  });
}

export function buildAllMeetingPlatformRolloutPlans(options = {}) {
  const platforms = options.platforms ?? options.platform_keys ?? MEETING_PLATFORM_KEYS;
  return asArray(platforms).map((platform) => buildMeetingPlatformRolloutPlan(platform, options));
}

export function buildMeetingPlatformRolloutSummary(options = {}) {
  const plans = buildAllMeetingPlatformRolloutPlans(options);
  return {
    type: 'meeting_platform_rollout_summary',
    ok: plans.every((plan) => plan.ready_for_realtime_annotations),
    production_ready: plans.every((plan) => plan.production_ready),
    plan_count: plans.length,
    production_ready_count: plans.filter((plan) => plan.production_ready).length,
    realtime_ready_count: plans.filter((plan) => plan.ready_for_realtime_annotations).length,
    blocked_count: plans.filter((plan) => plan.status === 'blocked').length,
    pending_platforms: plans
      .filter((plan) => plan.production_ready !== true)
      .map((plan) => plan.platform),
    next_actions: uniqueList(plans.flatMap((plan) => plan.next_actions ?? [])),
    plans,
  };
}

export function buildMeetingPlatformAdaptationRunbook(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const rolloutPlan = buildMeetingPlatformRolloutPlan(key, options);
  const integration = rolloutPlan.integration_plan ?? buildPlatformIntegrationPlan(key, options);
  const localDom = localDomRunbook(key, options);
  const provider = providerRunbook(key, integration, rolloutPlan.provider_events?.gate);
  const commands = captureCommands(key);
  return compactObject({
    type: 'meeting_platform_adaptation_runbook',
    platform: key,
    display_name: rolloutPlan.display_name,
    rollout_status: rolloutPlan.status,
    recommended_mode: rolloutPlan.recommended_mode,
    production_ready: rolloutPlan.production_ready,
    ready_for_realtime_annotations: rolloutPlan.ready_for_realtime_annotations,
    local_dom: localDom,
    provider_events: provider,
    steps: runbookSteps(key, rolloutPlan, localDom, provider),
    commands,
    handoff: {
      dom_evidence_input: 'meetingAppRecordSet or meetingAppSnapshots',
      provider_evidence_input: 'providerRecords or providerSamples',
      final_gate: 'buildMeetingPlatformRolloutPlan',
      pilot_condition: 'ready_for_realtime_annotations === true',
      production_condition: 'production_ready === true',
    },
    next_actions: rolloutPlan.next_actions,
    rollout_plan: rolloutPlan,
  });
}

export function buildAllMeetingPlatformAdaptationRunbooks(options = {}) {
  const platforms = options.platforms ?? options.platform_keys ?? MEETING_PLATFORM_KEYS;
  return asArray(platforms).map((platform) => buildMeetingPlatformAdaptationRunbook(platform, options));
}

export function buildMeetingPlatformAdaptationRunbookSummary(options = {}) {
  const runbooks = buildAllMeetingPlatformAdaptationRunbooks(options);
  return {
    type: 'meeting_platform_adaptation_runbook_summary',
    runbook_count: runbooks.length,
    production_ready_count: runbooks.filter((item) => item.production_ready).length,
    realtime_ready_count: runbooks.filter((item) => item.ready_for_realtime_annotations).length,
    platforms: runbooks.map((item) => item.platform),
    next_actions: uniqueList(runbooks.flatMap((item) => item.next_actions ?? [])),
    runbooks,
  };
}
