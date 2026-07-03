import { compactObject } from '../index.mjs';
import { MEETING_APP_FIXTURE_PLATFORMS } from './meeting-app-fixtures.mjs';
import { buildMeetingAppLaunchGate } from './meeting-app-gate.mjs';
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
