import { compactObject } from '../index.mjs';
import { MEETING_APP_FIXTURE_PLATFORMS } from './meeting-app-fixtures.mjs';
import { buildMeetingAppLiveSnapshotCapturePlan } from './meeting-app-profile.mjs';
import {
  buildMeetingPlatformEvidencePackageSummary,
} from './platform-evidence-package.mjs';
import { buildMeetingPlatformProviderConnectionPack } from './platform-provider-connection.mjs';
import {
  MEETING_PLATFORM_KEYS,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';
import { buildMeetingPlatformRuntimeProfile } from './platform-runtime-profile.mjs';

export const MEETING_PLATFORM_FIELD_CAPTURE_PLAN_SCHEMA = 'meeting_platform_field_capture_plan';
export const MEETING_PLATFORM_FIELD_CAPTURE_MATRIX_SCHEMA = 'meeting_platform_field_capture_matrix';
export const MEETING_PLATFORM_FIELD_CAPTURE_SCHEMA_VERSION = 1;

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

function selectedPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_PLATFORM_KEYS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

function matchesPlatform(value, platform) {
  if (!value) return true;
  try {
    return normalizeMeetingPlatform(value) === platform;
  } catch {
    return false;
  }
}

function selectedEvidencePackage(platform, options = {}) {
  const input = firstNonEmpty(
    options.evidencePackage,
    options.evidence_package,
    options.package,
    options.platformEvidencePackage,
    options.platform_evidence_package,
  );
  if (!input) return undefined;
  if (Array.isArray(input)) {
    return input.find((item) => matchesPlatform(item?.platform ?? item?.rollout_plan?.platform, platform));
  }
  if (input.schema === 'meeting_platform_evidence_package' || input.rollout_plan || input.evidence_summary) {
    return matchesPlatform(input.platform ?? input.rollout_plan?.platform, platform) ? input : undefined;
  }
  if (typeof input === 'object') {
    for (const [key, value] of Object.entries(input)) {
      if (matchesPlatform(key, platform)) return value;
    }
  }
  return undefined;
}

function outputPaths(platform, options = {}) {
  const root = String(firstNonEmpty(options.evidenceDir, options.evidence_dir, 'data'));
  return {
    meeting_app_evidence: `${root}/meeting-app-evidence/${platform}.json`,
    provider_evidence: platform === 'local_detector'
      ? `${root}/local-detector-evidence/${platform}.json`
      : `${root}/provider-evidence/${platform}.json`,
    evidence_package: `${root}/meeting-platform-evidence-packages/${platform}.json`,
    capture_report: `${root}/meeting-platform-field-capture/${platform}.json`,
  };
}

function requiredProviderCoverage(platform, runtime = {}) {
  if (platform === 'local_detector') return ['meeting_start', 'meeting_end'];
  return [
    runtime.axis?.start?.provider_reconcile_events?.length > 0 ? 'meeting_start' : undefined,
    runtime.axis?.end?.provider_reconcile_events?.length > 0 ? 'meeting_end' : undefined,
  ].filter(Boolean);
}

function providerEventPlan(platform, runtime = {}, provider = {}) {
  const coverage = requiredProviderCoverage(platform, runtime);
  return {
    source: platform === 'local_detector' ? 'trusted_host_detector' : provider.transport,
    endpoint: provider.endpoint,
    status_endpoint: provider.status_endpoint,
    required_coverage: coverage,
    minimum_record_count: coverage.length,
    start_events: runtime.axis?.start?.provider_reconcile_events ?? [],
    end_events: runtime.axis?.end?.provider_reconcile_events ?? [],
    participant_events: runtime.provider_events?.participant_events ?? [],
    artifact_events: runtime.provider_events?.artifact_events ?? [],
    lifecycle_events: runtime.provider_events?.lifecycle_events ?? [],
    security: {
      verifier: provider.security?.verifier,
      missing_env: provider.security?.missing_env ?? [],
      required_env: provider.security?.required_env ?? [],
    },
  };
}

function localSnapshotPlan(platform, options = {}) {
  if (!MEETING_APP_FIXTURE_PLATFORMS.includes(platform)) return undefined;
  const plan = buildMeetingAppLiveSnapshotCapturePlan(platform, options);
  return {
    source: 'browser_extension_or_native_host_observer',
    required_snapshots: plan.required_snapshots,
    recommended_snapshots: plan.recommended_snapshots,
    minimum_record_count: plan.minimum_record_count,
    validation: plan.validation,
    runtime_config: plan.runtime_config,
    recorder: plan.recorder,
  };
}

function summaryForPackage(platform, options = {}) {
  const pkg = selectedEvidencePackage(platform, options);
  return pkg ? buildMeetingPlatformEvidencePackageSummary(pkg, options) : undefined;
}

function missingLocalItems(localPlan = undefined, summary = undefined) {
  if (!localPlan) return [];
  if (!summary) {
    return [
      'capture_live_dom_snapshots_for_local_observer',
      ...localPlan.required_snapshots.map((item) => `capture_required_snapshot:${item.id}`),
    ];
  }
  const missing = [];
  if (Number(summary.meeting_app_record_count ?? 0) < Number(localPlan.minimum_record_count ?? 0)) {
    missing.push('capture_live_dom_snapshots_for_local_observer');
  }
  for (const item of summary.local_dom_missing_required_coverage ?? []) {
    missing.push(`local_dom_missing:${item}`);
  }
  return unique(missing);
}

function missingProviderItems(providerPlan = {}, summary = undefined) {
  if (!summary) {
    return [
      'capture_real_provider_start_end_events',
      ...providerPlan.required_coverage.map((item) => `provider_missing:${item}`),
    ];
  }
  const missing = [];
  if (Number(summary.provider_record_count ?? 0) < Number(providerPlan.minimum_record_count ?? 0)) {
    missing.push('capture_real_provider_start_end_events');
  }
  for (const item of summary.provider_missing_required_coverage ?? []) {
    missing.push(`provider_missing:${item}`);
  }
  return unique(missing);
}

function fieldStatus(summary = undefined, localMissing = [], providerMissing = []) {
  if (summary?.production_ready === true) return 'production_ready';
  if (summary?.ready_for_realtime_annotations === true && providerMissing.length > 0) return 'pilot_ready_provider_pending';
  if (localMissing.length === 0 && providerMissing.length > 0) return 'needs_provider_events';
  if (localMissing.length > 0 && providerMissing.length === 0) return 'needs_local_observer_snapshots';
  return 'needs_local_and_provider_evidence';
}

function captureChecklist(platform, localPlan = undefined, providerPlan = {}, runtime = {}) {
  const items = [];
  if (localPlan) {
    items.push({
      id: 'capture_active_speaker_snapshot',
      channel: 'meeting_app_observer',
      when: 'after_joining_the_real_meeting_and_a_speaker_is_visibly_active',
      output: 'meeting_app_snapshot_record',
      required_snapshot: 'active_speaker',
    });
    if ((localPlan.required_snapshots ?? []).some((item) => item.id === 'meeting_ended')) {
      items.push({
        id: 'capture_meeting_ended_snapshot',
        channel: 'meeting_app_observer',
        when: 'after_leaving_or_ending_the_same_real_meeting',
        output: 'meeting_app_snapshot_record',
        required_snapshot: 'meeting_ended',
      });
    }
    items.push({
      id: 'insert_live_annotation_sample',
      channel: 'timeline_sdk',
      when: 'while_the_local_axis_is_active',
      output: 'annotation_mark_with_captured_at_ms',
      invariant: runtime.runtime_contract?.annotation_timestamp_field,
    });
  }
  items.push({
    id: platform === 'local_detector' ? 'capture_detector_start_end' : 'capture_provider_start_end',
    channel: platform === 'local_detector' ? 'trusted_host_detector' : 'provider_webhook',
    when: platform === 'local_detector'
      ? 'when_the_host_detector_reports_meeting_started_and_meeting_ended'
      : 'when_the_provider_delivers_real_start_and_end_events_for_the_same_meeting',
    output: 'platform_capture_records',
    required_coverage: providerPlan.required_coverage,
  });
  return items;
}

export function buildMeetingPlatformFieldCapturePlan(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const runtime = buildMeetingPlatformRuntimeProfile(key, options);
  const provider = buildMeetingPlatformProviderConnectionPack(key, options);
  const localPlan = localSnapshotPlan(key, options);
  const providerPlan = providerEventPlan(key, runtime, provider);
  const evidenceSummary = summaryForPackage(key, options);
  const localMissing = missingLocalItems(localPlan, evidenceSummary);
  const providerMissing = missingProviderItems(providerPlan, evidenceSummary);
  const status = fieldStatus(evidenceSummary, localMissing, providerMissing);
  return compactObject({
    type: 'meeting_platform_field_capture_plan',
    schema: MEETING_PLATFORM_FIELD_CAPTURE_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_FIELD_CAPTURE_SCHEMA_VERSION,
    platform: key,
    display_name: runtime.display_name,
    status,
    production_ready: status === 'production_ready',
    ready_for_realtime_annotations: evidenceSummary?.ready_for_realtime_annotations === true,
    objective: 'capture_real_meeting_evidence_for_timeline_annotation_rollout',
    output_paths: outputPaths(key, options),
    runtime_contract: runtime.runtime_contract,
    local_observer: localPlan,
    provider_events: providerPlan,
    checklist: captureChecklist(key, localPlan, providerPlan, runtime),
    current_evidence: evidenceSummary,
    missing_items: unique([
      ...localMissing,
      ...providerMissing,
    ]),
    validation: {
      build_package: 'buildMeetingPlatformEvidencePackage(platform, { providerRecords, meetingAppRecordSet, env, baseUrl })',
      verify_package: 'verifyMeetingPlatformEvidencePackage(evidencePackage)',
      production_condition: 'evidencePackage.rollout_plan.production_ready === true',
      pilot_condition: 'evidencePackage.rollout_plan.ready_for_realtime_annotations === true',
    },
    next_actions: status === 'production_ready'
      ? ['ship_with_monitoring_and_keep_collecting_regression_evidence']
      : unique([
        ...localMissing,
        ...providerMissing,
        'export_meeting_platform_evidence_package',
        'verify_meeting_platform_evidence_package',
      ]),
  });
}

export function buildMeetingPlatformFieldCaptureMatrix(options = {}) {
  const plans = selectedPlatforms(options).map((platform) => buildMeetingPlatformFieldCapturePlan(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_field_capture_matrix',
    schema: MEETING_PLATFORM_FIELD_CAPTURE_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_FIELD_CAPTURE_SCHEMA_VERSION,
    platform_count: plans.length,
    production_ready_count: plans.filter((plan) => plan.production_ready).length,
    realtime_ready_count: plans.filter((plan) => plan.ready_for_realtime_annotations).length,
    missing_item_count: plans.reduce((total, plan) => total + (plan.missing_items?.length ?? 0), 0),
    platforms: plans.map((plan) => plan.platform),
    rows: plans.map((plan) => ({
      platform: plan.platform,
      display_name: plan.display_name,
      status: plan.status,
      production_ready: plan.production_ready,
      ready_for_realtime_annotations: plan.ready_for_realtime_annotations,
      local_required_records: plan.local_observer?.minimum_record_count ?? 0,
      provider_required_records: plan.provider_events?.minimum_record_count ?? 0,
      missing_items: plan.missing_items,
    })),
    plans,
    next_actions: unique(plans.flatMap((plan) => plan.next_actions ?? [])),
  };
}
