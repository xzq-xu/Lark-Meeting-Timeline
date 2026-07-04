import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  buildMeetingPlatformEvidencePackage,
  buildMeetingPlatformEvidencePackageSummary,
  verifyMeetingPlatformEvidencePackage,
} from './platform-evidence-package.mjs';
import {
  runMeetingPlatformRuntimeHostReplay,
} from './meeting-platform-runtime-host-verifier.mjs';
import {
  buildMeetingPlatformLiveAdapterReadiness,
} from './platform-live-adapter.mjs';
import {
  normalizeMeetingPlatform,
  platformEventEndpoint,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_REAL_INTAKE_SCHEMA = 'meeting_platform_real_evidence_intake';
export const MEETING_PLATFORM_REAL_INTAKE_MATRIX_SCHEMA = 'meeting_platform_real_evidence_intake_matrix';
export const MEETING_PLATFORM_REAL_INTAKE_PLAN_SCHEMA = 'meeting_platform_real_evidence_intake_plan';
export const MEETING_PLATFORM_REAL_INTAKE_SCHEMA_VERSION = 1;

const DEFAULT_REAL_INTAKE_PLATFORMS = Object.freeze([
  'google_meet',
  'microsoft_teams',
  'zoom',
  'webex',
  'lark',
]);

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
  return unique(asArray(firstNonEmpty(
    options.platforms,
    options.platform_keys,
    DEFAULT_REAL_INTAKE_PLATFORMS,
  )).map((platform) => normalizeMeetingPlatform(platform)));
}

function issue(severity, code, message, details = {}) {
  return compactObject({ severity, code, message, ...details });
}

function includesFixtureMarker(value) {
  return String(value ?? '').toLowerCase().includes('fixture');
}

function evidenceStrings(record = {}) {
  const snapshot = record.snapshot ?? record;
  return [
    record.id,
    record.label,
    record.source,
    record.fixture_state,
    record.provider,
    snapshot.id,
    snapshot.label,
    snapshot.source,
    snapshot.fixture_state,
    snapshot.title,
    record.body?.id,
    record.body?.event_id,
    record.body?.event,
    record.body?.type,
    record.body?.data?.conferenceRecord?.name,
    record.body?.payload?.object?.uuid,
    record.body?.resourceData?.onlineMeetingId,
    record.body?.data?.meetingId,
  ];
}

function fixtureEvidence(records = []) {
  return records.filter((record) => evidenceStrings(record).some((value) => includesFixtureMarker(value)));
}

function providerRecordsFrom(input = {}, options = {}) {
  return [
    ...asArray(input.providerRecords),
    ...asArray(input.provider_records),
    ...asArray(input.providerCaptureRecords),
    ...asArray(input.provider_capture_records),
    ...asArray(input.captureRecords),
    ...asArray(input.capture_records),
    ...asArray(options.providerRecords),
    ...asArray(options.provider_records),
    ...asArray(options.providerCaptureRecords),
    ...asArray(options.provider_capture_records),
  ];
}

function meetingAppRecordsFrom(input = {}, options = {}) {
  return [
    ...asArray(input.meetingAppRecords),
    ...asArray(input.meeting_app_records),
    ...asArray(input.meetingAppSnapshotRecords),
    ...asArray(input.meeting_app_snapshot_records),
    ...asArray(input.meetingAppSnapshots),
    ...asArray(input.meeting_app_snapshots),
    ...asArray(input.domSnapshots),
    ...asArray(input.dom_snapshots),
    ...asArray(input.snapshots),
    ...asArray(options.meetingAppRecords),
    ...asArray(options.meeting_app_records),
    ...asArray(options.meetingAppSnapshotRecords),
    ...asArray(options.meeting_app_snapshot_records),
    ...asArray(options.meetingAppSnapshots),
    ...asArray(options.meeting_app_snapshots),
    ...asArray(options.domSnapshots),
    ...asArray(options.dom_snapshots),
  ];
}

function evidencePackageFrom(platform, input = {}, options = {}) {
  const explicit = firstNonEmpty(input.evidencePackage, input.evidence_package, input.package, options.evidencePackage, options.evidence_package);
  if (explicit?.schema === 'meeting_platform_evidence_package') return explicit;
  return buildMeetingPlatformEvidencePackage(platform, input, options);
}

function platformInput(input = {}, platform) {
  if (!input || typeof input !== 'object' || Array.isArray(input)) return {};
  const key = normalizeMeetingPlatform(platform);
  return firstNonEmpty(input[key], input.platforms?.[key], input.byPlatform?.[key], input.by_platform?.[key], input) ?? {};
}

function check(code, passed, severity, message, details = {}) {
  return compactObject({
    code,
    passed: Boolean(passed),
    severity: passed ? 'info' : severity,
    message,
    ...details,
  });
}

export function buildMeetingPlatformRealEvidenceIntakePlan(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const baseUrl = String(firstNonEmpty(options.baseUrl, options.base_url, 'http://localhost:8787'));
  return {
    type: 'meeting_platform_real_evidence_intake_plan',
    schema: MEETING_PLATFORM_REAL_INTAKE_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_REAL_INTAKE_SCHEMA_VERSION,
    platform: key,
    base_url: baseUrl,
    provider_endpoint: platformEventEndpoint(baseUrl, key),
    required_inputs: [
      'provider_capture_records_from_real_webhook_or_event_subscription',
      'meeting_app_dom_records_from_real_browser_or_native_observer',
      'captured_at_ms_on_every_record',
    ],
    forbidden_inputs: [
      'meeting_app_fixture_source',
      'fixture_provider_payload_ids',
      'synthetic_sample_only_evidence',
    ],
    output_contract: {
      evidence_package_schema: 'meeting_platform_evidence_package',
      readiness_schema: 'meeting_platform_live_adapter_readiness',
      verification_method: 'verifyMeetingPlatformEvidencePackage',
      production_condition: 'verification.passed === true && readiness.passed === true && no_fixture_evidence',
    },
  };
}

export function buildMeetingPlatformRealEvidenceIntakeReport(platform, input = {}, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const baseUrl = String(firstNonEmpty(options.baseUrl, options.base_url, input.baseUrl, input.base_url, 'http://localhost:8787'));
  const mergedOptions = {
    ...options,
    baseUrl,
    includeRunbook: firstNonEmpty(options.includeRunbook, options.include_runbook, false),
  };
  const providerRecords = providerRecordsFrom(input, options);
  const meetingAppRecords = meetingAppRecordsFrom(input, options);
  const records = [...providerRecords, ...meetingAppRecords];
  const fixtureRecords = fixtureEvidence(records);
  const allowFixtureEvidence = options.allowFixtureEvidence === true || options.allow_fixture_evidence === true;
  const evidencePackage = evidencePackageFrom(key, input, {
    ...mergedOptions,
    providerRecords,
    meetingAppRecords,
  });
  const verification = verifyMeetingPlatformEvidencePackage(evidencePackage, {
    ...mergedOptions,
    requireProductionReady: firstNonEmpty(options.requireProductionReady, options.require_production_ready, true),
    includePackage: false,
  });
  const readiness = buildMeetingPlatformLiveAdapterReadiness(key, {
    ...mergedOptions,
    target: firstNonEmpty(options.target, 'production'),
    evidencePackage,
  });
  const summary = buildMeetingPlatformEvidencePackageSummary(evidencePackage);
  const checks = [
    check(
      'provider_records_present',
      key === 'local_detector' || providerRecords.length > 0,
      'error',
      'Real platform intake requires captured provider records for reconcile evidence.',
      { provider_record_count: providerRecords.length },
    ),
    check(
      'meeting_app_records_present',
      meetingAppRecords.length > 0,
      'error',
      'Real platform intake requires captured meeting app DOM/native observer records.',
      { meeting_app_record_count: meetingAppRecords.length },
    ),
    check(
      'no_fixture_evidence',
      allowFixtureEvidence || fixtureRecords.length === 0,
      'error',
      'Real intake must not be accepted from fixture or synthetic sample evidence.',
      { fixture_evidence_count: fixtureRecords.length },
    ),
    check(
      'evidence_package_verified',
      verification.passed === true,
      'error',
      'Evidence package must verify against production readiness gates.',
      { verification_requirement: verification.requirement },
    ),
    check(
      'live_adapter_readiness_passed',
      readiness.passed === true,
      'error',
      'Live adapter readiness must pass with the provided evidence package.',
      { readiness_status: readiness.status },
    ),
  ];
  const blocking = checks.filter((item) => item.passed !== true && item.severity === 'error');
  return compactObject({
    type: 'meeting_platform_real_evidence_intake',
    schema: MEETING_PLATFORM_REAL_INTAKE_SCHEMA,
    schema_version: MEETING_PLATFORM_REAL_INTAKE_SCHEMA_VERSION,
    platform: key,
    accepted: blocking.length === 0,
    base_url: baseUrl,
    provider_record_count: providerRecords.length,
    meeting_app_record_count: meetingAppRecords.length,
    fixture_evidence_count: fixtureRecords.length,
    blocking_count: blocking.length,
    checks,
    blocking_checks: blocking,
    summary,
    verification,
    readiness,
    plan: buildMeetingPlatformRealEvidenceIntakePlan(key, { ...options, baseUrl }),
    evidence_package: options.includeEvidencePackage === true || options.include_evidence_package === true
      ? evidencePackage
      : undefined,
    issues: [
      ...blocking.map((item) => issue('error', item.code, item.message, item)),
    ],
    next_actions: blocking.length > 0
      ? unique(blocking.map((item) => item.code))
      : ['handoff_evidence_package_to_host_project', 'enable_pilot_with_monitoring'],
  });
}

function replayReportFrom(input = {}, options = {}) {
  const explicit = firstNonEmpty(
    input.runtimeHostReplayReport,
    input.runtime_host_replay_report,
    input.runtimeHostReplay,
    input.runtime_host_replay,
    options.runtimeHostReplayReport,
    options.runtime_host_replay_report,
    options.runtimeHostReplay,
    options.runtime_host_replay,
  );
  return explicit?.schema === 'meeting_platform_runtime_host_replay' ? explicit : undefined;
}

function replayInputFrom(input = {}, evidencePackage = undefined, options = {}) {
  return firstNonEmpty(
    options.runtimeHostReplayInput,
    options.runtime_host_replay_input,
    input.runtimeHostReplayInput,
    input.runtime_host_replay_input,
    input.replayInput,
    input.replay_input,
    evidencePackage,
    input,
  );
}

function replayErrorReport(platform, error) {
  return {
    type: 'meeting_platform_runtime_host_replay',
    schema: 'meeting_platform_runtime_host_replay',
    version: 1,
    platform,
    accepted: false,
    input_record_count: 0,
    replay_row_count: 0,
    call_count: 0,
    actions: [],
    signal_types: [],
    result_actions: [],
    coverage: {
      runtime_host_replay_error_free: false,
    },
    missing: ['runtime_host_replay_error_free'],
    error: String(error?.message ?? error),
    next_actions: ['fix_runtime_host_replay_input_or_adapter_error'],
  };
}

function mergeRuntimeReplay(report = {}, runtimeHostReplay = {}, options = {}) {
  const includeEvidencePackage = options.includeEvidencePackage === true || options.include_evidence_package === true;
  const replayCheck = check(
    'runtime_host_replay_accepted',
    runtimeHostReplay.accepted === true,
    'error',
    'Runtime host replay must write start, speaker mark, and end into the timeline from captured meeting app snapshots.',
    {
      replay_missing: runtimeHostReplay.missing ?? [],
      replay_row_count: runtimeHostReplay.replay_row_count ?? 0,
    },
  );
  const checks = [...(report.checks ?? []), replayCheck];
  const blocking = checks.filter((item) => item.passed !== true && item.severity === 'error');
  return compactObject({
    ...report,
    accepted: blocking.length === 0,
    blocking_count: blocking.length,
    checks,
    blocking_checks: blocking,
    runtime_host_replay: runtimeHostReplay,
    runtime_host_replay_accepted: runtimeHostReplay.accepted === true,
    evidence_package: includeEvidencePackage ? report.evidence_package : undefined,
    issues: [
      ...blocking.map((item) => issue('error', item.code, item.message, item)),
    ],
    next_actions: blocking.length > 0
      ? unique([
        ...blocking.map((item) => item.code),
        ...(runtimeHostReplay.next_actions ?? []),
      ])
      : ['handoff_evidence_package_to_host_project', 'enable_pilot_with_monitoring'],
  });
}

export async function runMeetingPlatformRealEvidenceIntakeReport(platform, input = {}, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const report = buildMeetingPlatformRealEvidenceIntakeReport(key, input, {
    ...options,
    includeEvidencePackage: true,
    include_evidence_package: true,
  });
  const explicitReplay = replayReportFrom(input, options);
  let runtimeHostReplay = explicitReplay;
  if (!runtimeHostReplay) {
    try {
      runtimeHostReplay = await runMeetingPlatformRuntimeHostReplay(
        key,
        replayInputFrom(input, report.evidence_package, options),
        options.runtimeHostReplayOptions ?? options.runtime_host_replay_options ?? options,
      );
    } catch (error) {
      runtimeHostReplay = replayErrorReport(key, error);
    }
  }
  return mergeRuntimeReplay(report, runtimeHostReplay, options);
}

function matrixFromReports(reports = []) {
  return {
    type: 'meeting_platform_real_evidence_intake_matrix',
    schema: MEETING_PLATFORM_REAL_INTAKE_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_REAL_INTAKE_SCHEMA_VERSION,
    platform_count: reports.length,
    accepted_count: reports.filter((report) => report.accepted).length,
    rejected_count: reports.filter((report) => !report.accepted).length,
    production_ready_count: reports.filter((report) => report.readiness?.production_ready === true).length,
    realtime_ready_count: reports.filter((report) => report.readiness?.ready_for_realtime_annotations === true).length,
    runtime_host_replay_ready_count: reports.filter((report) => report.runtime_host_replay_accepted === true).length,
    platforms: reports.map((report) => report.platform),
    rows: reports.map((report) => ({
      platform: report.platform,
      accepted: report.accepted,
      provider_record_count: report.provider_record_count,
      meeting_app_record_count: report.meeting_app_record_count,
      fixture_evidence_count: report.fixture_evidence_count,
      runtime_host_replay_accepted: report.runtime_host_replay_accepted,
      runtime_host_replay_missing: report.runtime_host_replay?.missing ?? [],
      readiness_status: report.readiness?.status,
      verification_passed: report.verification?.passed,
      blocking_count: report.blocking_count,
      next_actions: report.next_actions,
    })),
    reports,
  };
}

export function buildMeetingPlatformRealEvidenceIntakeMatrix(input = {}, options = {}) {
  const platforms = selectedPlatforms({ ...input, ...options });
  const reports = platforms.map((platform) => buildMeetingPlatformRealEvidenceIntakeReport(
    platform,
    platformInput(input, platform),
    options,
  ));
  return {
    ...matrixFromReports(reports),
    platforms,
  };
}

export async function runMeetingPlatformRealEvidenceIntakeMatrix(input = {}, options = {}) {
  const platforms = selectedPlatforms({ ...input, ...options });
  const reports = [];
  for (const platform of platforms) {
    reports.push(await runMeetingPlatformRealEvidenceIntakeReport(
      platform,
      platformInput(input, platform),
      options,
    ));
  }
  return {
    ...matrixFromReports(reports),
    platforms,
  };
}

export function assertMeetingPlatformRealEvidenceIntake(platform, input = {}, options = {}) {
  const report = buildMeetingPlatformRealEvidenceIntakeReport(platform, input, options);
  if (report.accepted !== true) {
    throw new MeetingTimelineSdkError(`Meeting platform real evidence intake failed for ${report.platform}`, {
      platform: report.platform,
      blocking_checks: report.blocking_checks,
      next_actions: report.next_actions,
      report,
    });
  }
  return report;
}

export function assertMeetingPlatformRealEvidenceIntakeMatrix(input = {}, options = {}) {
  const matrix = buildMeetingPlatformRealEvidenceIntakeMatrix(input, options);
  if (matrix.rejected_count > 0) {
    throw new MeetingTimelineSdkError('Meeting platform real evidence intake matrix failed', {
      platform_count: matrix.platform_count,
      accepted_count: matrix.accepted_count,
      rejected_count: matrix.rejected_count,
      failed_platforms: matrix.rows.filter((row) => !row.accepted).map((row) => row.platform),
      matrix,
    });
  }
  return matrix;
}
