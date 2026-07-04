import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import { buildMeetingAppDomAdaptationDiagnosis } from './meeting-app-profile.mjs';
import { meetingAppSnapshotRecords } from './meeting-app-snapshot-recorder.mjs';
import {
  buildMeetingPlatformAdapterContract,
  buildMeetingPlatformAdapterContractAcceptanceReport,
} from './platform-adapter-contract.mjs';
import { buildMeetingPlatformFieldIntakePlan } from './platform-field-intake.mjs';
import {
  buildMeetingPlatformLiveAdapterHandoff,
  buildMeetingPlatformLiveAdapterReadiness,
} from './platform-live-adapter.mjs';
import { buildMeetingPlatformProviderConnectionPack } from './platform-provider-connection.mjs';
import {
  buildMeetingPlatformRealEvidenceIntakeReport,
} from './platform-real-intake.mjs';
import { normalizeMeetingPlatform } from './platform-setup.mjs';

export const MEETING_PLATFORM_HANDOFF_READINESS_SCHEMA = 'meeting_platform_handoff_readiness';
export const MEETING_PLATFORM_HANDOFF_READINESS_MATRIX_SCHEMA = 'meeting_platform_handoff_readiness_matrix';
export const MEETING_PLATFORM_HANDOFF_READINESS_SCHEMA_VERSION = 1;

const DEFAULT_HANDOFF_PLATFORMS = Object.freeze([
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
    DEFAULT_HANDOFF_PLATFORMS,
  )).map((platform) => normalizeMeetingPlatform(platform)));
}

function matchesPlatform(value, platform) {
  if (!value) return true;
  try {
    return normalizeMeetingPlatform(value) === platform;
  } catch {
    return false;
  }
}

function packageFromInput(input = {}, platform) {
  const explicit = firstNonEmpty(
    input.evidencePackage,
    input.evidence_package,
    input.package,
    input.handoffPackage,
    input.handoff_package,
  );
  if (!explicit) return undefined;
  if (Array.isArray(explicit)) {
    return explicit.find((item) => matchesPlatform(item?.platform ?? item?.rollout_plan?.platform, platform));
  }
  if (explicit.schema === 'meeting_platform_evidence_package') {
    return matchesPlatform(explicit.platform ?? explicit.rollout_plan?.platform, platform) ? explicit : undefined;
  }
  if (typeof explicit === 'object') {
    for (const [key, value] of Object.entries(explicit)) {
      if (matchesPlatform(key, platform)) return value;
    }
  }
  return undefined;
}

function platformInput(input = {}, platform) {
  if (!input || typeof input !== 'object' || Array.isArray(input)) return {};
  return firstNonEmpty(
    input[platform],
    input.platforms?.[platform],
    input.byPlatform?.[platform],
    input.by_platform?.[platform],
    input,
  ) ?? {};
}

function providerRecordsFrom(input = {}, evidencePackage = {}) {
  return [
    ...asArray(input.providerRecords),
    ...asArray(input.provider_records),
    ...asArray(input.providerCaptureRecords),
    ...asArray(input.provider_capture_records),
    ...asArray(evidencePackage.provider_records),
    ...asArray(evidencePackage.providerRecords),
  ];
}

function meetingAppRecordsFrom(input = {}, evidencePackage = {}) {
  return [
    ...asArray(input.meetingAppRecords),
    ...asArray(input.meeting_app_records),
    ...asArray(input.meetingAppSnapshotRecords),
    ...asArray(input.meeting_app_snapshot_records),
    ...asArray(input.domSnapshots),
    ...asArray(input.dom_snapshots),
    ...asArray(input.snapshots),
    ...meetingAppSnapshotRecords(firstNonEmpty(
      input.meetingAppRecordSet,
      input.meeting_app_record_set,
      input.recordSet,
      input.record_set,
      evidencePackage.meeting_app_record_set,
      evidencePackage.meetingAppRecordSet,
    )),
  ];
}

function evidenceFor(platform, input = {}, options = {}) {
  const rawInput = platformInput(input, platform);
  const merged = {
    ...rawInput,
    ...options,
  };
  const evidencePackage = firstNonEmpty(
    packageFromInput(rawInput, platform),
    packageFromInput(options, platform),
  );
  const providerRecords = providerRecordsFrom(merged, evidencePackage);
  const meetingAppRecords = meetingAppRecordsFrom(merged, evidencePackage);
  return compactObject({
    evidencePackage,
    providerRecords,
    meetingAppRecords,
    records: meetingAppRecords,
    snapshots: firstNonEmpty(merged.snapshots, merged.domSnapshots, merged.dom_snapshots),
    recordSet: firstNonEmpty(
      merged.recordSet,
      merged.record_set,
      merged.meetingAppRecordSet,
      merged.meeting_app_record_set,
      evidencePackage?.meeting_app_record_set,
      evidencePackage?.meetingAppRecordSet,
    ),
  });
}

function blockingCodes(items = []) {
  return unique(asArray(items).map((item) => item.code).filter(Boolean));
}

function statusFor({
  productionReady,
  pilotReady,
  contractAccepted,
  providerReady,
  providerMissingEnv,
  candidateObservationReady,
  domAccepted,
  realAccepted,
} = {}) {
  if (productionReady) return 'production_ready';
  if (pilotReady && providerReady) return 'pilot_ready_provider_reconcile_pending';
  if (pilotReady) return 'pilot_ready_provider_setup_pending';
  if (!candidateObservationReady) return 'needs_candidate_observation_contract';
  if (!contractAccepted) return 'adapter_contract_blocked';
  if (!domAccepted) return 'needs_local_observer_evidence';
  if (providerMissingEnv.length > 0) return 'needs_provider_credentials';
  if (!realAccepted) return 'needs_provider_or_package_evidence';
  return 'needs_field_validation';
}

function handoffCommands(platform, fieldPlan = {}) {
  const key = normalizeMeetingPlatform(platform);
  return {
    export_handoff_readiness: `npm run meeting-platform:handoff-readiness -- --platforms=${key} --json=true`,
    export_field_intake_plan: fieldPlan.commands?.export_field_intake_plan,
    collect_dom_evidence: 'npm run meeting-app:evidence-matrix',
    build_field_evidence: fieldPlan.commands?.build_field_evidence,
    validate_candidate_observation: `npm run meeting-platform:runtime-event-plan -- --platforms=${key} --json=true`,
    validate_real_intake: fieldPlan.commands?.validate_real_intake,
    validate_live_readiness: fieldPlan.commands?.validate_live_readiness,
  };
}

function nextActions({
  status,
  contractReport,
  provider,
  domDiagnosis,
  fieldPlan,
  realIntake,
  liveReadiness,
  candidateObservation,
} = {}) {
  const statusAction = {
    production_ready: 'handoff_to_host_project_with_monitoring',
    pilot_ready_provider_reconcile_pending: 'start_local_observer_pilot_and_capture_provider_reconcile_evidence',
    pilot_ready_provider_setup_pending: 'start_local_observer_pilot_and_complete_provider_setup',
    adapter_contract_blocked: 'fix_adapter_contract_acceptance',
    needs_candidate_observation_contract: 'verify_candidate_observation_before_host_handoff',
    needs_local_observer_evidence: 'capture_real_meeting_app_snapshots',
    needs_provider_credentials: 'configure_provider_security_env',
    needs_provider_or_package_evidence: 'capture_real_provider_events_and_build_evidence_package',
    needs_field_validation: 'run_field_validation_commands',
  }[status];
  return unique([
    statusAction,
    ...(contractReport?.issues ?? []).map((item) => item.code),
    ...(provider?.security?.missing_env ?? []).map((name) => `configure_env:${name}`),
    ...(domDiagnosis?.next_actions ?? []),
    ...(fieldPlan?.next_actions ?? []),
    ...(realIntake?.next_actions ?? []),
    ...(liveReadiness?.next_actions ?? []),
    ...(candidateObservation?.ready === false ? ['verify_candidate_observation_before_host_handoff'] : []),
    ...((candidateObservation?.issues ?? []).map((code) => `candidate_observation:${code}`)),
  ]);
}

export function buildMeetingPlatformHandoffReadiness(platformOrInput = {}, input = {}, options = {}) {
  const objectInput = platformOrInput && typeof platformOrInput === 'object' && !Array.isArray(platformOrInput);
  const rawPlatform = objectInput
    ? firstNonEmpty(platformOrInput.platform, platformOrInput.provider, platformOrInput.key, platformOrInput.name)
    : platformOrInput;
  const platform = normalizeMeetingPlatform(rawPlatform);
  const rawInput = objectInput ? platformOrInput : input;
  const mergedOptions = objectInput ? { ...input, ...options } : { ...input, ...options };
  const evidence = evidenceFor(platform, rawInput, mergedOptions);
  const baseOptions = {
    ...mergedOptions,
    env: {
      ...(mergedOptions.env ?? {}),
    },
  };

  const contract = buildMeetingPlatformAdapterContract(platform, baseOptions);
  const contractReport = buildMeetingPlatformAdapterContractAcceptanceReport(contract, {
    ...baseOptions,
    target: firstNonEmpty(baseOptions.contractTarget, baseOptions.contract_target, 'contract'),
  });
  const provider = buildMeetingPlatformProviderConnectionPack(platform, baseOptions);
  const fieldPlan = buildMeetingPlatformFieldIntakePlan(platform, baseOptions);
  const domDiagnosis = buildMeetingAppDomAdaptationDiagnosis({
    platform,
    ...evidence,
  }, baseOptions);
  const realIntake = buildMeetingPlatformRealEvidenceIntakeReport(platform, evidence, {
    ...baseOptions,
    requireProductionReady: firstNonEmpty(baseOptions.requireProductionReady, baseOptions.require_production_ready, false),
    require_production_ready: firstNonEmpty(baseOptions.requireProductionReady, baseOptions.require_production_ready, false),
  });
  const liveReadiness = buildMeetingPlatformLiveAdapterReadiness(platform, {
    ...baseOptions,
    target: firstNonEmpty(baseOptions.target, baseOptions.readinessTarget, baseOptions.readiness_target, 'pilot'),
    evidencePackage: evidence.evidencePackage,
  });
  const liveHandoff = buildMeetingPlatformLiveAdapterHandoff(platform, {
    ...baseOptions,
    target: firstNonEmpty(baseOptions.target, baseOptions.readinessTarget, baseOptions.readiness_target, 'pilot'),
    evidencePackage: evidence.evidencePackage,
  });
  const providerMissingEnv = provider.security?.missing_env ?? [];
  const providerReady = provider.readiness?.ready === true;
  const contractAccepted = contractReport.accepted === true;
  const domAccepted = domDiagnosis.accepted === true;
  const realAccepted = realIntake.accepted === true;
  const candidateObservation = liveReadiness.candidate_observation ?? liveHandoff.candidate_observation_contract ?? {};
  const candidateObservationReady = candidateObservation.ready === true;
  const pilotReady = contractAccepted && candidateObservationReady && domAccepted;
  const productionReady = pilotReady && realAccepted && liveReadiness.production_ready === true;
  const handoffReady = contractAccepted && candidateObservationReady && (pilotReady || liveReadiness.passed === true);
  const status = statusFor({
    productionReady,
    pilotReady,
    contractAccepted,
    providerReady,
    providerMissingEnv,
    candidateObservationReady,
    domAccepted,
    realAccepted,
  });

  return compactObject({
    type: 'meeting_platform_handoff_readiness',
    schema: MEETING_PLATFORM_HANDOFF_READINESS_SCHEMA,
    schema_version: MEETING_PLATFORM_HANDOFF_READINESS_SCHEMA_VERSION,
    platform,
    display_name: contract.display_name ?? fieldPlan.display_name ?? provider.display_name,
    status,
    handoff_ready: handoffReady,
    pilot_ready: pilotReady,
    production_ready: productionReady,
    provider_reconcile_ready: providerReady && realIntake.provider_record_count > 0,
    local_observer_ready: domAccepted,
    candidate_observation_ready: candidateObservationReady,
    candidate_observer_message_type: candidateObservation.message_type,
    candidate_observer_permission: candidateObservation.required_permission,
    candidate_observer_endpoint: candidateObservation.endpoint,
    adapter_contract_accepted: contractAccepted,
    real_intake_accepted: realAccepted,
    evidence_counts: {
      provider_records: realIntake.provider_record_count,
      meeting_app_records: realIntake.meeting_app_record_count,
      dom_records: domDiagnosis.record_count,
      fixture_evidence: realIntake.fixture_evidence_count,
    },
    missing: {
      provider_env: providerMissingEnv,
      dom_issue_codes: blockingCodes(domDiagnosis.issues),
      real_intake_blocking_codes: blockingCodes(realIntake.blocking_checks),
      live_readiness_blocking_codes: blockingCodes(liveReadiness.blocking_checks),
      contract_issue_codes: blockingCodes(contractReport.issues),
      candidate_observation_issue_codes: candidateObservation.issues ?? [],
    },
    commands: handoffCommands(platform, fieldPlan),
    required_host_contract: {
      annotation_timestamp_field: 'captured_at_ms',
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      meeting_axis_source_order: contract.realtime_axis?.rules,
      per_meeting_annotation_isolation_required: true,
      candidate_observation_required_for_host_axis_binding: true,
      candidate_observation_message_type: candidateObservation.message_type,
      candidate_observation_runtime_action: candidateObservation.runtime_event_action,
      candidate_observation_endpoint: candidateObservation.endpoint,
      candidate_observation_runtime_event_endpoint: candidateObservation.runtime_event_endpoint,
      candidate_observation_required_permission: candidateObservation.required_permission,
    },
    sdk_methods: [
      'platformHandoffReadiness',
      'platformHandoffReadinessMatrix',
      'platformLiveAdapterHandoff',
      'platformFieldIntakePlan',
      'meetingAppDomAdaptationDiagnosis',
      'platformRealEvidenceIntake',
    ],
    reports: {
      contract_acceptance: contractReport,
      provider_connection: provider,
      candidate_observation: candidateObservation,
      dom_diagnosis: domDiagnosis,
      field_intake: fieldPlan,
      real_intake: realIntake,
      live_readiness: liveReadiness,
      live_handoff: liveHandoff,
    },
    next_actions: nextActions({
      status,
      contractReport,
      provider,
      domDiagnosis,
      fieldPlan,
      realIntake,
      liveReadiness,
      candidateObservation,
    }),
  });
}

export function buildMeetingPlatformHandoffReadinessMatrix(input = {}, options = {}) {
  const platforms = selectedPlatforms({ ...input, ...options });
  const reports = platforms.map((platform) => buildMeetingPlatformHandoffReadiness(
    platform,
    platformInput(input, platform),
    {
      ...options,
      platforms: undefined,
      platform_keys: undefined,
    },
  ));
  return {
    type: 'meeting_platform_handoff_readiness_matrix',
    schema: MEETING_PLATFORM_HANDOFF_READINESS_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_HANDOFF_READINESS_SCHEMA_VERSION,
    platform_count: reports.length,
    handoff_ready_count: reports.filter((report) => report.handoff_ready).length,
    pilot_ready_count: reports.filter((report) => report.pilot_ready).length,
    production_ready_count: reports.filter((report) => report.production_ready).length,
    local_observer_ready_count: reports.filter((report) => report.local_observer_ready).length,
    candidate_observer_count: reports.filter((report) => report.candidate_observation_ready).length,
    provider_reconcile_ready_count: reports.filter((report) => report.provider_reconcile_ready).length,
    provider_setup_needed_count: reports.filter((report) => (report.missing?.provider_env?.length ?? 0) > 0).length,
    local_evidence_needed_count: reports.filter((report) => report.local_observer_ready !== true).length,
    platforms,
    rows: reports.map((report) => ({
      platform: report.platform,
      display_name: report.display_name,
      status: report.status,
      handoff_ready: report.handoff_ready,
      pilot_ready: report.pilot_ready,
      production_ready: report.production_ready,
      local_observer_ready: report.local_observer_ready,
      candidate_observation_ready: report.candidate_observation_ready,
      candidate_observer_message_type: report.candidate_observer_message_type,
      candidate_observer_permission: report.candidate_observer_permission,
      candidate_observer_endpoint: report.candidate_observer_endpoint,
      provider_reconcile_ready: report.provider_reconcile_ready,
      provider_missing_env: report.missing?.provider_env ?? [],
      provider_record_count: report.evidence_counts?.provider_records ?? 0,
      meeting_app_record_count: report.evidence_counts?.meeting_app_records ?? 0,
      dom_record_count: report.evidence_counts?.dom_records ?? 0,
      first_next_action: report.next_actions?.[0],
      next_actions: report.next_actions,
    })),
    reports,
    next_actions: unique(reports.flatMap((report) => report.next_actions ?? [])),
  };
}

export function assertMeetingPlatformHandoffReadiness(platformOrInput = {}, input = {}, options = {}) {
  const report = buildMeetingPlatformHandoffReadiness(platformOrInput, input, options);
  if (report.handoff_ready !== true) {
    throw new MeetingTimelineSdkError(`Meeting platform handoff readiness failed for ${report.platform}`, {
      platform: report.platform,
      status: report.status,
      next_actions: report.next_actions,
      report,
    });
  }
  return report;
}

export function assertMeetingPlatformHandoffReadinessMatrix(input = {}, options = {}) {
  const matrix = buildMeetingPlatformHandoffReadinessMatrix(input, options);
  if (matrix.handoff_ready_count !== matrix.platform_count) {
    throw new MeetingTimelineSdkError('Meeting platform handoff readiness matrix failed', {
      platform_count: matrix.platform_count,
      handoff_ready_count: matrix.handoff_ready_count,
      failed_platforms: matrix.rows.filter((row) => row.handoff_ready !== true).map((row) => row.platform),
      next_actions: matrix.next_actions,
      matrix,
    });
  }
  return matrix;
}
