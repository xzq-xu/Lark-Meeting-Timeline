import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import { buildMeetingAppDomAdaptationDiagnosis } from './meeting-app-profile.mjs';
import { meetingAppSnapshotRecords } from './meeting-app-snapshot-recorder.mjs';
import {
  runMeetingPlatformRuntimeHostReplay,
} from './meeting-platform-runtime-host-verifier.mjs';
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

function replayInputFrom(input = {}, evidence = {}, options = {}) {
  return firstNonEmpty(
    options.runtimeHostReplayInput,
    options.runtime_host_replay_input,
    input.runtimeHostReplayInput,
    input.runtime_host_replay_input,
    input.replayInput,
    input.replay_input,
    evidence.evidencePackage,
    compactObject({
      records: evidence.records,
      snapshots: evidence.snapshots,
      meetingAppRecords: evidence.meetingAppRecords,
      meeting_app_records: evidence.meetingAppRecords,
      meetingAppRecordSet: evidence.recordSet,
      meeting_app_record_set: evidence.recordSet,
    }),
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

function runtimeReplayRequired(options = {}, runtimeHostReplay = undefined) {
  return firstNonEmpty(options.requireRuntimeHostReplay, options.require_runtime_host_replay, false) === true
    || Boolean(runtimeHostReplay);
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
  runtimeHostReplayRequired,
  runtimeHostReplayAccepted,
} = {}) {
  if (productionReady) return 'production_ready';
  if (runtimeHostReplayRequired && !runtimeHostReplayAccepted && realAccepted) return 'needs_runtime_host_replay';
  if (pilotReady && providerReady) return 'pilot_ready_provider_reconcile_pending';
  if (pilotReady) return 'pilot_ready_provider_setup_pending';
  if (!candidateObservationReady) return 'needs_candidate_observation_contract';
  if (!contractAccepted) return 'adapter_contract_blocked';
  if (!domAccepted) return 'needs_local_observer_evidence';
  if (providerMissingEnv.length > 0) return 'needs_provider_credentials';
  if (!realAccepted) return 'needs_provider_or_package_evidence';
  if (runtimeHostReplayRequired && !runtimeHostReplayAccepted) return 'needs_runtime_host_replay';
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
  runtimeHostReplay,
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
    needs_runtime_host_replay: 'fix_runtime_host_replay_before_sdk_handoff',
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
    ...(runtimeHostReplay?.next_actions ?? []),
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
  const runtimeHostReplay = replayReportFrom(rawInput, mergedOptions);
  const requireRuntimeReplay = runtimeReplayRequired(mergedOptions, runtimeHostReplay);
  const runtimeReplayAccepted = runtimeHostReplay?.accepted === true;
  const runtimeReplayGatePassed = !requireRuntimeReplay || runtimeReplayAccepted;
  const candidateObservation = liveReadiness.candidate_observation ?? liveHandoff.candidate_observation_contract ?? {};
  const candidateObservationReady = candidateObservation.ready === true;
  const pilotReady = contractAccepted && candidateObservationReady && domAccepted;
  const productionReady = pilotReady && realAccepted && liveReadiness.production_ready === true && runtimeReplayGatePassed;
  const handoffReady = contractAccepted && candidateObservationReady && (pilotReady || liveReadiness.passed === true) && runtimeReplayGatePassed;
  const status = statusFor({
    productionReady,
    pilotReady,
    contractAccepted,
    providerReady,
    providerMissingEnv,
    candidateObservationReady,
    domAccepted,
    realAccepted,
    runtimeHostReplayRequired: requireRuntimeReplay,
    runtimeHostReplayAccepted: runtimeReplayAccepted,
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
    runtime_host_replay_required: requireRuntimeReplay,
    runtime_host_replay_accepted: runtimeHostReplay ? runtimeReplayAccepted : undefined,
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
      runtime_host_replay_missing: runtimeHostReplay?.missing ?? [],
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
      runtime_host_replay_required_for_external_sdk_handoff: true,
    },
    sdk_methods: [
      'platformHandoffReadiness',
      'platformHandoffReadinessMatrix',
      'runPlatformHandoffReadiness',
      'runPlatformHandoffReadinessMatrix',
      'platformLiveAdapterHandoff',
      'platformFieldIntakePlan',
      'meetingAppDomAdaptationDiagnosis',
      'platformRealEvidenceIntake',
      'replayPlatformRuntimeHost',
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
      runtime_host_replay: runtimeHostReplay,
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
      runtimeHostReplay,
    }),
  });
}

function matrixFromReports(reports = [], platforms = reports.map((report) => report.platform)) {
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
    runtime_host_replay_ready_count: reports.filter((report) => report.runtime_host_replay_accepted === true).length,
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
      runtime_host_replay_required: report.runtime_host_replay_required,
      runtime_host_replay_accepted: report.runtime_host_replay_accepted,
      runtime_host_replay_missing: report.missing?.runtime_host_replay_missing ?? [],
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
  return matrixFromReports(reports, platforms);
}

export async function runMeetingPlatformHandoffReadiness(platformOrInput = {}, input = {}, options = {}) {
  const objectInput = platformOrInput && typeof platformOrInput === 'object' && !Array.isArray(platformOrInput);
  const rawPlatform = objectInput
    ? firstNonEmpty(platformOrInput.platform, platformOrInput.provider, platformOrInput.key, platformOrInput.name)
    : platformOrInput;
  const platform = normalizeMeetingPlatform(rawPlatform);
  const rawInput = objectInput ? platformOrInput : input;
  const mergedOptions = objectInput ? { ...input, ...options } : { ...input, ...options };
  const evidence = evidenceFor(platform, rawInput, mergedOptions);
  const explicitReplay = replayReportFrom(rawInput, mergedOptions);
  let runtimeHostReplay = explicitReplay;
  if (!runtimeHostReplay) {
    try {
      runtimeHostReplay = await runMeetingPlatformRuntimeHostReplay(
        platform,
        replayInputFrom(rawInput, evidence, mergedOptions),
        mergedOptions.runtimeHostReplayOptions ?? mergedOptions.runtime_host_replay_options ?? mergedOptions,
      );
    } catch (error) {
      runtimeHostReplay = replayErrorReport(platform, error);
    }
  }
  return buildMeetingPlatformHandoffReadiness(platform, rawInput, {
    ...mergedOptions,
    runtimeHostReplay,
    runtime_host_replay: runtimeHostReplay,
    requireRuntimeHostReplay: true,
    require_runtime_host_replay: true,
  });
}

export async function runMeetingPlatformHandoffReadinessMatrix(input = {}, options = {}) {
  const platforms = selectedPlatforms({ ...input, ...options });
  const reports = [];
  for (const platform of platforms) {
    reports.push(await runMeetingPlatformHandoffReadiness(
      platform,
      platformInput(input, platform),
      {
        ...options,
        platforms: undefined,
        platform_keys: undefined,
      },
    ));
  }
  return matrixFromReports(reports, platforms);
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
