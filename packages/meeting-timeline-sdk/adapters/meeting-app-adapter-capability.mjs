import { compactObject } from '../index.mjs';
import {
  buildMeetingAppAdapterFitMatrix,
  buildMeetingAppAdapterFitReport,
} from './meeting-apps.mjs';
import {
  buildMeetingAppAdapterHandoffPackage,
  buildMeetingAppAdapterVerificationReport,
} from './meeting-app-adapter-handoff-package.mjs';
import {
  MEETING_PLATFORM_KEYS,
  normalizeMeetingPlatform,
  platformCapabilityContract,
} from './platform-setup.mjs';

export const MEETING_APP_ADAPTER_CAPABILITY_REPORT_SCHEMA = 'meeting_app_adapter_capability_report';
export const MEETING_APP_ADAPTER_CAPABILITY_MATRIX_SCHEMA = 'meeting_app_adapter_capability_matrix';
export const MEETING_APP_ADAPTER_EXECUTION_PLAN_SCHEMA = 'meeting_app_adapter_execution_plan';
export const MEETING_APP_ADAPTER_EXECUTION_PLAN_MATRIX_SCHEMA = 'meeting_app_adapter_execution_plan_matrix';
export const MEETING_APP_ADAPTER_CAPABILITY_SCHEMA_VERSION = 1;

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value == null) return [];
  return [value];
}

function unique(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function normalizeOptionalPlatform(platform) {
  if (platform == null || platform === '') return undefined;
  return normalizeMeetingPlatform(platform);
}

function keyAliases(platform) {
  const dashed = platform.replaceAll('_', '-');
  const aliases = platform === 'microsoft_teams'
    ? ['teams', 'microsoft-teams']
    : platform === 'google_meet'
      ? ['meet', 'google-meet']
      : platform === 'lark'
        ? ['feishu', 'larksuite']
        : [];
  return unique([platform, dashed, ...aliases]);
}

function inputForPlatform(platform, options = {}) {
  const source = firstNonEmpty(
    options.inputs,
    options.inputByPlatform,
    options.input_by_platform,
    options.snapshots,
    options.snapshotByPlatform,
    options.snapshot_by_platform,
  );
  if (source && typeof source === 'object' && !Array.isArray(source)) {
    for (const key of keyAliases(platform)) {
      if (source[key] != null) return source[key];
    }
  }
  return firstNonEmpty(options.input, options.snapshot, options.sample);
}

function evidenceForPlatform(platform, options = {}) {
  const source = firstNonEmpty(
    options.evidenceByAdapter,
    options.evidence_by_adapter,
    options.evidenceByPlatform,
    options.evidence_by_platform,
  );
  if (source && typeof source === 'object' && !Array.isArray(source)) {
    for (const key of keyAliases(platform)) {
      if (source[key] != null) return source[key];
    }
  }
  return firstNonEmpty(options.evidence, options.evidence_package, options.evidencePackage);
}

function hasMeaningfulInput(value) {
  if (value == null) return false;
  if (Array.isArray(value)) return value.length > 0;
  if (typeof value !== 'object') return true;
  return Object.keys(value).length > 0;
}

function featureStatus(contract = {}, feature) {
  return String(contract[feature]?.status ?? 'not_declared');
}

function supported(status) {
  const text = String(status || '');
  return text === 'supported'
    || text.startsWith('supported_')
    || text === 'metadata_supported'
    || text === 'supported_when_cloud_recording_transcript_enabled';
}

function featureSource(contract = {}, feature) {
  return contract[feature]?.source ?? contract[feature]?.detail;
}

function rowFeature(name, values = {}) {
  return compactObject({
    name,
    status: values.status,
    provider_status: values.providerStatus,
    provider_declared: values.providerDeclared,
    local_ready: values.localReady,
    post_meeting_backfill: values.postMeetingBackfill,
    source: values.source,
    fallback: values.fallback,
    signal_types: values.signalTypes,
  });
}

function recommendedMode({ realtimeAxis, localAxisReady, speakerLocalReady, participantReady }) {
  if (localAxisReady && speakerLocalReady && participantReady) return 'hybrid_local_observer_first';
  if (localAxisReady) return 'local_observer_axis_with_provider_backfill';
  if (realtimeAxis.providerDeclared) return 'provider_event_primary_with_local_snapshot_required';
  return 'post_meeting_artifact_only';
}

function riskLevel({ mode, verification, fitReport }) {
  if (verification?.production_ready === true) return 'low';
  if (mode === 'hybrid_local_observer_first' && fitReport?.accepted === true) return 'medium';
  if (mode === 'provider_event_primary_with_local_snapshot_required') return 'medium_high';
  return 'high';
}

function nextActions({ fitInputProvided, fitReport, verification, mode }) {
  const actions = [];
  if (!fitInputProvided) actions.push('capture_live_meeting_app_snapshot');
  if (fitReport && !fitReport.ready_for_speaker_track) actions.push('capture_active_speaker_tile_or_audio_activity');
  if (fitReport && !fitReport.ready_for_participant_track) actions.push('capture_participant_tiles_or_roster_rows');
  if (!verification || verification.live_evidence_ready !== true) actions.push('attach_live_evidence_for_production_readiness');
  if (mode.includes('provider_event')) actions.push('configure_provider_event_subscription_and_reconcile_with_local_axis');
  actions.push('keep_transcript_import_post_meeting_and_nonblocking');
  return unique(actions);
}

function capabilityInput(input = {}, options = {}) {
  return input?.schema === MEETING_APP_ADAPTER_CAPABILITY_REPORT_SCHEMA
    ? input
    : buildMeetingAppAdapterCapabilityReport(input, options);
}

function planStep(id, values = {}) {
  return compactObject({
    id,
    role: values.role,
    source: values.source,
    status: values.status,
    required_for_realtime: values.requiredForRealtime,
    required_for_pilot: values.requiredForPilot,
    required_for_production: values.requiredForProduction,
    blocks_realtime_if_missing: values.blocksRealtimeIfMissing,
    sdk_modules: values.sdkModules,
    evidence: values.evidence,
    fallback: values.fallback,
  });
}

function statusFromReady(ready, fallbackStatus = 'pending') {
  return ready ? 'ready' : fallbackStatus;
}

function executionSteps(capability = {}) {
  const caps = capability.timeline_capabilities ?? {};
  const realtimeAxis = caps.realtime_axis ?? {};
  const speakerTrack = caps.speaker_track ?? {};
  const participantTrack = caps.participant_track ?? {};
  const transcript = caps.post_meeting_transcript ?? {};
  const recording = caps.recording ?? {};
  const localAxisReady = realtimeAxis.local_ready === true;
  const providerAxisDeclared = realtimeAxis.provider_declared === true;
  const speakerLocalReady = speakerTrack.local_ready === true;
  const participantReady = participantTrack.local_ready === true || participantTrack.provider_declared === true;
  return [
    planStep('capture_meeting_app_snapshot', {
      role: 'observe_current_meeting_window_and_extract_candidate_axis_signals',
      source: 'browser_extension_or_native_detector',
      status: statusFromReady(localAxisReady, 'needs_live_snapshot'),
      requiredForRealtime: true,
      requiredForPilot: true,
      blocksRealtimeIfMissing: providerAxisDeclared ? false : true,
      sdkModules: [
        '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-capture',
        '@ai-annotation/meeting-timeline-sdk/adapters/meeting-apps',
      ],
      evidence: ['live_dom_snapshot', 'candidate_observation'],
    }),
    planStep('start_or_reconcile_meeting_axis', {
      role: localAxisReady ? 'start_low_latency_axis_from_local_observer' : 'reconcile_axis_from_provider_event_until_local_snapshot_exists',
      source: localAxisReady ? 'local_observer' : 'provider_event',
      status: localAxisReady || providerAxisDeclared ? 'ready' : 'blocked',
      requiredForRealtime: true,
      requiredForPilot: true,
      blocksRealtimeIfMissing: true,
      sdkModules: [
        '@ai-annotation/meeting-timeline-sdk/adapters/timeline-bridge',
        '@ai-annotation/meeting-timeline-sdk/adapters/platform-ingest',
      ],
      evidence: localAxisReady ? ['candidate_observation'] : ['provider_meeting_started_event'],
      fallback: realtimeAxis.fallback,
    }),
    planStep('insert_annotation_on_current_axis', {
      role: 'place_human_marks_by_captured_at_ms_on_the_selected_meeting_axis',
      source: 'device_or_host_annotation_stream',
      status: capability.pilot_ready ? 'ready' : 'blocked',
      requiredForRealtime: true,
      requiredForPilot: true,
      requiredForProduction: true,
      blocksRealtimeIfMissing: true,
      sdkModules: ['@ai-annotation/meeting-timeline-sdk'],
      evidence: ['annotation_insert_current_axis'],
    }),
    planStep('emit_speaker_position_markers', {
      role: 'draw_speaker_positions_without_realtime_transcript_content',
      source: speakerLocalReady ? 'local_observer_active_speaker' : 'post_meeting_transcript_backfill',
      status: speakerLocalReady ? 'ready' : transcript.provider_declared ? 'backfill_only' : 'missing',
      requiredForRealtime: false,
      requiredForPilot: false,
      blocksRealtimeIfMissing: false,
      sdkModules: [
        '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-track-runtime',
        '@ai-annotation/meeting-timeline-sdk/adapters/platform-speaker-track',
      ],
      evidence: ['speaker_track'],
      fallback: speakerTrack.fallback,
    }),
    planStep('track_visible_participants', {
      role: 'draw_participant_presence_positions_and_improve_speaker_attribution',
      source: participantTrack.local_ready ? 'local_observer_roster' : 'provider_participant_events',
      status: participantReady ? 'ready' : 'missing',
      requiredForRealtime: false,
      requiredForPilot: false,
      blocksRealtimeIfMissing: false,
      sdkModules: ['@ai-annotation/meeting-timeline-sdk/adapters/platform-participant-track'],
      evidence: ['participant_track'],
    }),
    planStep('configure_provider_reconcile', {
      role: 'authoritative_start_end_participant_artifact_reconcile_and_production_evidence',
      source: 'official_provider_webhook_or_long_connection',
      status: providerAxisDeclared ? 'ready' : 'not_supported',
      requiredForRealtime: false,
      requiredForProduction: true,
      blocksRealtimeIfMissing: false,
      sdkModules: [
        '@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-router',
        '@ai-annotation/meeting-timeline-sdk/adapters/platform-provider-connection',
      ],
      evidence: ['provider_meeting_started_event', 'provider_meeting_ended_event'],
    }),
    planStep('import_post_meeting_artifacts', {
      role: 'backfill_transcript_recording_and_speaker_text_after_meeting',
      source: 'provider_artifact_api',
      status: transcript.provider_declared || recording.provider_declared ? 'ready' : 'not_supported',
      requiredForRealtime: false,
      requiredForProduction: false,
      blocksRealtimeIfMissing: false,
      sdkModules: ['@ai-annotation/meeting-timeline-sdk/adapters/transcript'],
      evidence: ['post_meeting_transcript', 'recording'],
    }),
  ];
}

function executionGates(capability = {}) {
  const liveSnapshotAccepted = capability.evidence_state?.live_snapshot_accepted === true;
  const liveEvidenceReady = capability.evidence_state?.live_evidence_ready === true;
  return [
    {
      id: 'static_adapter_contract',
      required_for: ['pilot', 'production'],
      status: capability.static_ready ? 'passed' : 'failed',
    },
    {
      id: 'axis_signal_available',
      required_for: ['pilot', 'production'],
      status: capability.pilot_ready ? 'passed' : 'failed',
    },
    {
      id: 'live_snapshot_accepted',
      required_for: ['low_latency_pilot', 'production'],
      status: liveSnapshotAccepted ? 'passed' : 'missing',
    },
    {
      id: 'live_evidence_package',
      required_for: ['production'],
      status: liveEvidenceReady ? 'passed' : 'missing',
      missing_evidence_count: capability.evidence_state?.missing_evidence_count,
    },
  ];
}

function firstBlockedStep(steps = []) {
  return steps.find((step) => step.blocks_realtime_if_missing === true && step.status !== 'ready')?.id;
}

function verificationForPlatform(platform, options = {}) {
  const evidence = evidenceForPlatform(platform, options);
  const include = evidence != null
    || options.includeVerification === true
    || options.include_verification === true
    || options.verify === true;
  if (!include) return undefined;
  return buildMeetingAppAdapterVerificationReport(platform, {
    ...options,
    evidence,
  });
}

export function buildMeetingAppAdapterCapabilityReport(platformOrOptions = {}, options = {}) {
  const rawOptions = typeof platformOrOptions === 'string'
    ? { ...options, platform: platformOrOptions }
    : { ...platformOrOptions, ...options };
  const platform = normalizeOptionalPlatform(firstNonEmpty(rawOptions.platform, rawOptions.provider, rawOptions.adapter_key, rawOptions.adapterKey));
  if (!platform) {
    throw new Error('platform is required for buildMeetingAppAdapterCapabilityReport');
  }
  const contract = platformCapabilityContract(platform, rawOptions);
  const handoffPackage = buildMeetingAppAdapterHandoffPackage(platform, rawOptions);
  const fitInput = inputForPlatform(platform, rawOptions);
  const fitInputProvided = hasMeaningfulInput(fitInput);
  const fitReport = fitInputProvided
    ? buildMeetingAppAdapterFitReport(fitInput, { ...rawOptions, platform })
    : undefined;
  const verification = verificationForPlatform(platform, rawOptions);

  const providerRealtimeStatus = featureStatus(contract, 'realtime_axis');
  const providerParticipantStatus = featureStatus(contract, 'participant_track');
  const providerSpeakerStatus = featureStatus(contract, 'speaker_activity');
  const transcriptStatus = featureStatus(contract, 'post_meeting_transcript');
  const recordingStatus = featureStatus(contract, 'recording');
  const subscriptionStatus = featureStatus(contract, 'subscription_lifecycle');

  const localAxisReady = fitReport?.ready_for_realtime_axis === true;
  const speakerLocalReady = fitReport?.ready_for_speaker_track === true;
  const participantLocalReady = fitReport?.ready_for_participant_track === true;
  const realtimeAxis = {
    providerStatus: providerRealtimeStatus,
    providerDeclared: supported(providerRealtimeStatus),
    localReady: localAxisReady,
  };
  const participantReady = supported(providerParticipantStatus) || participantLocalReady;
  const speakerBackfill = supported(transcriptStatus);
  const mode = recommendedMode({
    realtimeAxis,
    localAxisReady,
    speakerLocalReady,
    participantReady,
  });
  const staticReady = handoffPackage.accepted === true;
  const pilotReady = staticReady && (localAxisReady || realtimeAxis.providerDeclared);
  const productionReady = verification?.production_ready === true;
  const actions = nextActions({
    fitInputProvided,
    fitReport,
    verification,
    mode,
  });

  return {
    type: 'meeting_app_adapter_capability_report',
    schema: MEETING_APP_ADAPTER_CAPABILITY_REPORT_SCHEMA,
    schema_version: MEETING_APP_ADAPTER_CAPABILITY_SCHEMA_VERSION,
    platform,
    display_name: contract.display_name,
    accepted: pilotReady,
    static_ready: staticReady,
    pilot_ready: pilotReady,
    production_ready: productionReady,
    recommended_mode: mode,
    risk_level: riskLevel({ mode, verification, fitReport }),
    timeline_capabilities: {
      realtime_axis: rowFeature('realtime_axis', {
        status: localAxisReady ? 'local_ready' : realtimeAxis.providerDeclared ? 'provider_declared' : 'missing',
        providerStatus: providerRealtimeStatus,
        providerDeclared: realtimeAxis.providerDeclared,
        localReady: localAxisReady,
        source: featureSource(contract, 'realtime_axis'),
        fallback: contract.realtime_axis?.fallback,
        signalTypes: contract.realtime_axis?.signal_types,
      }),
      speaker_track: rowFeature('speaker_track', {
        status: speakerLocalReady ? 'local_ready' : speakerBackfill ? 'post_meeting_backfill' : 'missing',
        providerStatus: providerSpeakerStatus,
        providerDeclared: supported(providerSpeakerStatus),
        localReady: speakerLocalReady,
        postMeetingBackfill: speakerBackfill,
        source: featureSource(contract, 'speaker_activity'),
        fallback: contract.speaker_activity?.fallback,
        signalTypes: contract.speaker_activity?.signal_types,
      }),
      participant_track: rowFeature('participant_track', {
        status: participantReady ? 'available' : 'missing',
        providerStatus: providerParticipantStatus,
        providerDeclared: supported(providerParticipantStatus),
        localReady: participantLocalReady,
        source: featureSource(contract, 'participant_track'),
        signalTypes: contract.participant_track?.signal_types,
      }),
      annotation_timeline: rowFeature('annotation_timeline', {
        status: pilotReady ? 'ready_to_insert_current_axis' : 'needs_axis_signal',
        localReady: localAxisReady,
        providerDeclared: realtimeAxis.providerDeclared,
        source: 'captured_at_ms annotation intake with local/provider axis reconciliation',
        signalTypes: ['annotation_insert_current_axis'],
      }),
      post_meeting_transcript: rowFeature('post_meeting_transcript', {
        status: transcriptStatus,
        providerStatus: transcriptStatus,
        providerDeclared: supported(transcriptStatus),
        source: featureSource(contract, 'post_meeting_transcript'),
      }),
      recording: rowFeature('recording', {
        status: recordingStatus,
        providerStatus: recordingStatus,
        providerDeclared: supported(recordingStatus),
        source: featureSource(contract, 'recording'),
      }),
      subscription_lifecycle: rowFeature('subscription_lifecycle', {
        status: subscriptionStatus,
        providerStatus: subscriptionStatus,
        providerDeclared: supported(subscriptionStatus),
        source: featureSource(contract, 'subscription_lifecycle'),
      }),
    },
    evidence_state: {
      live_snapshot_supplied: fitInputProvided,
      live_snapshot_accepted: fitReport?.accepted === true,
      live_evidence_ready: verification?.live_evidence_ready === true,
      verification_target: verification?.target,
      missing_evidence_count: verification?.missing_evidence_count,
    },
    fit_report: fitReport,
    verification_report: verification,
    handoff_package: rawOptions.includeHandoffPackage === true || rawOptions.include_handoff_package === true
      ? handoffPackage
      : undefined,
    next_actions: actions,
  };
}

export function buildMeetingAppAdapterCapabilityMatrix(options = {}) {
  const platforms = unique(asArray(firstNonEmpty(
    options.platforms,
    options.platform_keys,
    options.platformKeys,
    MEETING_PLATFORM_KEYS.filter((platform) => platform !== 'local_detector'),
  )).map((platform) => normalizeMeetingPlatform(platform)));
  const reports = platforms.map((platform) => buildMeetingAppAdapterCapabilityReport(platform, options));
  const fitMatrix = buildMeetingAppAdapterFitMatrix({
    platforms,
    inputs: firstNonEmpty(options.inputs, options.inputByPlatform, options.input_by_platform, options.snapshots),
  });
  return {
    type: 'meeting_app_adapter_capability_matrix',
    schema: MEETING_APP_ADAPTER_CAPABILITY_MATRIX_SCHEMA,
    schema_version: MEETING_APP_ADAPTER_CAPABILITY_SCHEMA_VERSION,
    platform_count: reports.length,
    accepted_count: reports.filter((report) => report.accepted).length,
    static_ready_count: reports.filter((report) => report.static_ready).length,
    pilot_ready_count: reports.filter((report) => report.pilot_ready).length,
    production_ready_count: reports.filter((report) => report.production_ready).length,
    local_axis_ready_count: reports.filter((report) => report.timeline_capabilities.realtime_axis.local_ready).length,
    local_speaker_ready_count: reports.filter((report) => report.timeline_capabilities.speaker_track.local_ready).length,
    provider_axis_declared_count: reports.filter((report) => report.timeline_capabilities.realtime_axis.provider_declared).length,
    platforms: reports.map((report) => report.platform),
    rows: reports.map((report) => ({
      platform: report.platform,
      display_name: report.display_name,
      accepted: report.accepted,
      static_ready: report.static_ready,
      pilot_ready: report.pilot_ready,
      production_ready: report.production_ready,
      recommended_mode: report.recommended_mode,
      risk_level: report.risk_level,
      realtime_axis_status: report.timeline_capabilities.realtime_axis.status,
      speaker_track_status: report.timeline_capabilities.speaker_track.status,
      participant_track_status: report.timeline_capabilities.participant_track.status,
      live_snapshot_accepted: report.evidence_state.live_snapshot_accepted,
      missing_evidence_count: report.evidence_state.missing_evidence_count,
      first_next_action: report.next_actions[0],
    })),
    fit_matrix: fitMatrix,
    reports,
    next_actions: unique(reports.flatMap((report) => report.next_actions)),
  };
}

export function buildMeetingAppAdapterExecutionPlan(capabilityOrPlatform = {}, options = {}) {
  const capability = capabilityInput(capabilityOrPlatform, options);
  const steps = executionSteps(capability);
  const gates = executionGates(capability);
  const blockedStep = firstBlockedStep(steps);
  return {
    type: 'meeting_app_adapter_execution_plan',
    schema: MEETING_APP_ADAPTER_EXECUTION_PLAN_SCHEMA,
    schema_version: MEETING_APP_ADAPTER_CAPABILITY_SCHEMA_VERSION,
    platform: capability.platform,
    display_name: capability.display_name,
    accepted: capability.pilot_ready === true,
    production_ready: capability.production_ready === true,
    recommended_mode: capability.recommended_mode,
    risk_level: capability.risk_level,
    first_blocked_step: blockedStep,
    realtime_ready: blockedStep == null,
    steps,
    gates,
    commands: {
      generate_capability_report: `npm run meeting-app:adapter-capability -- --platforms=${capability.platform}`,
      verify_adapter_evidence: `npm run meeting-app:adapter-verify -- --platforms=${capability.platform}`,
      verify_handoff_readiness: `npm run meeting-platform:handoff-readiness -- --platforms=${capability.platform}`,
    },
    handoff_requirements: [
      'capture_live_meeting_app_snapshot_with_captured_at_ms',
      'prove_annotation_insert_current_axis',
      'keep_provider_events_nonblocking_for_realtime_annotation',
      'import_transcript_after_meeting_only',
    ],
    next_actions: capability.next_actions,
    capability_report: capability,
  };
}

export function buildMeetingAppAdapterExecutionPlanMatrix(options = {}) {
  const capabilityMatrix = options.capabilityMatrix?.schema === MEETING_APP_ADAPTER_CAPABILITY_MATRIX_SCHEMA
    ? options.capabilityMatrix
    : buildMeetingAppAdapterCapabilityMatrix(options);
  const plans = capabilityMatrix.reports.map((report) => buildMeetingAppAdapterExecutionPlan(report, options));
  return {
    type: 'meeting_app_adapter_execution_plan_matrix',
    schema: MEETING_APP_ADAPTER_EXECUTION_PLAN_MATRIX_SCHEMA,
    schema_version: MEETING_APP_ADAPTER_CAPABILITY_SCHEMA_VERSION,
    platform_count: plans.length,
    accepted_count: plans.filter((plan) => plan.accepted).length,
    production_ready_count: plans.filter((plan) => plan.production_ready).length,
    realtime_ready_count: plans.filter((plan) => plan.realtime_ready).length,
    platforms: plans.map((plan) => plan.platform),
    rows: plans.map((plan) => ({
      platform: plan.platform,
      accepted: plan.accepted,
      production_ready: plan.production_ready,
      realtime_ready: plan.realtime_ready,
      recommended_mode: plan.recommended_mode,
      risk_level: plan.risk_level,
      first_blocked_step: plan.first_blocked_step,
      first_next_action: plan.next_actions[0],
    })),
    plans,
    next_actions: unique(plans.flatMap((plan) => plan.next_actions)),
  };
}
