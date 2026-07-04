import {
  MeetingTimelineSdkError,
  compactObject,
} from '../index.mjs';
import { buildPlatformAcceptanceReport, buildMeetingPlatformAcceptanceSummary } from './platform-acceptance.mjs';
import {
  buildPlatformFixtureAcceptanceInput,
  buildPlatformFixtureSamples,
} from './platform-fixtures.mjs';
import {
  MEETING_APP_FIXTURE_PLATFORMS,
  buildAllMeetingAppFixtureSnapshots,
  buildMeetingAppFixtureAcceptanceReport,
  buildMeetingAppFixtureSnapshot,
  diagnoseMeetingAppFixture,
  diagnoseMeetingAppFixtureLifecycle,
} from './meeting-app-fixtures.mjs';
import {
  buildMeetingAppTrackPipeline,
  createMeetingAppTrackPipeline,
} from './meeting-app-track-pipeline.mjs';
import {
  createMeetingAppTrackRuntime,
} from './meeting-app-track-runtime.mjs';
import { diagnosePlatformEvent } from './platform-ingest.mjs';
import { buildMeetingPlatformOnboardingReport, buildMeetingPlatformOnboardingSummary } from './platform-onboarding.mjs';
import {
  MEETING_PLATFORM_KEYS,
  allPlatformCapabilityContracts,
  buildPlatformIntegrationPlan,
  buildPlatformPermissionPlan,
  buildPlatformSetup,
  normalizeMeetingPlatform,
  platformCapabilityContract,
} from './platform-setup.mjs';
import {
  assertMeetingPlatformRegistryManifest,
  buildMeetingPlatformRegistryAcceptanceReport,
  buildMeetingPlatformRegistryEntry,
  buildMeetingPlatformRegistryManifest,
} from './platform-registry.mjs';
import { createMeetingTimelineBridge } from './timeline-bridge.mjs';
import {
  buildPlatformWebhookRouteTable,
  buildPlatformWebhookRouterSetup,
  buildPlatformWebhookRouterStatus,
  createMeetingPlatformWebhookRouter,
  platformWebhookRoutePath,
} from './platform-webhook-router.mjs';
import { createMeetingPlatformFetchHandler } from './platform-http.mjs';
import { createMeetingPlatformNodeHandler } from './platform-node.mjs';
import {
  buildMeetingPlatformCaptureAcceptanceSummary,
  buildPlatformCaptureAcceptanceReport,
  buildPlatformCaptureSamples,
  capturePlatformWebRequest,
  capturePlatformWebhookEvent,
} from './platform-capture.mjs';
import {
  assertAllPlatformLaunchGates,
  assertPlatformLaunchGate,
  buildMeetingPlatformLaunchGateSummary,
  buildPlatformLaunchGate,
} from './platform-gate.mjs';
import {
  buildAllMeetingPlatformAdaptationRunbooks,
  buildAllMeetingPlatformRolloutPlans,
  buildMeetingPlatformAdaptationRunbook,
  buildMeetingPlatformAdaptationRunbookSummary,
  buildMeetingPlatformRolloutPlan,
  buildMeetingPlatformRolloutSummary,
} from './platform-rollout.mjs';
import {
  buildMeetingPlatformEvidencePackage,
  buildMeetingPlatformEvidencePackageSummary,
  createMeetingPlatformEvidencePackageBuilder,
  verifyMeetingPlatformEvidencePackage,
} from './platform-evidence-package.mjs';
import {
  buildMeetingPlatformEvidenceCorrelation,
} from './platform-evidence-correlation.mjs';
import {
  createMeetingPlatformEvidenceSession,
} from './platform-evidence-session.mjs';
import {
  assertMeetingPlatformLiveAdapterReadiness,
  assertMeetingPlatformLiveAdapterReadinessMatrix,
  buildMeetingPlatformLiveAdapterHandoff,
  buildMeetingPlatformLiveAdapterHandoffBundle,
  buildMeetingPlatformLiveAdapterMatrix,
  buildMeetingPlatformLiveAdapterPlan,
  buildMeetingPlatformLiveAdapterReadiness,
  buildMeetingPlatformLiveAdapterReadinessMatrix,
  createMeetingPlatformLiveAdapter,
  createMeetingPlatformLiveAdapterSuite,
} from './platform-live-adapter.mjs';
import {
  buildAllMeetingPlatformAdaptationStrategies,
  buildMeetingPlatformAdaptationStrategy,
  buildMeetingPlatformAdaptationStrategyMatrix,
} from './platform-strategy.mjs';
import {
  assertMeetingPlatformHostIntegrationScaffold,
  buildMeetingPlatformHostIntegrationPlan,
  buildMeetingPlatformHostIntegrationScaffold,
  buildMeetingPlatformHostIntegrationScaffoldAcceptanceReport,
} from './platform-host-integration.mjs';
import {
  buildMeetingPlatformProviderConnectionMatrix,
  buildMeetingPlatformProviderConnectionPack,
} from './platform-provider-connection.mjs';
import {
  assertMeetingPlatformSubscriptionHandoff,
  assertMeetingPlatformSubscriptionHandoffMatrix,
  buildMeetingPlatformSubscriptionHandoff,
  buildMeetingPlatformSubscriptionHandoffMatrix,
} from './platform-subscription-handoff.mjs';
import {
  buildMeetingPlatformSpeakerTrack,
  buildMeetingPlatformSpeakerTrackMatrix,
  buildMeetingPlatformSpeakerTrackPlan,
} from './platform-speaker-track.mjs';
import {
  buildMeetingPlatformParticipantTrack,
  buildMeetingPlatformParticipantTrackMatrix,
  buildMeetingPlatformParticipantTrackPlan,
} from './platform-participant-track.mjs';
import {
  buildMeetingPlatformTimelineView,
  buildMeetingPlatformTimelineViewMatrix,
  buildMeetingPlatformTimelineViewPlan,
  zoomMeetingPlatformTimelineViewport,
} from './platform-timeline-view.mjs';
import {
  buildMeetingPlatformAnnotationIntake,
  buildMeetingPlatformAnnotationIntakeMatrix,
  buildMeetingPlatformAnnotationIntakePlan,
} from './platform-annotation-intake.mjs';
import {
  buildMeetingPlatformClockSyncMatrix,
  buildMeetingPlatformClockSyncPlan,
  buildMeetingPlatformClockSyncReport,
} from './platform-clock-sync.mjs';
import {
  buildMeetingPlatformSessionBinding,
  buildMeetingPlatformSessionBindingMatrix,
  buildMeetingPlatformSessionBindingPlan,
} from './platform-session-binding.mjs';
import {
  buildMeetingPlatformRealtimeAnnotation,
  buildMeetingPlatformRealtimeAnnotationMatrix,
  buildMeetingPlatformRealtimeAnnotationPlan,
} from './platform-realtime-annotation.mjs';
import {
  buildMeetingPlatformArtifactHandoff,
  buildMeetingPlatformArtifactHandoffMatrix,
  buildMeetingPlatformArtifactHandoffPlan,
} from './platform-artifact-handoff.mjs';
import {
  buildMeetingPlatformRuntimeProfile,
  buildMeetingPlatformRuntimeProfileMatrix,
} from './platform-runtime-profile.mjs';
import {
  buildMeetingPlatformAdaptationPackage,
  buildMeetingPlatformAdaptationPackageMatrix,
} from './platform-adaptation-package.mjs';
import {
  buildMeetingPlatformRuntimeBundle,
  buildMeetingPlatformRuntimeBundleMatrix,
} from './platform-runtime-bundle.mjs';
import {
  buildMeetingPlatformRuntimeEventPlan,
  buildMeetingPlatformRuntimeEventPlanMatrix,
} from './platform-runtime-event.mjs';
import {
  assertMeetingPlatformAdapterContract,
  buildMeetingPlatformAdapterContractAcceptanceMatrix,
  buildMeetingPlatformAdapterContractAcceptanceReport,
  buildMeetingPlatformAdapterContract,
  buildMeetingPlatformAdapterContractMatrix,
} from './platform-adapter-contract.mjs';
import {
  assertMeetingPlatformAdapterSample,
  assertMeetingPlatformAdapterSampleMatrix,
  buildMeetingPlatformAdapterSamplePlan,
  runMeetingPlatformAdapterSample,
  runMeetingPlatformAdapterSampleMatrix,
} from './platform-adapter-sample.mjs';
import {
  assertMeetingPlatformRealEvidenceIntake,
  assertMeetingPlatformRealEvidenceIntakeMatrix,
  buildMeetingPlatformRealEvidenceIntakeMatrix,
  buildMeetingPlatformRealEvidenceIntakePlan,
  buildMeetingPlatformRealEvidenceIntakeReport,
} from './platform-real-intake.mjs';
import {
  buildMeetingPlatformFieldCaptureManifest,
  buildMeetingPlatformFieldCaptureManifestMatrix,
  buildMeetingPlatformFieldCollectorConfig,
  buildMeetingPlatformFieldCollectorConfigMatrix,
  buildMeetingPlatformFieldEvidenceBundle,
  buildMeetingPlatformFieldEvidenceMatrix,
  buildMeetingPlatformFieldCaptureMatrix,
  buildMeetingPlatformFieldCapturePlan,
} from './platform-field-capture.mjs';
import {
  buildMeetingPlatformFieldIntakeMatrix,
  buildMeetingPlatformFieldIntakePlan,
} from './platform-field-intake.mjs';
import {
  assertMeetingPlatformHandoffReadiness,
  assertMeetingPlatformHandoffReadinessMatrix,
  buildMeetingPlatformHandoffReadiness,
  buildMeetingPlatformHandoffReadinessMatrix,
} from './platform-handoff-readiness.mjs';
import {
  assertAllMeetingAppLaunchGates,
  assertMeetingAppLaunchGate,
  buildMeetingAppLaunchGate,
  buildMeetingAppLaunchGateSummary,
} from './meeting-app-gate.mjs';
import {
  buildMeetingAppLaunchGateInputFromRecords,
  createMeetingAppSnapshotRecorder,
} from './meeting-app-snapshot-recorder.mjs';
import {
  assertMeetingAppRuntimeAdapterConfig,
  assertMeetingAppRuntimeAdapterValidation,
  assertMeetingAppDeploymentManifest,
  buildAllMeetingAppDomAdaptationDiagnoses,
  buildAllMeetingAppDeploymentManifestAcceptanceReports,
  buildAllMeetingAppDeploymentManifests,
  buildAllMeetingAppIntegrationProfiles,
  buildAllMeetingAppLiveSnapshotCapturePlans,
  buildAllMeetingAppRuntimeAdapterAcceptanceReports,
  buildAllMeetingAppRuntimeAdapterConfigs,
  buildAllMeetingAppRuntimeAdapterValidationReports,
  buildMeetingAppDeploymentManifestAcceptanceReport,
  buildMeetingAppDeploymentManifestAcceptanceSummary,
  buildMeetingAppDomAdaptationDiagnosis,
  buildMeetingAppDomAdaptationDiagnosisMatrix,
  buildMeetingAppDeploymentManifest,
  buildMeetingAppIntegrationMatrix,
  buildMeetingAppIntegrationProfile,
  buildMeetingAppLiveEvidencePackage,
  buildMeetingAppLiveEvidencePackageSummary,
  buildMeetingAppLiveSnapshotCapturePlan,
  buildMeetingAppRuntimeAdapterAcceptanceReport,
  buildMeetingAppRuntimeAdapterConfig,
  buildMeetingAppRuntimeAdapterValidationReport,
} from './meeting-app-profile.mjs';
import {
  buildMeetingAppExtensionAttachedMessage,
  buildMeetingAppContentScriptManifest,
  buildMeetingAppExtensionClientCallMessage,
  buildMeetingAppExtensionInstallPlan,
  buildMeetingAppExtensionMatchPatterns,
  buildMeetingAppExtensionObserveCandidatesMessage,
  buildMeetingAppExtensionScaffold,
  buildMeetingAppExtensionScaffoldAcceptanceReport,
  buildMeetingAppExtensionStatusMessage,
  assertMeetingAppExtensionScaffold,
  meetingAppExtensionTimelineEndpoint,
  normalizeMeetingAppExtensionMessageType,
} from './meeting-app-extension.mjs';

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function isTimelineClient(value) {
  return Boolean(value)
    && typeof value === 'object'
    && typeof value.startMeeting === 'function'
    && typeof value.endMeeting === 'function'
    && typeof value.insertMark === 'function';
}

function kitDefaults(clientOrOptions, options = {}) {
  const baseOptions = isTimelineClient(clientOrOptions) ? {} : (clientOrOptions ?? {});
  const clientOptions = {
    ...(baseOptions.clientOptions ?? baseOptions.client_options ?? {}),
    ...(options.clientOptions ?? options.client_options ?? {}),
  };
  return compactObject({
    ...baseOptions,
    ...options,
    clientOptions,
    baseUrl: firstNonEmpty(options.baseUrl, options.base_url, baseOptions.baseUrl, baseOptions.base_url, clientOptions.baseUrl, clientOptions.base_url),
    basePath: firstNonEmpty(options.basePath, options.base_path, baseOptions.basePath, baseOptions.base_path, '/api/platform-events'),
    env: {
      ...(baseOptions.env ?? {}),
      ...(options.env ?? {}),
    },
  });
}

function withDefaults(defaults = {}, options = {}) {
  return compactObject({
    ...defaults,
    ...options,
    clientOptions: {
      ...(defaults.clientOptions ?? defaults.client_options ?? {}),
      ...(options.clientOptions ?? options.client_options ?? {}),
    },
    env: {
      ...(defaults.env ?? {}),
      ...(options.env ?? {}),
    },
  });
}

function bridgeInput(clientOrOptions, defaults = {}) {
  if (isTimelineClient(clientOrOptions)) return clientOrOptions;
  const clientOptions = {
    ...(clientOrOptions ?? {}),
    ...(defaults.clientOptions ?? defaults.client_options ?? {}),
    baseUrl: firstNonEmpty(defaults.baseUrl, defaults.base_url, clientOrOptions?.baseUrl, clientOrOptions?.base_url),
  };
  if (!clientOptions.baseUrl && !clientOptions.base_url) {
    throw new MeetingTimelineSdkError('baseUrl or MeetingTimelineClient is required for createMeetingPlatformTimelineKit');
  }
  return clientOptions;
}

function acceptanceOptions(defaults = {}, options = {}) {
  const merged = withDefaults(defaults, options);
  return {
    ...merged,
    requireEndEvent: firstNonEmpty(options.requireEndEvent, options.require_end_event, true),
  };
}

function platformOverview(platform, defaults = {}, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const merged = withDefaults(defaults, options);
  return {
    platform: key,
    capabilities: platformCapabilityContract(key, merged),
    setup: buildPlatformSetup(key, merged),
    permission_plan: buildPlatformPermissionPlan(key, merged),
    integration_plan: buildPlatformIntegrationPlan(key, merged),
    webhook_route: platformWebhookRoutePath(key, merged),
    fixture_samples: buildPlatformFixtureSamples(key, merged),
    meeting_app_fixture: MEETING_APP_FIXTURE_PLATFORMS.includes(key)
      ? diagnoseMeetingAppFixture(key, merged)
      : undefined,
  };
}

export function buildMeetingPlatformKitReport(options = {}) {
  const fixtureInput = buildPlatformFixtureAcceptanceInput(options);
  const meetingAppFixtureAcceptance = buildMeetingAppFixtureAcceptanceReport(options);
  const meetingAppLaunchGateSummary = buildMeetingAppLaunchGateSummary(options);
  return {
    base_url: options.baseUrl ?? options.base_url,
    base_path: options.basePath ?? options.base_path ?? '/api/platform-events',
    supported_platforms: MEETING_PLATFORM_KEYS,
    supported_meeting_app_platforms: MEETING_APP_FIXTURE_PLATFORMS,
    platform_registry_manifest: buildMeetingPlatformRegistryManifest(options),
    capabilities: allPlatformCapabilityContracts(options),
    webhook_router: buildPlatformWebhookRouterStatus(options),
    onboarding: buildMeetingPlatformOnboardingSummary(options),
    fixture_acceptance: buildMeetingPlatformAcceptanceSummary({
      ...fixtureInput,
      requireEndEvent: true,
      ...options,
      samples: fixtureInput.samples,
      env: fixtureInput.env,
    }),
    meeting_app_extension_install_plan: buildMeetingAppExtensionInstallPlan(options),
    meeting_app_extension_acceptance: buildMeetingAppExtensionScaffoldAcceptanceReport(options),
    meeting_app_integration_matrix: buildMeetingAppIntegrationMatrix(options),
    meeting_app_deployment_manifests: buildAllMeetingAppDeploymentManifests(options),
    meeting_app_deployment_manifest_acceptance: buildAllMeetingAppDeploymentManifestAcceptanceReports(options),
    meeting_app_deployment_manifest_acceptance_summary: buildMeetingAppDeploymentManifestAcceptanceSummary(options),
    meeting_app_dom_adaptation_diagnosis: buildAllMeetingAppDomAdaptationDiagnoses(options),
    meeting_app_dom_adaptation_diagnosis_matrix: buildMeetingAppDomAdaptationDiagnosisMatrix(options),
    meeting_app_runtime_adapter_configs: buildAllMeetingAppRuntimeAdapterConfigs(options),
    meeting_app_runtime_adapter_acceptance: buildAllMeetingAppRuntimeAdapterAcceptanceReports(options),
    meeting_app_live_snapshot_capture_plans: buildAllMeetingAppLiveSnapshotCapturePlans(options),
    meeting_app_runtime_adapter_validation: buildAllMeetingAppRuntimeAdapterValidationReports(options),
    meeting_app_fixture_acceptance: meetingAppFixtureAcceptance,
    meeting_app_launch_gate: meetingAppLaunchGateSummary,
    platform_rollout: buildMeetingPlatformRolloutSummary(options),
    platform_adaptation_runbook: buildMeetingPlatformAdaptationRunbookSummary(options),
    platform_adaptation_strategy: buildMeetingPlatformAdaptationStrategyMatrix(options),
    platform_live_adapter_matrix: buildMeetingPlatformLiveAdapterMatrix(options),
    platform_live_adapter_readiness_matrix: buildMeetingPlatformLiveAdapterReadinessMatrix(options),
    platform_live_adapter_handoff_bundle: buildMeetingPlatformLiveAdapterHandoffBundle(options),
    platform_host_integration: buildMeetingPlatformHostIntegrationPlan(options),
    platform_host_integration_acceptance: buildMeetingPlatformHostIntegrationScaffoldAcceptanceReport(options),
    platform_provider_connection_matrix: buildMeetingPlatformProviderConnectionMatrix(options),
    platform_subscription_handoff_matrix: buildMeetingPlatformSubscriptionHandoffMatrix(options),
    platform_speaker_track_matrix: buildMeetingPlatformSpeakerTrackMatrix(options),
    platform_participant_track_matrix: buildMeetingPlatformParticipantTrackMatrix(options),
    platform_timeline_view_matrix: buildMeetingPlatformTimelineViewMatrix(options),
    platform_annotation_intake_matrix: buildMeetingPlatformAnnotationIntakeMatrix(options),
    platform_clock_sync_matrix: buildMeetingPlatformClockSyncMatrix(options),
    platform_session_binding_matrix: buildMeetingPlatformSessionBindingMatrix(options),
    platform_realtime_annotation_matrix: buildMeetingPlatformRealtimeAnnotationMatrix(options),
    platform_artifact_handoff_matrix: buildMeetingPlatformArtifactHandoffMatrix(options),
    platform_runtime_profile_matrix: buildMeetingPlatformRuntimeProfileMatrix(options),
    platform_adaptation_package_matrix: buildMeetingPlatformAdaptationPackageMatrix(options),
    platform_runtime_event_plan_matrix: buildMeetingPlatformRuntimeEventPlanMatrix(options),
    platform_runtime_bundle_matrix: buildMeetingPlatformRuntimeBundleMatrix(options),
    platform_adapter_contract_matrix: buildMeetingPlatformAdapterContractMatrix(options),
    platform_adapter_contract_acceptance_matrix: buildMeetingPlatformAdapterContractAcceptanceMatrix(options),
    platform_field_capture_matrix: buildMeetingPlatformFieldCaptureMatrix(options),
    platform_field_capture_manifest_matrix: buildMeetingPlatformFieldCaptureManifestMatrix(options),
    platform_field_collector_config_matrix: buildMeetingPlatformFieldCollectorConfigMatrix(options),
    platform_field_intake_matrix: buildMeetingPlatformFieldIntakeMatrix(options),
    platform_handoff_readiness_matrix: buildMeetingPlatformHandoffReadinessMatrix(options),
  };
}

export function createMeetingPlatformTimelineKit(clientOrOptions, options = {}) {
  const defaults = kitDefaults(clientOrOptions, options);
  const bridge = createMeetingTimelineBridge(bridgeInput(clientOrOptions, defaults), defaults);
  const router = createMeetingPlatformWebhookRouter(bridge.client, {
    ...defaults,
    reconcile: firstNonEmpty(defaults.reconcile, defaults.reconciled, true),
  });
  const fetchHandler = createMeetingPlatformFetchHandler(router, defaults);
  const nodeHandler = createMeetingPlatformNodeHandler(router, defaults);

  return {
    client: bridge.client,
    bridge,
    webhookRouter: router,
    fetchHandler,
    nodeHandler,
    platforms: MEETING_PLATFORM_KEYS,
    observe(snapshot = {}, observeOptions = {}) {
      return bridge.observe(snapshot, observeOptions);
    },
    observeCandidates(candidates = [], observeOptions = {}) {
      return bridge.observeCandidates(candidates, observeOptions);
    },
    ingest(platformOrInput, payload, ingestOptions = {}) {
      return bridge.ingest(platformOrInput, payload, ingestOptions);
    },
    insertMark(input = {}, markOptions = {}) {
      return bridge.insertMark(input, markOptions);
    },
    insertAnnotation(input = {}, markOptions = {}) {
      return bridge.insertAnnotation(input, markOptions);
    },
    insertMarks(inputs = [], markOptions = {}) {
      return bridge.insertMarks(inputs, markOptions);
    },
    importTranscript(input = {}, transcriptOptions = {}) {
      return bridge.importTranscript(input, transcriptOptions);
    },
    startMeeting(input = {}) {
      return bridge.startMeeting(input);
    },
    endMeeting(input = {}) {
      return bridge.endMeeting(input);
    },
    handleWebhook(input = {}, webhookOptions = {}) {
      return router(input, webhookOptions);
    },
    handleFetchRequest(request, fetchOptions = {}) {
      return fetchHandler(request, fetchOptions);
    },
    handleNodeRequest(req, res, nodeOptions = {}) {
      return nodeHandler(req, res, nodeOptions);
    },
    captureWebhook(input = {}, captureOptions = {}) {
      return capturePlatformWebhookEvent(input, undefined, withDefaults(defaults, captureOptions));
    },
    captureFetchRequest(request, captureOptions = {}) {
      return capturePlatformWebRequest(request, withDefaults(defaults, captureOptions));
    },
    capturedSamples(records = [], sampleOptions = {}) {
      return buildPlatformCaptureSamples(records, withDefaults(defaults, sampleOptions));
    },
    capturedAcceptance(platform, records = [], reportOptions = {}) {
      return buildPlatformCaptureAcceptanceReport(platform, records, withDefaults(defaults, reportOptions));
    },
    capturedAcceptanceSummary(records = [], reportOptions = {}) {
      return buildMeetingPlatformCaptureAcceptanceSummary(records, withDefaults(defaults, reportOptions));
    },
    launchGate(platform, gateOptions = {}) {
      return buildPlatformLaunchGate(platform, withDefaults(defaults, gateOptions));
    },
    launchGateSummary(gateOptions = {}) {
      return buildMeetingPlatformLaunchGateSummary(withDefaults(defaults, gateOptions));
    },
    assertLaunchGate(platform, gateOptions = {}) {
      return assertPlatformLaunchGate(platform, withDefaults(defaults, gateOptions));
    },
    assertAllLaunchGates(gateOptions = {}) {
      return assertAllPlatformLaunchGates(withDefaults(defaults, gateOptions));
    },
    meetingAppLaunchGate(platform, gateOptions = {}) {
      return buildMeetingAppLaunchGate(platform, withDefaults(defaults, gateOptions));
    },
    meetingAppLaunchGateSummary(gateOptions = {}) {
      return buildMeetingAppLaunchGateSummary(withDefaults(defaults, gateOptions));
    },
    platformRolloutPlan(platform, rolloutOptions = {}) {
      return buildMeetingPlatformRolloutPlan(platform, withDefaults(defaults, rolloutOptions));
    },
    allPlatformRolloutPlans(rolloutOptions = {}) {
      return buildAllMeetingPlatformRolloutPlans(withDefaults(defaults, rolloutOptions));
    },
    platformRolloutSummary(rolloutOptions = {}) {
      return buildMeetingPlatformRolloutSummary(withDefaults(defaults, rolloutOptions));
    },
    platformAdaptationRunbook(platform, runbookOptions = {}) {
      return buildMeetingPlatformAdaptationRunbook(platform, withDefaults(defaults, runbookOptions));
    },
    allPlatformAdaptationRunbooks(runbookOptions = {}) {
      return buildAllMeetingPlatformAdaptationRunbooks(withDefaults(defaults, runbookOptions));
    },
    platformAdaptationRunbookSummary(runbookOptions = {}) {
      return buildMeetingPlatformAdaptationRunbookSummary(withDefaults(defaults, runbookOptions));
    },
    platformAdaptationStrategy(platform, strategyOptions = {}) {
      return buildMeetingPlatformAdaptationStrategy(platform, withDefaults(defaults, strategyOptions));
    },
    allPlatformAdaptationStrategies(strategyOptions = {}) {
      return buildAllMeetingPlatformAdaptationStrategies(withDefaults(defaults, strategyOptions));
    },
    platformAdaptationStrategyMatrix(strategyOptions = {}) {
      return buildMeetingPlatformAdaptationStrategyMatrix(withDefaults(defaults, strategyOptions));
    },
    platformEvidencePackage(platformOrInput, input = {}, packageOptions = {}) {
      if (platformOrInput && typeof platformOrInput === 'object' && !Array.isArray(platformOrInput)) {
        return buildMeetingPlatformEvidencePackage(platformOrInput, withDefaults(defaults, input));
      }
      return buildMeetingPlatformEvidencePackage(platformOrInput, input, withDefaults(defaults, packageOptions));
    },
    platformEvidencePackageSummary(packageOrInput, input = {}, summaryOptions = {}) {
      if (packageOrInput?.schema === 'meeting_platform_evidence_package') {
        return buildMeetingPlatformEvidencePackageSummary(packageOrInput, withDefaults(defaults, input));
      }
      if (packageOrInput && typeof packageOrInput === 'object' && !Array.isArray(packageOrInput)) {
        return buildMeetingPlatformEvidencePackageSummary(packageOrInput, withDefaults(defaults, input));
      }
      return buildMeetingPlatformEvidencePackageSummary(
        buildMeetingPlatformEvidencePackage(packageOrInput, input, withDefaults(defaults, summaryOptions)),
      );
    },
    platformEvidencePackageBuilder(platform, builderOptions = {}) {
      return createMeetingPlatformEvidencePackageBuilder(platform, withDefaults(defaults, builderOptions));
    },
    verifyPlatformEvidencePackage(packageOrInput, verifyOptions = {}) {
      return verifyMeetingPlatformEvidencePackage(packageOrInput, withDefaults(defaults, verifyOptions));
    },
    platformEvidenceCorrelation(platform, input = {}, correlationOptions = {}) {
      return buildMeetingPlatformEvidenceCorrelation(platform, input, withDefaults(defaults, correlationOptions));
    },
    platformEvidenceSession(platform, sessionOptions = {}) {
      return createMeetingPlatformEvidenceSession(platform, withDefaults(defaults, sessionOptions));
    },
    platformLiveAdapter(platform, adapterOptions = {}) {
      return createMeetingPlatformLiveAdapter(platform, bridge.client, withDefaults(defaults, adapterOptions));
    },
    platformLiveAdapterPlan(platform, planOptions = {}) {
      return buildMeetingPlatformLiveAdapterPlan(platform, withDefaults(defaults, planOptions));
    },
    platformLiveAdapterReadiness(platform, readinessOptions = {}) {
      return buildMeetingPlatformLiveAdapterReadiness(platform, withDefaults(defaults, readinessOptions));
    },
    assertPlatformLiveAdapterReadiness(platform, readinessOptions = {}) {
      return assertMeetingPlatformLiveAdapterReadiness(platform, withDefaults(defaults, readinessOptions));
    },
    platformLiveAdapterMatrix(matrixOptions = {}) {
      return buildMeetingPlatformLiveAdapterMatrix(withDefaults(defaults, matrixOptions));
    },
    platformLiveAdapterReadinessMatrix(matrixOptions = {}) {
      return buildMeetingPlatformLiveAdapterReadinessMatrix(withDefaults(defaults, matrixOptions));
    },
    assertPlatformLiveAdapterReadinessMatrix(matrixOptions = {}) {
      return assertMeetingPlatformLiveAdapterReadinessMatrix(withDefaults(defaults, matrixOptions));
    },
    platformLiveAdapterHandoff(platform, handoffOptions = {}) {
      return buildMeetingPlatformLiveAdapterHandoff(platform, withDefaults(defaults, handoffOptions));
    },
    platformLiveAdapterHandoffBundle(handoffOptions = {}) {
      return buildMeetingPlatformLiveAdapterHandoffBundle(withDefaults(defaults, handoffOptions));
    },
    platformLiveAdapterSuite(suiteOptions = {}) {
      return createMeetingPlatformLiveAdapterSuite(bridge.client, withDefaults(defaults, suiteOptions));
    },
    platformHostIntegrationPlan(hostOptions = {}) {
      return buildMeetingPlatformHostIntegrationPlan(withDefaults(defaults, hostOptions));
    },
    platformHostIntegrationScaffold(hostOptions = {}) {
      return buildMeetingPlatformHostIntegrationScaffold(withDefaults(defaults, hostOptions));
    },
    platformHostIntegrationAcceptance(scaffoldOrOptions = {}, acceptanceOptions = {}) {
      if (scaffoldOrOptions?.schema === 'meeting_platform_host_integration_scaffold') {
        return buildMeetingPlatformHostIntegrationScaffoldAcceptanceReport(
          scaffoldOrOptions,
          withDefaults(defaults, acceptanceOptions),
        );
      }
      return buildMeetingPlatformHostIntegrationScaffoldAcceptanceReport(withDefaults(defaults, scaffoldOrOptions));
    },
    assertPlatformHostIntegrationScaffold(scaffoldOrOptions = {}, acceptanceOptions = {}) {
      if (scaffoldOrOptions?.schema === 'meeting_platform_host_integration_scaffold') {
        return assertMeetingPlatformHostIntegrationScaffold(
          scaffoldOrOptions,
          withDefaults(defaults, acceptanceOptions),
        );
      }
      return assertMeetingPlatformHostIntegrationScaffold(withDefaults(defaults, scaffoldOrOptions));
    },
    platformProviderConnectionPack(platform, providerOptions = {}) {
      return buildMeetingPlatformProviderConnectionPack(platform, withDefaults(defaults, providerOptions));
    },
    platformProviderConnectionMatrix(providerOptions = {}) {
      return buildMeetingPlatformProviderConnectionMatrix(withDefaults(defaults, providerOptions));
    },
    platformSubscriptionHandoff(platform, handoffOptions = {}) {
      return buildMeetingPlatformSubscriptionHandoff(platform, withDefaults(defaults, handoffOptions));
    },
    platformSubscriptionHandoffMatrix(handoffOptions = {}) {
      return buildMeetingPlatformSubscriptionHandoffMatrix(withDefaults(defaults, handoffOptions));
    },
    assertPlatformSubscriptionHandoff(platform, handoffOptions = {}) {
      return assertMeetingPlatformSubscriptionHandoff(platform, withDefaults(defaults, handoffOptions));
    },
    assertPlatformSubscriptionHandoffMatrix(handoffOptions = {}) {
      return assertMeetingPlatformSubscriptionHandoffMatrix(withDefaults(defaults, handoffOptions));
    },
    platformSpeakerTrackPlan(platform, trackOptions = {}) {
      return buildMeetingPlatformSpeakerTrackPlan(platform, withDefaults(defaults, trackOptions));
    },
    platformSpeakerTrackMatrix(trackOptions = {}) {
      return buildMeetingPlatformSpeakerTrackMatrix(withDefaults(defaults, trackOptions));
    },
    platformSpeakerTrack(platform, input = {}, trackOptions = {}) {
      return buildMeetingPlatformSpeakerTrack(platform, input, withDefaults(defaults, trackOptions));
    },
    platformParticipantTrackPlan(platform, trackOptions = {}) {
      return buildMeetingPlatformParticipantTrackPlan(platform, withDefaults(defaults, trackOptions));
    },
    platformParticipantTrackMatrix(trackOptions = {}) {
      return buildMeetingPlatformParticipantTrackMatrix(withDefaults(defaults, trackOptions));
    },
    platformParticipantTrack(platform, input = {}, trackOptions = {}) {
      return buildMeetingPlatformParticipantTrack(platform, input, withDefaults(defaults, trackOptions));
    },
    meetingAppTrackPipeline(input = {}, trackOptions = {}) {
      return buildMeetingAppTrackPipeline(input, withDefaults(defaults, trackOptions));
    },
    createMeetingAppTrackPipeline(trackOptions = {}) {
      return createMeetingAppTrackPipeline(withDefaults(defaults, trackOptions));
    },
    meetingAppTrackRuntime(runtimeOptions = {}) {
      return createMeetingAppTrackRuntime(bridge.client, withDefaults(defaults, runtimeOptions));
    },
    platformTimelineViewPlan(platform, viewOptions = {}) {
      return buildMeetingPlatformTimelineViewPlan(platform, withDefaults(defaults, viewOptions));
    },
    platformTimelineViewMatrix(viewOptions = {}) {
      return buildMeetingPlatformTimelineViewMatrix(withDefaults(defaults, viewOptions));
    },
    platformTimelineView(platform, input = {}, viewOptions = {}) {
      return buildMeetingPlatformTimelineView(platform, input, withDefaults(defaults, viewOptions));
    },
    zoomPlatformTimelineViewport(viewport = {}, factor = 1, anchorRatio = 0.5, viewOptions = {}) {
      return zoomMeetingPlatformTimelineViewport(viewport, factor, anchorRatio, withDefaults(defaults, viewOptions));
    },
    platformAnnotationIntakePlan(platform, intakeOptions = {}) {
      return buildMeetingPlatformAnnotationIntakePlan(platform, withDefaults(defaults, intakeOptions));
    },
    platformAnnotationIntakeMatrix(intakeOptions = {}) {
      return buildMeetingPlatformAnnotationIntakeMatrix(withDefaults(defaults, intakeOptions));
    },
    platformAnnotationIntake(platform, input = {}, intakeOptions = {}) {
      return buildMeetingPlatformAnnotationIntake(platform, input, withDefaults(defaults, intakeOptions));
    },
    platformClockSyncPlan(platform, clockOptions = {}) {
      return buildMeetingPlatformClockSyncPlan(platform, withDefaults(defaults, clockOptions));
    },
    platformClockSyncMatrix(clockOptions = {}) {
      return buildMeetingPlatformClockSyncMatrix(withDefaults(defaults, clockOptions));
    },
    platformClockSync(platform, input = {}, clockOptions = {}) {
      return buildMeetingPlatformClockSyncReport(platform, input, withDefaults(defaults, clockOptions));
    },
    platformSessionBindingPlan(platform, bindingOptions = {}) {
      return buildMeetingPlatformSessionBindingPlan(platform, withDefaults(defaults, bindingOptions));
    },
    platformSessionBindingMatrix(bindingOptions = {}) {
      return buildMeetingPlatformSessionBindingMatrix(withDefaults(defaults, bindingOptions));
    },
    platformSessionBinding(platform, input = {}, bindingOptions = {}) {
      return buildMeetingPlatformSessionBinding(platform, input, withDefaults(defaults, bindingOptions));
    },
    platformRealtimeAnnotationPlan(platform, realtimeOptions = {}) {
      return buildMeetingPlatformRealtimeAnnotationPlan(platform, withDefaults(defaults, realtimeOptions));
    },
    platformRealtimeAnnotationMatrix(realtimeOptions = {}) {
      return buildMeetingPlatformRealtimeAnnotationMatrix(withDefaults(defaults, realtimeOptions));
    },
    platformRealtimeAnnotation(platform, input = {}, realtimeOptions = {}) {
      return buildMeetingPlatformRealtimeAnnotation(platform, input, withDefaults(defaults, realtimeOptions));
    },
    platformArtifactHandoffPlan(platform, artifactOptions = {}) {
      return buildMeetingPlatformArtifactHandoffPlan(platform, withDefaults(defaults, artifactOptions));
    },
    platformArtifactHandoffMatrix(artifactOptions = {}) {
      return buildMeetingPlatformArtifactHandoffMatrix(withDefaults(defaults, artifactOptions));
    },
    platformArtifactHandoff(platform, input = {}, artifactOptions = {}) {
      return buildMeetingPlatformArtifactHandoff(platform, input, withDefaults(defaults, artifactOptions));
    },
    platformRuntimeProfile(platform, profileOptions = {}) {
      return buildMeetingPlatformRuntimeProfile(platform, withDefaults(defaults, profileOptions));
    },
    platformRuntimeProfileMatrix(profileOptions = {}) {
      return buildMeetingPlatformRuntimeProfileMatrix(withDefaults(defaults, profileOptions));
    },
    platformAdaptationPackage(platform, packageOptions = {}) {
      return buildMeetingPlatformAdaptationPackage(platform, withDefaults(defaults, packageOptions));
    },
    platformAdaptationPackageMatrix(packageOptions = {}) {
      return buildMeetingPlatformAdaptationPackageMatrix(withDefaults(defaults, packageOptions));
    },
    platformRuntimeBundle(platform, bundleOptions = {}) {
      return buildMeetingPlatformRuntimeBundle(platform, withDefaults(defaults, bundleOptions));
    },
    platformRuntimeBundleMatrix(bundleOptions = {}) {
      return buildMeetingPlatformRuntimeBundleMatrix(withDefaults(defaults, bundleOptions));
    },
    platformRuntimeEventPlan(platform, planOptions = {}) {
      return buildMeetingPlatformRuntimeEventPlan(platform, withDefaults(defaults, planOptions));
    },
    platformRuntimeEventPlanMatrix(planOptions = {}) {
      return buildMeetingPlatformRuntimeEventPlanMatrix(withDefaults(defaults, planOptions));
    },
    platformAdapterContract(platform, contractOptions = {}) {
      return buildMeetingPlatformAdapterContract(platform, withDefaults(defaults, contractOptions));
    },
    platformAdapterContractMatrix(contractOptions = {}) {
      return buildMeetingPlatformAdapterContractMatrix(withDefaults(defaults, contractOptions));
    },
    platformAdapterContractAcceptance(contractOrPlatform, acceptanceOptions = {}) {
      return buildMeetingPlatformAdapterContractAcceptanceReport(contractOrPlatform, withDefaults(defaults, acceptanceOptions));
    },
    platformAdapterContractAcceptanceMatrix(acceptanceOptions = {}) {
      return buildMeetingPlatformAdapterContractAcceptanceMatrix(withDefaults(defaults, acceptanceOptions));
    },
    assertPlatformAdapterContract(contractOrPlatform, acceptanceOptions = {}) {
      return assertMeetingPlatformAdapterContract(contractOrPlatform, withDefaults(defaults, acceptanceOptions));
    },
    platformAdapterSamplePlan(platform, sampleOptions = {}) {
      return buildMeetingPlatformAdapterSamplePlan(platform, withDefaults(defaults, sampleOptions));
    },
    runPlatformAdapterSample(platform, sampleOptions = {}) {
      return runMeetingPlatformAdapterSample(platform, withDefaults(defaults, sampleOptions));
    },
    runPlatformAdapterSampleMatrix(sampleOptions = {}) {
      return runMeetingPlatformAdapterSampleMatrix(withDefaults(defaults, sampleOptions));
    },
    assertPlatformAdapterSample(sampleOrPlatform, sampleOptions = {}) {
      return assertMeetingPlatformAdapterSample(sampleOrPlatform, withDefaults(defaults, sampleOptions));
    },
    assertPlatformAdapterSampleMatrix(sampleOptions = {}) {
      return assertMeetingPlatformAdapterSampleMatrix(withDefaults(defaults, sampleOptions));
    },
    platformRealEvidenceIntakePlan(platform, intakeOptions = {}) {
      return buildMeetingPlatformRealEvidenceIntakePlan(platform, withDefaults(defaults, intakeOptions));
    },
    platformRealEvidenceIntake(platform, input = {}, intakeOptions = {}) {
      return buildMeetingPlatformRealEvidenceIntakeReport(platform, input, withDefaults(defaults, intakeOptions));
    },
    platformRealEvidenceIntakeMatrix(input = {}, intakeOptions = {}) {
      return buildMeetingPlatformRealEvidenceIntakeMatrix(input, withDefaults(defaults, intakeOptions));
    },
    assertPlatformRealEvidenceIntake(platform, input = {}, intakeOptions = {}) {
      return assertMeetingPlatformRealEvidenceIntake(platform, input, withDefaults(defaults, intakeOptions));
    },
    assertPlatformRealEvidenceIntakeMatrix(input = {}, intakeOptions = {}) {
      return assertMeetingPlatformRealEvidenceIntakeMatrix(input, withDefaults(defaults, intakeOptions));
    },
    platformFieldCapturePlan(platform, captureOptions = {}) {
      return buildMeetingPlatformFieldCapturePlan(platform, withDefaults(defaults, captureOptions));
    },
    platformFieldCaptureMatrix(captureOptions = {}) {
      return buildMeetingPlatformFieldCaptureMatrix(withDefaults(defaults, captureOptions));
    },
    platformFieldCaptureManifest(platform, manifestOptions = {}) {
      return buildMeetingPlatformFieldCaptureManifest(platform, withDefaults(defaults, manifestOptions));
    },
    platformFieldCaptureManifestMatrix(manifestOptions = {}) {
      return buildMeetingPlatformFieldCaptureManifestMatrix(withDefaults(defaults, manifestOptions));
    },
    platformFieldCollectorConfig(platform, collectorOptions = {}) {
      return buildMeetingPlatformFieldCollectorConfig(platform, withDefaults(defaults, collectorOptions));
    },
    platformFieldCollectorConfigMatrix(collectorOptions = {}) {
      return buildMeetingPlatformFieldCollectorConfigMatrix(withDefaults(defaults, collectorOptions));
    },
    platformFieldEvidenceBundle(platform, input = {}, bundleOptions = {}) {
      return buildMeetingPlatformFieldEvidenceBundle(platform, input, withDefaults(defaults, bundleOptions));
    },
    platformFieldEvidenceMatrix(matrixOptions = {}) {
      return buildMeetingPlatformFieldEvidenceMatrix(withDefaults(defaults, matrixOptions));
    },
    platformFieldIntakePlan(platform, intakeOptions = {}) {
      return buildMeetingPlatformFieldIntakePlan(platform, withDefaults(defaults, intakeOptions));
    },
    platformFieldIntakeMatrix(intakeOptions = {}) {
      return buildMeetingPlatformFieldIntakeMatrix(withDefaults(defaults, intakeOptions));
    },
    platformHandoffReadiness(platform, input = {}, readinessOptions = {}) {
      return buildMeetingPlatformHandoffReadiness(platform, input, withDefaults(defaults, readinessOptions));
    },
    platformHandoffReadinessMatrix(input = {}, readinessOptions = {}) {
      return buildMeetingPlatformHandoffReadinessMatrix(input, withDefaults(defaults, readinessOptions));
    },
    assertPlatformHandoffReadiness(platform, input = {}, readinessOptions = {}) {
      return assertMeetingPlatformHandoffReadiness(platform, input, withDefaults(defaults, readinessOptions));
    },
    assertPlatformHandoffReadinessMatrix(input = {}, readinessOptions = {}) {
      return assertMeetingPlatformHandoffReadinessMatrix(input, withDefaults(defaults, readinessOptions));
    },
    meetingAppSnapshotRecorder(recorderOptions = {}) {
      return createMeetingAppSnapshotRecorder(withDefaults(defaults, recorderOptions));
    },
    meetingAppGateInputFromRecords(records = [], recordOptions = {}) {
      return buildMeetingAppLaunchGateInputFromRecords(records, withDefaults(defaults, recordOptions));
    },
    meetingAppIntegrationProfile(platform, profileOptions = {}) {
      return buildMeetingAppIntegrationProfile(platform, withDefaults(defaults, profileOptions));
    },
    allMeetingAppIntegrationProfiles(profileOptions = {}) {
      return buildAllMeetingAppIntegrationProfiles(withDefaults(defaults, profileOptions));
    },
    meetingAppIntegrationMatrix(profileOptions = {}) {
      return buildMeetingAppIntegrationMatrix(withDefaults(defaults, profileOptions));
    },
    meetingAppDeploymentManifest(platform, manifestOptions = {}) {
      return buildMeetingAppDeploymentManifest(platform, withDefaults(defaults, manifestOptions));
    },
    allMeetingAppDeploymentManifests(manifestOptions = {}) {
      return buildAllMeetingAppDeploymentManifests(withDefaults(defaults, manifestOptions));
    },
    meetingAppDeploymentManifestAcceptance(manifestOrPlatform = {}, acceptanceOptions = {}) {
      return buildMeetingAppDeploymentManifestAcceptanceReport(manifestOrPlatform, withDefaults(defaults, acceptanceOptions));
    },
    allMeetingAppDeploymentManifestAcceptanceReports(acceptanceOptions = {}) {
      return buildAllMeetingAppDeploymentManifestAcceptanceReports(withDefaults(defaults, acceptanceOptions));
    },
    meetingAppDeploymentManifestAcceptanceSummary(acceptanceOptions = {}) {
      return buildMeetingAppDeploymentManifestAcceptanceSummary(withDefaults(defaults, acceptanceOptions));
    },
    assertMeetingAppDeploymentManifest(manifestOrPlatform = {}, acceptanceOptions = {}) {
      return assertMeetingAppDeploymentManifest(manifestOrPlatform, withDefaults(defaults, acceptanceOptions));
    },
    meetingAppLiveEvidencePackage(input = {}, evidenceOptions = {}) {
      return buildMeetingAppLiveEvidencePackage(input, withDefaults(defaults, evidenceOptions));
    },
    meetingAppLiveEvidencePackageSummary(input = {}, evidenceOptions = {}) {
      return buildMeetingAppLiveEvidencePackageSummary(input, withDefaults(defaults, evidenceOptions));
    },
    meetingAppDomAdaptationDiagnosis(platformOrInput = {}, diagnosisOptions = {}) {
      return buildMeetingAppDomAdaptationDiagnosis(platformOrInput, withDefaults(defaults, diagnosisOptions));
    },
    allMeetingAppDomAdaptationDiagnoses(diagnosisOptions = {}) {
      return buildAllMeetingAppDomAdaptationDiagnoses(withDefaults(defaults, diagnosisOptions));
    },
    meetingAppDomAdaptationDiagnosisMatrix(diagnosisOptions = {}) {
      return buildMeetingAppDomAdaptationDiagnosisMatrix(withDefaults(defaults, diagnosisOptions));
    },
    meetingAppRuntimeAdapterConfig(platform, configOptions = {}) {
      return buildMeetingAppRuntimeAdapterConfig(platform, withDefaults(defaults, configOptions));
    },
    meetingAppLiveSnapshotCapturePlan(platform, planOptions = {}) {
      return buildMeetingAppLiveSnapshotCapturePlan(platform, withDefaults(defaults, planOptions));
    },
    allMeetingAppLiveSnapshotCapturePlans(planOptions = {}) {
      return buildAllMeetingAppLiveSnapshotCapturePlans(withDefaults(defaults, planOptions));
    },
    allMeetingAppRuntimeAdapterConfigs(configOptions = {}) {
      return buildAllMeetingAppRuntimeAdapterConfigs(withDefaults(defaults, configOptions));
    },
    allMeetingAppRuntimeAdapterAcceptanceReports(acceptanceOptions = {}) {
      return buildAllMeetingAppRuntimeAdapterAcceptanceReports(withDefaults(defaults, acceptanceOptions));
    },
    meetingAppRuntimeAdapterAcceptance(configOrPlatform = {}, acceptanceOptions = {}) {
      return buildMeetingAppRuntimeAdapterAcceptanceReport(configOrPlatform, withDefaults(defaults, acceptanceOptions));
    },
    assertMeetingAppRuntimeAdapterConfig(configOrPlatform = {}, acceptanceOptions = {}) {
      return assertMeetingAppRuntimeAdapterConfig(configOrPlatform, withDefaults(defaults, acceptanceOptions));
    },
    meetingAppRuntimeAdapterValidation(configOrPlatform = {}, validationReportOptions = {}) {
      return buildMeetingAppRuntimeAdapterValidationReport(configOrPlatform, withDefaults(defaults, validationReportOptions));
    },
    allMeetingAppRuntimeAdapterValidationReports(validationReportOptions = {}) {
      return buildAllMeetingAppRuntimeAdapterValidationReports(withDefaults(defaults, validationReportOptions));
    },
    assertMeetingAppRuntimeAdapterValidation(configOrPlatform = {}, validationReportOptions = {}) {
      return assertMeetingAppRuntimeAdapterValidation(configOrPlatform, withDefaults(defaults, validationReportOptions));
    },
    meetingAppExtensionInstallPlan(extensionOptions = {}) {
      return buildMeetingAppExtensionInstallPlan(withDefaults(defaults, extensionOptions));
    },
    meetingAppContentScriptManifest(extensionOptions = {}) {
      return buildMeetingAppContentScriptManifest(withDefaults(defaults, extensionOptions));
    },
    meetingAppExtensionMatches(extensionOptions = {}) {
      return buildMeetingAppExtensionMatchPatterns(withDefaults(defaults, extensionOptions));
    },
    meetingAppExtensionScaffold(extensionOptions = {}) {
      return buildMeetingAppExtensionScaffold(withDefaults(defaults, extensionOptions));
    },
    meetingAppExtensionAcceptance(extensionOptions = {}) {
      return buildMeetingAppExtensionScaffoldAcceptanceReport(withDefaults(defaults, extensionOptions));
    },
    assertMeetingAppExtensionScaffold(extensionOptions = {}) {
      return assertMeetingAppExtensionScaffold(withDefaults(defaults, extensionOptions));
    },
    meetingAppExtensionAttachedMessage(input = {}, messageOptions = {}) {
      return buildMeetingAppExtensionAttachedMessage(input, withDefaults(defaults, messageOptions));
    },
    meetingAppExtensionStatusMessage(input = {}, messageOptions = {}) {
      return buildMeetingAppExtensionStatusMessage(input, withDefaults(defaults, messageOptions));
    },
    meetingAppExtensionObserveCandidatesMessage(input = {}, messageOptions = {}) {
      return buildMeetingAppExtensionObserveCandidatesMessage(input, withDefaults(defaults, messageOptions));
    },
    meetingAppExtensionClientCallMessage(methodOrInput, input = {}, messageOptions = {}) {
      const objectInput = Boolean(methodOrInput) && typeof methodOrInput === 'object' && !Array.isArray(methodOrInput);
      const resolvedInput = objectInput ? {} : input;
      const resolvedOptions = objectInput && arguments.length === 2 ? input : messageOptions;
      return buildMeetingAppExtensionClientCallMessage(methodOrInput, resolvedInput, withDefaults(defaults, resolvedOptions));
    },
    meetingAppExtensionTimelineEndpoint(method) {
      return meetingAppExtensionTimelineEndpoint(method);
    },
    normalizeMeetingAppExtensionMessageType(messageType) {
      return normalizeMeetingAppExtensionMessageType(messageType);
    },
    assertMeetingAppLaunchGate(platform, gateOptions = {}) {
      return assertMeetingAppLaunchGate(platform, withDefaults(defaults, gateOptions));
    },
    assertAllMeetingAppLaunchGates(gateOptions = {}) {
      return assertAllMeetingAppLaunchGates(withDefaults(defaults, gateOptions));
    },
    diagnose(platform, payload, diagnosticOptions = {}) {
      return diagnosePlatformEvent(platform, payload, diagnosticOptions);
    },
    platform(platform, platformOptions = {}) {
      return platformOverview(platform, defaults, platformOptions);
    },
    allPlatforms(platformOptions = {}) {
      return MEETING_PLATFORM_KEYS.map((platform) => platformOverview(platform, defaults, platformOptions));
    },
    platformRegistryEntry(platform, registryOptions = {}) {
      return buildMeetingPlatformRegistryEntry(platform, withDefaults(defaults, registryOptions));
    },
    platformRegistryManifest(registryOptions = {}) {
      return buildMeetingPlatformRegistryManifest(withDefaults(defaults, registryOptions));
    },
    platformRegistryAcceptance(manifestOrOptions = {}, registryOptions = {}) {
      return buildMeetingPlatformRegistryAcceptanceReport(
        manifestOrOptions,
        withDefaults(defaults, registryOptions),
      );
    },
    assertPlatformRegistryManifest(manifestOrOptions = {}, registryOptions = {}) {
      return assertMeetingPlatformRegistryManifest(
        manifestOrOptions,
        withDefaults(defaults, registryOptions),
      );
    },
    capability(platform, platformOptions = {}) {
      return platformCapabilityContract(platform, withDefaults(defaults, platformOptions));
    },
    routeTable(routeOptions = {}) {
      return buildPlatformWebhookRouteTable(withDefaults(defaults, routeOptions));
    },
    routerStatus(routeOptions = {}) {
      return buildPlatformWebhookRouterStatus(withDefaults(defaults, routeOptions));
    },
    routerSetup(routeOptions = {}) {
      return buildPlatformWebhookRouterSetup(withDefaults(defaults, routeOptions));
    },
    acceptance(platform, reportOptions = {}) {
      return buildPlatformAcceptanceReport(platform, acceptanceOptions(defaults, reportOptions));
    },
    fixtureAcceptance(platform, reportOptions = {}) {
      const fixtureInput = buildPlatformFixtureAcceptanceInput(withDefaults(defaults, reportOptions));
      return buildPlatformAcceptanceReport(platform, {
        ...fixtureInput,
        requireEndEvent: true,
        ...reportOptions,
        samples: fixtureInput.samples,
        env: fixtureInput.env,
      });
    },
    fixtureInput(fixtureOptions = {}) {
      return buildPlatformFixtureAcceptanceInput(withDefaults(defaults, fixtureOptions));
    },
    meetingAppFixture(platform, fixtureOptions = {}) {
      return buildMeetingAppFixtureSnapshot(platform, withDefaults(defaults, fixtureOptions));
    },
    allMeetingAppFixtures(fixtureOptions = {}) {
      return buildAllMeetingAppFixtureSnapshots(withDefaults(defaults, fixtureOptions));
    },
    diagnoseMeetingAppFixture(platform, diagnosticOptions = {}) {
      return diagnoseMeetingAppFixture(platform, withDefaults(defaults, diagnosticOptions));
    },
    diagnoseMeetingAppFixtureLifecycle(platform, diagnosticOptions = {}) {
      return diagnoseMeetingAppFixtureLifecycle(platform, withDefaults(defaults, diagnosticOptions));
    },
    meetingAppFixtureAcceptance(reportOptions = {}) {
      return buildMeetingAppFixtureAcceptanceReport(withDefaults(defaults, reportOptions));
    },
    onboarding(platform, onboardingOptions = {}) {
      return buildMeetingPlatformOnboardingReport(platform, withDefaults(defaults, onboardingOptions));
    },
    report(reportOptions = {}) {
      return buildMeetingPlatformKitReport(withDefaults(defaults, reportOptions));
    },
    getState() {
      return {
        bridge: bridge.getBridgeState(),
        webhook_router: router.getReconciliationState(),
      };
    },
    reset(nextState = {}) {
      return {
        bridge: bridge.reset(nextState.bridge ?? nextState),
        webhook_router: router.resetReconciliationState(nextState.webhook_router ?? nextState.webhookRouter ?? {}),
      };
    },
  };
}
