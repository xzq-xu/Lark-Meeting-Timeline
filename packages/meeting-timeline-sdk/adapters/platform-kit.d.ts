import type { MeetingTimelineClient, MeetingTimelineClientOptions, TimelineMarkInput, TranscriptImportInput } from '../index.mjs';
import type { MeetingTimelineBridge, MeetingTimelineBridgeOptions } from './timeline-bridge.mjs';
import type {
  PlatformWebhookRouterOptions,
  PlatformWebhookRouterRequest,
} from './platform-webhook-router.mjs';
import type { MeetingPlatformFetchHandler, PlatformWebhookHttpOptions } from './platform-http.mjs';
import type { MeetingPlatformNodeHandler, PlatformWebhookNodeOptions } from './platform-node.mjs';
import type {
  MeetingPlatformRegistryAcceptanceReport,
  MeetingPlatformRegistryEntry,
  MeetingPlatformRegistryManifest,
  MeetingPlatformRegistryOptions,
} from './platform-registry.mjs';
import type {
  MeetingPlatformConnector,
  MeetingPlatformConnectorAcceptanceReport,
  MeetingPlatformConnectorBrowserRuntime,
  MeetingPlatformConnectorContentScriptBridge,
  MeetingPlatformConnectorHub,
  MeetingPlatformConnectorHubRuntime,
  MeetingPlatformConnectorMatrix,
  MeetingPlatformConnectorOptions,
  MeetingPlatformConnectorResolution,
  MeetingPlatformConnectorRuntime,
} from './meeting-platform-connector.mjs';
import type { PlatformCaptureOptions, PlatformCaptureRecord } from './platform-capture.mjs';
import type { PlatformLaunchGateOptions } from './platform-gate.mjs';
import type {
  MeetingPlatformAdaptationRunbook,
  MeetingPlatformAdaptationRunbookSummary,
  MeetingPlatformRolloutOptions,
  MeetingPlatformRolloutPlan,
  MeetingPlatformRolloutSummary,
} from './platform-rollout.mjs';
import type {
  MeetingPlatformEvidencePackage,
  MeetingPlatformEvidencePackageBuilder,
  MeetingPlatformEvidencePackageOptions,
  MeetingPlatformEvidencePackageSummary,
  MeetingPlatformEvidencePackageVerification,
} from './platform-evidence-package.mjs';
import type {
  MeetingPlatformEvidenceCorrelation,
  MeetingPlatformEvidenceCorrelationInput,
  MeetingPlatformEvidenceCorrelationOptions,
} from './platform-evidence-correlation.mjs';
import type {
  MeetingPlatformEvidenceSession,
  MeetingPlatformEvidenceSessionOptions,
} from './platform-evidence-session.mjs';
import type {
  MeetingPlatformLiveAdapter,
  MeetingPlatformLiveAdapterHandoff,
  MeetingPlatformLiveAdapterHandoffBundle,
  MeetingPlatformLiveAdapterMatrix,
  MeetingPlatformLiveAdapterOptions,
  MeetingPlatformLiveAdapterPlan,
  MeetingPlatformLiveAdapterReadiness,
  MeetingPlatformLiveAdapterReadinessMatrix,
  MeetingPlatformLiveAdapterReadinessOptions,
  MeetingPlatformLiveAdapterSuite,
} from './platform-live-adapter.mjs';
import type {
  MeetingPlatformAdaptationStrategy,
  MeetingPlatformAdaptationStrategyMatrix,
} from './platform-strategy.mjs';
import type {
  MeetingPlatformAdapterRoute,
  MeetingPlatformAdapterRouteMatrix,
  MeetingPlatformAdapterRouteOptions,
} from './platform-adapter-route.mjs';
import type {
  MeetingPlatformHostIntegrationOptions,
  MeetingPlatformHostIntegrationPlan,
  MeetingPlatformHostIntegrationScaffold,
  MeetingPlatformHostIntegrationScaffoldAcceptanceReport,
} from './platform-host-integration.mjs';
import type {
  MeetingPlatformProviderConnectionMatrix,
  MeetingPlatformProviderConnectionOptions,
  MeetingPlatformProviderConnectionPack,
} from './platform-provider-connection.mjs';
import type {
  MeetingPlatformSubscriptionHandoff,
  MeetingPlatformSubscriptionHandoffMatrix,
  MeetingPlatformSubscriptionHandoffOptions,
} from './platform-subscription-handoff.mjs';
import type {
  MeetingPlatformSpeakerTrack,
  MeetingPlatformSpeakerTrackInput,
  MeetingPlatformSpeakerTrackMatrix,
  MeetingPlatformSpeakerTrackOptions,
  MeetingPlatformSpeakerTrackPlan,
} from './platform-speaker-track.mjs';
import type {
  MeetingPlatformParticipantTrack,
  MeetingPlatformParticipantTrackInput,
  MeetingPlatformParticipantTrackMatrix,
  MeetingPlatformParticipantTrackOptions,
  MeetingPlatformParticipantTrackPlan,
} from './platform-participant-track.mjs';
import type {
  MeetingPlatformTimelineView,
  MeetingPlatformTimelineViewInput,
  MeetingPlatformTimelineViewMatrix,
  MeetingPlatformTimelineViewOptions,
  MeetingPlatformTimelineViewPlan,
  MeetingPlatformTimelineViewport,
} from './platform-timeline-view.mjs';
import type {
  MeetingPlatformAnnotationIntake,
  MeetingPlatformAnnotationIntakeInput,
  MeetingPlatformAnnotationIntakeMatrix,
  MeetingPlatformAnnotationIntakeOptions,
  MeetingPlatformAnnotationIntakePlan,
} from './platform-annotation-intake.mjs';
import type {
  MeetingPlatformClockSyncInput,
  MeetingPlatformClockSyncMatrix,
  MeetingPlatformClockSyncOptions,
  MeetingPlatformClockSyncPlan,
  MeetingPlatformClockSyncReport,
} from './platform-clock-sync.mjs';
import type {
  MeetingPlatformSessionBinding,
  MeetingPlatformSessionBindingInput,
  MeetingPlatformSessionBindingMatrix,
  MeetingPlatformSessionBindingOptions,
  MeetingPlatformSessionBindingPlan,
} from './platform-session-binding.mjs';
import type {
  MeetingPlatformRealtimeAnnotation,
  MeetingPlatformRealtimeAnnotationInput,
  MeetingPlatformRealtimeAnnotationMatrix,
  MeetingPlatformRealtimeAnnotationOptions,
  MeetingPlatformRealtimeAnnotationPlan,
} from './platform-realtime-annotation.mjs';
import type {
  MeetingPlatformArtifactHandoff,
  MeetingPlatformArtifactHandoffInput,
  MeetingPlatformArtifactHandoffMatrix,
  MeetingPlatformArtifactHandoffOptions,
  MeetingPlatformArtifactHandoffPlan,
} from './platform-artifact-handoff.mjs';
import type {
  MeetingPlatformRuntimeProfile,
  MeetingPlatformRuntimeProfileMatrix,
  MeetingPlatformRuntimeProfileOptions,
} from './platform-runtime-profile.mjs';
import type {
  MeetingPlatformAdaptationPackage,
  MeetingPlatformAdaptationPackageMatrix,
  MeetingPlatformAdaptationPackageOptions,
} from './platform-adaptation-package.mjs';
import type {
  MeetingPlatformRuntimeBundle,
  MeetingPlatformRuntimeBundleMatrix,
  MeetingPlatformRuntimeBundleOptions,
} from './platform-runtime-bundle.mjs';
import type {
  MeetingPlatformRuntimeHost,
  MeetingPlatformRuntimeHostConfig,
  MeetingPlatformRuntimeHostConfigMatrix,
  MeetingPlatformRuntimeHostHandoff,
  MeetingPlatformRuntimeHostHandoffMatrix,
  MeetingPlatformRuntimeHostOptions,
} from './meeting-platform-runtime-host.mjs';
import type {
  MeetingPlatformRuntimeHostFixtureEnvironment,
  MeetingPlatformRuntimeHostReplayMatrix,
  MeetingPlatformRuntimeHostReplayReport,
  MeetingPlatformRuntimeHostVerificationClient,
  MeetingPlatformRuntimeHostVerificationMatrix,
  MeetingPlatformRuntimeHostVerificationReport,
} from './meeting-platform-runtime-host-verifier.mjs';
import type {
  MeetingPlatformRuntimeEventClientOptions,
  MeetingPlatformRuntimeEventPlan,
  MeetingPlatformRuntimeEventPlanMatrix,
} from './platform-runtime-event.mjs';
import type {
  MeetingPlatformAdapterContract,
  MeetingPlatformAdapterContractAcceptanceMatrix,
  MeetingPlatformAdapterContractAcceptanceReport,
  MeetingPlatformAdapterContractMatrix,
  MeetingPlatformAdapterContractOptions,
} from './platform-adapter-contract.mjs';
import type {
  MeetingPlatformAdapterSample,
  MeetingPlatformAdapterSampleMatrix,
  MeetingPlatformAdapterSampleOptions,
  MeetingPlatformAdapterSamplePlan,
} from './platform-adapter-sample.mjs';
import type {
  MeetingPlatformRealEvidenceIntakeMatrix,
  MeetingPlatformRealEvidenceIntakeOptions,
  MeetingPlatformRealEvidenceIntakePlan,
  MeetingPlatformRealEvidenceIntakeReport,
} from './platform-real-intake.mjs';
import type {
  MeetingPlatformFieldCaptureManifest,
  MeetingPlatformFieldCaptureManifestMatrix,
  MeetingPlatformFieldCollectorConfig,
  MeetingPlatformFieldCollectorConfigMatrix,
  MeetingPlatformFieldEvidenceBundle,
  MeetingPlatformFieldEvidenceMatrix,
  MeetingPlatformFieldCaptureMatrix,
  MeetingPlatformFieldCaptureOptions,
  MeetingPlatformFieldCapturePlan,
} from './platform-field-capture.mjs';
import type {
  MeetingPlatformFieldIntakeMatrix,
  MeetingPlatformFieldIntakeOptions,
  MeetingPlatformFieldIntakePlan,
} from './platform-field-intake.mjs';
import type {
  MeetingPlatformHandoffReadiness,
  MeetingPlatformHandoffReadinessMatrix,
  MeetingPlatformHandoffReadinessOptions,
} from './platform-handoff-readiness.mjs';
import type {
  MeetingPlatformConformanceOptions,
  MeetingPlatformConformanceReport,
} from './platform-conformance.mjs';
import type {
  MeetingPlatformConsumerHandoff,
  MeetingPlatformConsumerHandoffOptions,
} from './platform-consumer-handoff.mjs';
import type {
  MeetingAppAdapterManifest,
  MeetingAppAdapterManifestMatrix,
  MeetingAppAdapterManifestOptions,
} from './meeting-app-adapter-manifest.mjs';
import type {
  MeetingAppAdapterSpec,
  MeetingAppAdapterSpecInput,
  MeetingAppAdapterSpecMatrix,
  MeetingAppAdapterSpecMatrixOptions,
} from './meeting-app-adapter-spec.mjs';
import type {
  MeetingAppAdapterRuntimeConfig,
  MeetingAppAdapterRuntimeConfigMatrix,
  MeetingAppAdapterRuntimeConfigMatrixOptions,
  MeetingAppAdapterRuntimeConfigOptions,
} from './meeting-app-adapter-runtime-config.mjs';
import type {
  MeetingAppAdapterHandoffPackage,
  MeetingAppAdapterHandoffPackageMatrix,
  MeetingAppAdapterHandoffPackageMatrixOptions,
  MeetingAppAdapterHandoffPackageOptions,
  MeetingAppAdapterVerificationMatrixOptions,
  MeetingAppAdapterVerificationOptions,
  MeetingAppAdapterVerificationReport,
  MeetingAppAdapterVerificationReportMatrix,
} from './meeting-app-adapter-handoff-package.mjs';
import type { MeetingAppLaunchGate, MeetingAppLaunchGateOptions, MeetingAppLaunchGateSummary } from './meeting-app-gate.mjs';
import type { MeetingAppSnapshotRecorder, MeetingAppSnapshotRecord, MeetingAppSnapshotRecordOptions, MeetingAppSnapshotRecordSet } from './meeting-app-snapshot-recorder.mjs';
import type {
  MeetingAppDeploymentManifest,
  MeetingAppDeploymentManifestAcceptanceReport,
  MeetingAppDeploymentManifestAcceptanceSummary,
  MeetingAppDomAdaptationDiagnosis,
  MeetingAppDomAdaptationDiagnosisMatrix,
  MeetingAppIntegrationMatrix,
  MeetingAppIntegrationProfile,
  MeetingAppIntegrationProfileOptions,
  MeetingAppLiveEvidencePackage,
  MeetingAppLiveEvidencePackageSummary,
  MeetingAppLiveSnapshotCapturePlan,
  MeetingAppRuntimeAdapterAcceptanceReport,
  MeetingAppRuntimeAdapterConfig,
  MeetingAppRuntimeAdapterHandoffAcceptanceReport,
  MeetingAppRuntimeAdapterHandoff,
  MeetingAppRuntimeAdapterHandoffMatrixAcceptanceReport,
  MeetingAppRuntimeAdapterHandoffMatrix,
  MeetingAppRuntimeAdapterHostPackage,
  MeetingAppRuntimeAdapterProfileMatrix,
  MeetingAppRuntimeAdapterProfileResolution,
  MeetingAppRuntimeAdapterSelection,
  MeetingAppRuntimeAdapterValidationReport,
  MeetingAppRuntimeObserverPlan,
  MeetingAppRuntimeObserverPlanMatrix,
} from './meeting-app-profile.mjs';
import type {
  MeetingAppExtensionInstallPlan,
  MeetingAppExtensionMatchPatterns,
  MeetingAppExtensionClientCallMessageInput,
  MeetingAppExtensionClientCallMethod,
  MeetingAppExtensionMessageOptions,
  MeetingAppExtensionMessageType,
  MeetingAppExtensionOptions,
  MeetingAppExtensionScaffold,
  MeetingAppExtensionScaffoldAcceptanceReport,
} from './meeting-app-extension.mjs';
import type {
  MeetingAppObserverSchedulerConfig,
  MeetingAppObserverSchedulerConfigMatrix,
  MeetingAppObserverSchedulerOptions,
} from './meeting-app-observer-scheduler.mjs';
import type { PlatformEventDiagnosticResult, PlatformEventIngestInput, PlatformEventIngestOptions, ReconciledPlatformEventIngestResult } from './platform-ingest.mjs';
import type {
  MeetingAppFixtureDiagnosis,
  MeetingAppFixtureLifecycleDiagnosis,
  MeetingAppFixtureOptions,
  MeetingAppFixturePlatform,
} from './meeting-app-fixtures.mjs';
import type {
  MeetingAppAdapterFitMatrix,
  MeetingAppAdapterFitReport,
  MeetingAppSnapshot,
} from './meeting-apps.mjs';
import type {
  MeetingAppAdapterCapabilityMatrix,
  MeetingAppAdapterCapabilityOptions,
  MeetingAppAdapterCapabilityReport,
  MeetingAppAdapterExecutionPlan,
  MeetingAppAdapterExecutionPlanMatrix,
} from './meeting-app-adapter-capability.mjs';
import type {
  MeetingAppAdapterIntegrationPackage,
  MeetingAppAdapterIntegrationPackageMatrix,
} from './meeting-app-adapter-integration-package.mjs';
import type {
  MeetingAppTrackPipeline,
  MeetingAppTrackPipelineAccumulator,
  MeetingAppTrackPipelineInput,
  MeetingAppTrackPipelineOptions,
} from './meeting-app-track-pipeline.mjs';
import type {
  MeetingAppTrackRuntime,
  MeetingAppTrackRuntimeOptions,
} from './meeting-app-track-runtime.mjs';

export interface MeetingPlatformTimelineKitOptions extends MeetingTimelineBridgeOptions, PlatformWebhookRouterOptions {
  baseUrl?: string;
  base_url?: string;
  clientOptions?: MeetingTimelineClientOptions;
  client_options?: MeetingTimelineClientOptions;
  env?: Record<string, unknown>;
  verify?: boolean;
  reconcile?: boolean;
  reconciled?: boolean;
  [key: string]: unknown;
}

export interface MeetingPlatformTimelineKit {
  client: MeetingTimelineClient;
  bridge: MeetingTimelineBridge;
  webhookRouter: unknown;
  fetchHandler: MeetingPlatformFetchHandler;
  nodeHandler: MeetingPlatformNodeHandler;
  platforms: readonly string[];
  observe(snapshot?: Record<string, unknown>, observeOptions?: Record<string, unknown>): Promise<unknown>;
  observeCandidates(candidates?: unknown[] | Record<string, unknown>, observeOptions?: Record<string, unknown>): Promise<unknown>;
  ingest(platformOrInput: string | PlatformEventIngestInput, payload?: unknown, ingestOptions?: PlatformEventIngestOptions): Promise<ReconciledPlatformEventIngestResult>;
  insertMark(input?: TimelineMarkInput, markOptions?: Record<string, unknown>): Promise<unknown>;
  insertAnnotation(input?: TimelineMarkInput, markOptions?: Record<string, unknown>): Promise<unknown>;
  insertMarks(inputs?: TimelineMarkInput[] | Record<string, unknown>, markOptions?: Record<string, unknown>): Promise<unknown>;
  importTranscript(input?: TranscriptImportInput & { raw?: unknown; content?: unknown; platform?: string }, transcriptOptions?: Record<string, unknown>): Promise<unknown>;
  startMeeting(input?: Record<string, unknown>): Promise<unknown>;
  endMeeting(input?: Record<string, unknown>): Promise<unknown>;
  handleWebhook(input?: PlatformWebhookRouterRequest, webhookOptions?: PlatformWebhookRouterOptions): Promise<Record<string, unknown>>;
  handleFetchRequest(request: Request, fetchOptions?: PlatformWebhookHttpOptions): Promise<Response>;
  handleNodeRequest(req: Record<string, unknown>, res?: Record<string, unknown>, nodeOptions?: PlatformWebhookNodeOptions): Promise<unknown>;
  captureWebhook(input?: PlatformWebhookRouterRequest, captureOptions?: PlatformCaptureOptions): PlatformCaptureRecord;
  captureFetchRequest(request: Request, captureOptions?: PlatformCaptureOptions): Promise<PlatformCaptureRecord>;
  capturedSamples(records?: PlatformCaptureRecord[], sampleOptions?: Record<string, unknown>): Record<string, Record<string, unknown>[]>;
  capturedAcceptance(platform: string, records?: PlatformCaptureRecord[], reportOptions?: Record<string, unknown>): Record<string, unknown>;
  capturedAcceptanceSummary(records?: PlatformCaptureRecord[], reportOptions?: Record<string, unknown>): Record<string, unknown>;
  launchGate(platform: string, gateOptions?: PlatformLaunchGateOptions): Record<string, unknown>;
  launchGateSummary(gateOptions?: PlatformLaunchGateOptions): Record<string, unknown>;
  assertLaunchGate(platform: string, gateOptions?: PlatformLaunchGateOptions): Record<string, unknown>;
  assertAllLaunchGates(gateOptions?: PlatformLaunchGateOptions): Record<string, unknown>;
  meetingAppLaunchGate(platform: string, gateOptions?: MeetingAppLaunchGateOptions): MeetingAppLaunchGate;
  meetingAppLaunchGateSummary(gateOptions?: MeetingAppLaunchGateOptions): MeetingAppLaunchGateSummary;
  platformRolloutPlan(platform: string, rolloutOptions?: MeetingPlatformRolloutOptions): MeetingPlatformRolloutPlan;
  allPlatformRolloutPlans(rolloutOptions?: MeetingPlatformRolloutOptions): MeetingPlatformRolloutPlan[];
  platformRolloutSummary(rolloutOptions?: MeetingPlatformRolloutOptions): MeetingPlatformRolloutSummary;
  platformAdaptationRunbook(platform: string, runbookOptions?: MeetingPlatformRolloutOptions): MeetingPlatformAdaptationRunbook;
  allPlatformAdaptationRunbooks(runbookOptions?: MeetingPlatformRolloutOptions): MeetingPlatformAdaptationRunbook[];
  platformAdaptationRunbookSummary(runbookOptions?: MeetingPlatformRolloutOptions): MeetingPlatformAdaptationRunbookSummary;
  platformAdaptationStrategy(platform: string, strategyOptions?: MeetingPlatformRolloutOptions): MeetingPlatformAdaptationStrategy;
  allPlatformAdaptationStrategies(strategyOptions?: MeetingPlatformRolloutOptions): MeetingPlatformAdaptationStrategy[];
  platformAdaptationStrategyMatrix(strategyOptions?: MeetingPlatformRolloutOptions): MeetingPlatformAdaptationStrategyMatrix;
  platformAdapterRoute(platform: string, routeOptions?: MeetingPlatformAdapterRouteOptions): MeetingPlatformAdapterRoute;
  platformAdapterRouteMatrix(routeOptions?: MeetingPlatformAdapterRouteOptions): MeetingPlatformAdapterRouteMatrix;
  platformConformance(conformanceOptions?: MeetingPlatformConformanceOptions): MeetingPlatformConformanceReport;
  assertPlatformConformance(conformanceOptions?: MeetingPlatformConformanceOptions): MeetingPlatformConformanceReport;
  platformConsumerHandoff(handoffOptions?: MeetingPlatformConsumerHandoffOptions): MeetingPlatformConsumerHandoff;
  assertPlatformConsumerHandoff(handoffOptions?: MeetingPlatformConsumerHandoffOptions): MeetingPlatformConsumerHandoff;
  platformEvidencePackage(platformOrInput: string | MeetingPlatformEvidencePackageOptions, input?: MeetingPlatformEvidencePackageOptions, packageOptions?: MeetingPlatformEvidencePackageOptions): MeetingPlatformEvidencePackage;
  platformEvidencePackageSummary(packageOrInput: MeetingPlatformEvidencePackage | string | MeetingPlatformEvidencePackageOptions, input?: MeetingPlatformEvidencePackageOptions, summaryOptions?: MeetingPlatformEvidencePackageOptions): MeetingPlatformEvidencePackageSummary;
  platformEvidencePackageBuilder(platform: string, builderOptions?: MeetingPlatformEvidencePackageOptions): MeetingPlatformEvidencePackageBuilder;
  verifyPlatformEvidencePackage(packageOrInput: MeetingPlatformEvidencePackage | string | MeetingPlatformEvidencePackageOptions, verifyOptions?: MeetingPlatformEvidencePackageOptions): MeetingPlatformEvidencePackageVerification;
  platformEvidenceCorrelation(platform: string, input?: MeetingPlatformEvidenceCorrelationInput, correlationOptions?: MeetingPlatformEvidenceCorrelationOptions): MeetingPlatformEvidenceCorrelation;
  platformEvidenceSession(platform: string, sessionOptions?: MeetingPlatformEvidenceSessionOptions): MeetingPlatformEvidenceSession;
  platformLiveAdapter(platform: string, adapterOptions?: MeetingPlatformLiveAdapterOptions): MeetingPlatformLiveAdapter;
  platformLiveAdapterPlan(platform: string, planOptions?: MeetingPlatformLiveAdapterOptions): MeetingPlatformLiveAdapterPlan;
  platformLiveAdapterReadiness(platform: string, readinessOptions?: MeetingPlatformLiveAdapterReadinessOptions): MeetingPlatformLiveAdapterReadiness;
  assertPlatformLiveAdapterReadiness(platform: string, readinessOptions?: MeetingPlatformLiveAdapterReadinessOptions): MeetingPlatformLiveAdapterReadiness;
  platformLiveAdapterMatrix(matrixOptions?: MeetingPlatformLiveAdapterOptions): MeetingPlatformLiveAdapterMatrix;
  platformLiveAdapterReadinessMatrix(matrixOptions?: MeetingPlatformLiveAdapterReadinessOptions): MeetingPlatformLiveAdapterReadinessMatrix;
  assertPlatformLiveAdapterReadinessMatrix(matrixOptions?: MeetingPlatformLiveAdapterReadinessOptions): MeetingPlatformLiveAdapterReadinessMatrix;
  platformLiveAdapterHandoff(platform: string, handoffOptions?: MeetingPlatformLiveAdapterReadinessOptions): MeetingPlatformLiveAdapterHandoff;
  platformLiveAdapterHandoffBundle(handoffOptions?: MeetingPlatformLiveAdapterReadinessOptions): MeetingPlatformLiveAdapterHandoffBundle;
  platformLiveAdapterSuite(suiteOptions?: MeetingPlatformLiveAdapterOptions): MeetingPlatformLiveAdapterSuite;
  platformHostIntegrationPlan(hostOptions?: MeetingPlatformHostIntegrationOptions): MeetingPlatformHostIntegrationPlan;
  platformHostIntegrationScaffold(hostOptions?: MeetingPlatformHostIntegrationOptions): MeetingPlatformHostIntegrationScaffold;
  platformHostIntegrationAcceptance(
    scaffoldOrOptions?: MeetingPlatformHostIntegrationScaffold | MeetingPlatformHostIntegrationOptions,
    acceptanceOptions?: MeetingPlatformHostIntegrationOptions,
  ): MeetingPlatformHostIntegrationScaffoldAcceptanceReport;
  assertPlatformHostIntegrationScaffold(
    scaffoldOrOptions?: MeetingPlatformHostIntegrationScaffold | MeetingPlatformHostIntegrationOptions,
    acceptanceOptions?: MeetingPlatformHostIntegrationOptions,
  ): MeetingPlatformHostIntegrationScaffoldAcceptanceReport;
  platformProviderConnectionPack(platform: string, providerOptions?: MeetingPlatformProviderConnectionOptions): MeetingPlatformProviderConnectionPack;
  platformProviderConnectionMatrix(providerOptions?: MeetingPlatformProviderConnectionOptions): MeetingPlatformProviderConnectionMatrix;
  platformSubscriptionHandoff(platform: string, handoffOptions?: MeetingPlatformSubscriptionHandoffOptions): MeetingPlatformSubscriptionHandoff;
  platformSubscriptionHandoffMatrix(handoffOptions?: MeetingPlatformSubscriptionHandoffOptions): MeetingPlatformSubscriptionHandoffMatrix;
  assertPlatformSubscriptionHandoff(platform: string, handoffOptions?: MeetingPlatformSubscriptionHandoffOptions): MeetingPlatformSubscriptionHandoff;
  assertPlatformSubscriptionHandoffMatrix(handoffOptions?: MeetingPlatformSubscriptionHandoffOptions): MeetingPlatformSubscriptionHandoffMatrix;
  platformSpeakerTrackPlan(platform: string, trackOptions?: MeetingPlatformSpeakerTrackOptions): MeetingPlatformSpeakerTrackPlan;
  platformSpeakerTrackMatrix(trackOptions?: MeetingPlatformSpeakerTrackOptions): MeetingPlatformSpeakerTrackMatrix;
  platformSpeakerTrack(
    platform: string,
    input?: MeetingPlatformSpeakerTrackInput | Record<string, unknown>[],
    trackOptions?: MeetingPlatformSpeakerTrackOptions,
  ): MeetingPlatformSpeakerTrack;
  platformParticipantTrackPlan(platform: string, trackOptions?: MeetingPlatformParticipantTrackOptions): MeetingPlatformParticipantTrackPlan;
  platformParticipantTrackMatrix(trackOptions?: MeetingPlatformParticipantTrackOptions): MeetingPlatformParticipantTrackMatrix;
  platformParticipantTrack(
    platform: string,
    input?: MeetingPlatformParticipantTrackInput | Record<string, unknown>[],
    trackOptions?: MeetingPlatformParticipantTrackOptions,
  ): MeetingPlatformParticipantTrack;
  meetingAppAdapterFit(
    input?: MeetingAppSnapshot | MeetingAppSnapshot[] | Record<string, unknown>,
    fitOptions?: Record<string, unknown>,
  ): MeetingAppAdapterFitReport;
  meetingAppAdapterFitMatrix(
    input?: MeetingAppSnapshot | MeetingAppSnapshot[] | Record<string, unknown>,
    fitOptions?: Record<string, unknown>,
  ): MeetingAppAdapterFitMatrix;
  meetingAppAdapterCapability(
    platformOrOptions?: string | MeetingAppAdapterCapabilityOptions,
    input?: MeetingAppSnapshot | MeetingAppSnapshot[] | Record<string, unknown>,
    capabilityOptions?: MeetingAppAdapterCapabilityOptions,
  ): MeetingAppAdapterCapabilityReport;
  meetingAppAdapterCapabilityMatrix(capabilityOptions?: MeetingAppAdapterCapabilityOptions): MeetingAppAdapterCapabilityMatrix;
  meetingAppAdapterExecutionPlan(
    platformOrCapability?: string | MeetingAppAdapterCapabilityReport | MeetingAppAdapterCapabilityOptions,
    input?: MeetingAppSnapshot | MeetingAppSnapshot[] | Record<string, unknown> | MeetingAppAdapterCapabilityOptions,
    planOptions?: MeetingAppAdapterCapabilityOptions,
  ): MeetingAppAdapterExecutionPlan;
  meetingAppAdapterExecutionPlanMatrix(planOptions?: MeetingAppAdapterCapabilityOptions): MeetingAppAdapterExecutionPlanMatrix;
  meetingAppAdapterIntegrationPackage(
    platformOrCapability?: string | MeetingAppAdapterCapabilityReport | MeetingAppAdapterExecutionPlan | MeetingAppAdapterCapabilityOptions,
    input?: MeetingAppSnapshot | MeetingAppSnapshot[] | Record<string, unknown> | MeetingAppAdapterCapabilityOptions,
    packageOptions?: MeetingAppAdapterCapabilityOptions,
  ): MeetingAppAdapterIntegrationPackage;
  meetingAppAdapterIntegrationPackageMatrix(packageOptions?: MeetingAppAdapterCapabilityOptions): MeetingAppAdapterIntegrationPackageMatrix;
  assertMeetingAppAdapterIntegrationPackage(
    packageOrPlatform?: MeetingAppAdapterIntegrationPackage | string | MeetingAppAdapterCapabilityReport | MeetingAppAdapterExecutionPlan | MeetingAppAdapterCapabilityOptions,
    input?: MeetingAppSnapshot | MeetingAppSnapshot[] | Record<string, unknown> | MeetingAppAdapterCapabilityOptions,
    packageOptions?: MeetingAppAdapterCapabilityOptions,
  ): MeetingAppAdapterIntegrationPackage;
  assertMeetingAppAdapterIntegrationPackageMatrix(
    matrixOrOptions?: MeetingAppAdapterIntegrationPackageMatrix | MeetingAppAdapterCapabilityOptions,
    packageOptions?: MeetingAppAdapterCapabilityOptions,
  ): MeetingAppAdapterIntegrationPackageMatrix;
  meetingAppAdapterManifest(platform: string, manifestOptions?: MeetingAppAdapterManifestOptions): MeetingAppAdapterManifest;
  meetingAppAdapterManifestMatrix(manifestOptions?: MeetingAppAdapterManifestOptions): MeetingAppAdapterManifestMatrix;
  assertMeetingAppAdapterManifest(
    manifestOrPlatform?: MeetingAppAdapterManifest | string | MeetingAppAdapterManifestOptions,
    manifestOptions?: MeetingAppAdapterManifestOptions,
  ): MeetingAppAdapterManifest;
  assertMeetingAppAdapterManifestMatrix(
    matrixOrOptions?: MeetingAppAdapterManifestMatrix | MeetingAppAdapterManifestOptions,
    manifestOptions?: MeetingAppAdapterManifestOptions,
  ): MeetingAppAdapterManifestMatrix;
  meetingAppAdapterSpec(
    specOrPlatform?: string | MeetingAppAdapterSpecInput,
    specOptions?: MeetingAppAdapterSpecInput,
  ): MeetingAppAdapterSpec;
  meetingAppAdapterSpecMatrix(specOptions?: MeetingAppAdapterSpecMatrixOptions): MeetingAppAdapterSpecMatrix;
  meetingAppAdapterSpecTemplate(input?: MeetingAppAdapterSpecInput): MeetingAppAdapterSpecInput;
  assertMeetingAppAdapterSpec(
    specOrPlatform?: MeetingAppAdapterSpec | string | MeetingAppAdapterSpecInput,
    specOptions?: MeetingAppAdapterSpecInput,
  ): MeetingAppAdapterSpec;
  assertMeetingAppAdapterSpecMatrix(
    matrixOrOptions?: MeetingAppAdapterSpecMatrix | MeetingAppAdapterSpecMatrixOptions,
    specOptions?: MeetingAppAdapterSpecMatrixOptions,
  ): MeetingAppAdapterSpecMatrix;
  meetingAppAdapterRuntimeConfig(
    specOrPlatform?: MeetingAppAdapterSpec | string | MeetingAppAdapterSpecInput,
    configOptions?: MeetingAppAdapterRuntimeConfigOptions,
  ): MeetingAppAdapterRuntimeConfig;
  meetingAppAdapterRuntimeConfigMatrix(configOptions?: MeetingAppAdapterRuntimeConfigMatrixOptions): MeetingAppAdapterRuntimeConfigMatrix;
  assertMeetingAppAdapterRuntimeConfig(
    configOrSpec?: MeetingAppAdapterRuntimeConfig | MeetingAppAdapterSpec | string | MeetingAppAdapterSpecInput,
    configOptions?: MeetingAppAdapterRuntimeConfigOptions,
  ): MeetingAppAdapterRuntimeConfig;
  assertMeetingAppAdapterRuntimeConfigMatrix(
    matrixOrOptions?: MeetingAppAdapterRuntimeConfigMatrix | MeetingAppAdapterRuntimeConfigMatrixOptions,
    configOptions?: MeetingAppAdapterRuntimeConfigMatrixOptions,
  ): MeetingAppAdapterRuntimeConfigMatrix;
  meetingAppAdapterHandoffPackage(
    specOrPlatform?: MeetingAppAdapterSpec | string | MeetingAppAdapterSpecInput,
    packageOptions?: MeetingAppAdapterHandoffPackageOptions,
  ): MeetingAppAdapterHandoffPackage;
  meetingAppAdapterHandoffPackageMatrix(packageOptions?: MeetingAppAdapterHandoffPackageMatrixOptions): MeetingAppAdapterHandoffPackageMatrix;
  assertMeetingAppAdapterHandoffPackage(
    packageOrSpec?: MeetingAppAdapterHandoffPackage | MeetingAppAdapterSpec | string | MeetingAppAdapterSpecInput,
    packageOptions?: MeetingAppAdapterHandoffPackageOptions,
  ): MeetingAppAdapterHandoffPackage;
  assertMeetingAppAdapterHandoffPackageMatrix(
    matrixOrOptions?: MeetingAppAdapterHandoffPackageMatrix | MeetingAppAdapterHandoffPackageMatrixOptions,
    packageOptions?: MeetingAppAdapterHandoffPackageMatrixOptions,
  ): MeetingAppAdapterHandoffPackageMatrix;
  meetingAppAdapterVerificationReport(
    packageOrSpec?: MeetingAppAdapterHandoffPackage | MeetingAppAdapterSpec | string | MeetingAppAdapterSpecInput,
    verificationOptions?: MeetingAppAdapterVerificationOptions,
  ): MeetingAppAdapterVerificationReport;
  meetingAppAdapterVerificationReportMatrix(verificationOptions?: MeetingAppAdapterVerificationMatrixOptions): MeetingAppAdapterVerificationReportMatrix;
  assertMeetingAppAdapterVerificationReport(
    reportOrPackage?: MeetingAppAdapterVerificationReport | MeetingAppAdapterHandoffPackage | MeetingAppAdapterSpec | string | MeetingAppAdapterSpecInput,
    verificationOptions?: MeetingAppAdapterVerificationOptions,
  ): MeetingAppAdapterVerificationReport;
  assertMeetingAppAdapterVerificationReportMatrix(
    matrixOrOptions?: MeetingAppAdapterVerificationReportMatrix | MeetingAppAdapterVerificationMatrixOptions,
    verificationOptions?: MeetingAppAdapterVerificationMatrixOptions,
  ): MeetingAppAdapterVerificationReportMatrix;
  meetingAppTrackPipeline(
    input?: MeetingAppTrackPipelineInput | Record<string, unknown>[],
    trackOptions?: MeetingAppTrackPipelineOptions,
  ): MeetingAppTrackPipeline;
  createMeetingAppTrackPipeline(trackOptions?: MeetingAppTrackPipelineOptions): MeetingAppTrackPipelineAccumulator;
  meetingAppTrackRuntime(runtimeOptions?: MeetingAppTrackRuntimeOptions): MeetingAppTrackRuntime;
  platformTimelineViewPlan(platform: string, viewOptions?: MeetingPlatformTimelineViewOptions): MeetingPlatformTimelineViewPlan;
  platformTimelineViewMatrix(viewOptions?: MeetingPlatformTimelineViewOptions): MeetingPlatformTimelineViewMatrix;
  platformTimelineView(
    platform: string,
    input?: MeetingPlatformTimelineViewInput,
    viewOptions?: MeetingPlatformTimelineViewOptions,
  ): MeetingPlatformTimelineView;
  zoomPlatformTimelineViewport(
    viewport?: Partial<MeetingPlatformTimelineViewport>,
    factor?: number,
    anchorRatio?: number,
    viewOptions?: MeetingPlatformTimelineViewOptions,
  ): MeetingPlatformTimelineViewport;
  platformAnnotationIntakePlan(platform: string, intakeOptions?: MeetingPlatformAnnotationIntakeOptions): MeetingPlatformAnnotationIntakePlan;
  platformAnnotationIntakeMatrix(intakeOptions?: MeetingPlatformAnnotationIntakeOptions): MeetingPlatformAnnotationIntakeMatrix;
  platformAnnotationIntake(
    platform: string,
    input?: MeetingPlatformAnnotationIntakeInput,
    intakeOptions?: MeetingPlatformAnnotationIntakeOptions,
  ): MeetingPlatformAnnotationIntake;
  platformClockSyncPlan(platform: string, clockOptions?: MeetingPlatformClockSyncOptions): MeetingPlatformClockSyncPlan;
  platformClockSyncMatrix(clockOptions?: MeetingPlatformClockSyncOptions): MeetingPlatformClockSyncMatrix;
  platformClockSync(
    platform: string,
    input?: MeetingPlatformClockSyncInput,
    clockOptions?: MeetingPlatformClockSyncOptions,
  ): MeetingPlatformClockSyncReport;
  platformSessionBindingPlan(platform: string, bindingOptions?: MeetingPlatformSessionBindingOptions): MeetingPlatformSessionBindingPlan;
  platformSessionBindingMatrix(bindingOptions?: MeetingPlatformSessionBindingOptions): MeetingPlatformSessionBindingMatrix;
  platformSessionBinding(
    platform: string,
    input?: MeetingPlatformSessionBindingInput,
    bindingOptions?: MeetingPlatformSessionBindingOptions,
  ): MeetingPlatformSessionBinding;
  platformRealtimeAnnotationPlan(platform: string, realtimeOptions?: MeetingPlatformRealtimeAnnotationOptions): MeetingPlatformRealtimeAnnotationPlan;
  platformRealtimeAnnotationMatrix(realtimeOptions?: MeetingPlatformRealtimeAnnotationOptions): MeetingPlatformRealtimeAnnotationMatrix;
  platformRealtimeAnnotation(
    platform: string,
    input?: MeetingPlatformRealtimeAnnotationInput,
    realtimeOptions?: MeetingPlatformRealtimeAnnotationOptions,
  ): MeetingPlatformRealtimeAnnotation;
  platformArtifactHandoffPlan(platform: string, artifactOptions?: MeetingPlatformArtifactHandoffOptions): MeetingPlatformArtifactHandoffPlan;
  platformArtifactHandoffMatrix(artifactOptions?: MeetingPlatformArtifactHandoffOptions): MeetingPlatformArtifactHandoffMatrix;
  platformArtifactHandoff(
    platform: string,
    input?: MeetingPlatformArtifactHandoffInput | Record<string, unknown>[],
    artifactOptions?: MeetingPlatformArtifactHandoffOptions,
  ): MeetingPlatformArtifactHandoff;
  platformRuntimeProfile(platform: string, profileOptions?: MeetingPlatformRuntimeProfileOptions): MeetingPlatformRuntimeProfile;
  platformRuntimeProfileMatrix(profileOptions?: MeetingPlatformRuntimeProfileOptions): MeetingPlatformRuntimeProfileMatrix;
  platformAdaptationPackage(platform: string, packageOptions?: MeetingPlatformAdaptationPackageOptions): MeetingPlatformAdaptationPackage;
  platformAdaptationPackageMatrix(packageOptions?: MeetingPlatformAdaptationPackageOptions): MeetingPlatformAdaptationPackageMatrix;
  platformRuntimeBundle(platform: string, bundleOptions?: MeetingPlatformRuntimeBundleOptions): MeetingPlatformRuntimeBundle;
  platformRuntimeBundleMatrix(bundleOptions?: MeetingPlatformRuntimeBundleOptions): MeetingPlatformRuntimeBundleMatrix;
  platformRuntimeHostConfig(platform: string | MeetingPlatformRuntimeBundle | Record<string, unknown>, hostOptions?: MeetingPlatformRuntimeHostOptions): MeetingPlatformRuntimeHostConfig;
  platformRuntimeHostConfigMatrix(hostOptions?: MeetingPlatformRuntimeHostOptions): MeetingPlatformRuntimeHostConfigMatrix;
  platformRuntimeHostHandoff(
    platformOrConfig: string | MeetingPlatformRuntimeBundle | MeetingPlatformRuntimeHostConfig | Record<string, unknown>,
    handoffOptions?: MeetingPlatformRuntimeHostOptions,
  ): MeetingPlatformRuntimeHostHandoff;
  platformRuntimeHostHandoffMatrix(handoffOptions?: MeetingPlatformRuntimeHostOptions): MeetingPlatformRuntimeHostHandoffMatrix;
  createPlatformRuntimeHostVerificationClient(clientOptions?: Record<string, unknown>): MeetingPlatformRuntimeHostVerificationClient;
  createPlatformRuntimeHostFixtureEnvironment(platform: string, envOptions?: Record<string, unknown>): MeetingPlatformRuntimeHostFixtureEnvironment;
  verifyPlatformRuntimeHost(platform: string, verifyOptions?: Record<string, unknown>): Promise<MeetingPlatformRuntimeHostVerificationReport>;
  verifyPlatformRuntimeHostMatrix(verifyOptions?: Record<string, unknown>): Promise<MeetingPlatformRuntimeHostVerificationMatrix>;
  replayPlatformRuntimeHost(platform: string, input?: Record<string, unknown> | unknown[], replayOptions?: Record<string, unknown>): Promise<MeetingPlatformRuntimeHostReplayReport>;
  replayPlatformRuntimeHostMatrix(replayOptions?: Record<string, unknown>): Promise<MeetingPlatformRuntimeHostReplayMatrix>;
  createPlatformRuntimeHost(
    clientOrRuntime?: MeetingTimelineClient | Record<string, unknown>,
    platformOrConfig?: string | MeetingPlatformRuntimeBundle | MeetingPlatformRuntimeHostConfig | Record<string, unknown>,
    hostOptions?: MeetingPlatformRuntimeHostOptions,
  ): MeetingPlatformRuntimeHost;
  platformRuntimeEventPlan(platform: string, planOptions?: MeetingPlatformRuntimeEventClientOptions): MeetingPlatformRuntimeEventPlan;
  platformRuntimeEventPlanMatrix(planOptions?: MeetingPlatformRuntimeEventClientOptions): MeetingPlatformRuntimeEventPlanMatrix;
  platformRegistryEntry(platform: string, registryOptions?: MeetingPlatformRegistryOptions): MeetingPlatformRegistryEntry;
  platformRegistryManifest(registryOptions?: MeetingPlatformRegistryOptions): MeetingPlatformRegistryManifest;
  platformRegistryAcceptance(
    manifestOrOptions?: MeetingPlatformRegistryManifest | MeetingPlatformRegistryEntry | MeetingPlatformRegistryOptions,
    registryOptions?: MeetingPlatformRegistryOptions,
  ): MeetingPlatformRegistryAcceptanceReport;
  assertPlatformRegistryManifest(
    manifestOrOptions?: MeetingPlatformRegistryManifest | MeetingPlatformRegistryEntry | MeetingPlatformRegistryOptions,
    registryOptions?: MeetingPlatformRegistryOptions,
  ): MeetingPlatformRegistryAcceptanceReport;
  platformConnector(platform: string, connectorOptions?: MeetingPlatformConnectorOptions): MeetingPlatformConnector;
  platformConnectorMatrix(connectorOptions?: MeetingPlatformConnectorOptions): MeetingPlatformConnectorMatrix;
  platformConnectorHub(connectorOptions?: MeetingPlatformConnectorOptions): MeetingPlatformConnectorHub;
  resolvePlatformConnector(input?: unknown, connectorOptions?: MeetingPlatformConnectorOptions): MeetingPlatformConnectorResolution;
  platformConnectorAcceptance(
    connectorOrOptions?: MeetingPlatformConnector | string | MeetingPlatformConnectorOptions,
    connectorOptions?: MeetingPlatformConnectorOptions,
  ): MeetingPlatformConnectorAcceptanceReport;
  assertPlatformConnector(
    connectorOrOptions?: MeetingPlatformConnector | string | MeetingPlatformConnectorOptions,
    connectorOptions?: MeetingPlatformConnectorOptions,
  ): MeetingPlatformConnectorAcceptanceReport;
  createPlatformConnectorRuntime(
    platformOrConnector: string | MeetingPlatformConnector,
    connectorOptions?: MeetingPlatformConnectorOptions,
  ): MeetingPlatformConnectorRuntime;
  createPlatformConnectorHub(connectorOptions?: MeetingPlatformConnectorOptions): MeetingPlatformConnectorHubRuntime;
  createPlatformConnectorBrowserRuntime(connectorOptions?: MeetingPlatformConnectorOptions): MeetingPlatformConnectorBrowserRuntime;
  createPlatformConnectorContentScriptBridge(connectorOptions?: MeetingPlatformConnectorOptions): MeetingPlatformConnectorContentScriptBridge;
  installPlatformConnectorContentScriptBridge(connectorOptions?: MeetingPlatformConnectorOptions): MeetingPlatformConnectorContentScriptBridge;
  platformAdapterContract(platform: string, contractOptions?: MeetingPlatformAdapterContractOptions): MeetingPlatformAdapterContract;
  platformAdapterContractMatrix(contractOptions?: MeetingPlatformAdapterContractOptions): MeetingPlatformAdapterContractMatrix;
  platformAdapterContractAcceptance(
    contractOrPlatform: string | MeetingPlatformAdapterContract,
    acceptanceOptions?: MeetingPlatformAdapterContractOptions,
  ): MeetingPlatformAdapterContractAcceptanceReport;
  platformAdapterContractAcceptanceMatrix(acceptanceOptions?: MeetingPlatformAdapterContractOptions): MeetingPlatformAdapterContractAcceptanceMatrix;
  assertPlatformAdapterContract(
    contractOrPlatform: string | MeetingPlatformAdapterContract,
    acceptanceOptions?: MeetingPlatformAdapterContractOptions,
  ): MeetingPlatformAdapterContractAcceptanceReport;
  platformAdapterSamplePlan(platform: string, sampleOptions?: MeetingPlatformAdapterSampleOptions): MeetingPlatformAdapterSamplePlan;
  runPlatformAdapterSample(platform: string, sampleOptions?: MeetingPlatformAdapterSampleOptions): Promise<MeetingPlatformAdapterSample>;
  runPlatformAdapterSampleMatrix(sampleOptions?: MeetingPlatformAdapterSampleOptions): Promise<MeetingPlatformAdapterSampleMatrix>;
  assertPlatformAdapterSample(
    sampleOrPlatform: string | MeetingPlatformAdapterSample,
    sampleOptions?: MeetingPlatformAdapterSampleOptions,
  ): Promise<MeetingPlatformAdapterSample>;
  assertPlatformAdapterSampleMatrix(sampleOptions?: MeetingPlatformAdapterSampleOptions): Promise<MeetingPlatformAdapterSampleMatrix>;
  platformRealEvidenceIntakePlan(platform: string, intakeOptions?: MeetingPlatformRealEvidenceIntakeOptions): MeetingPlatformRealEvidenceIntakePlan;
  platformRealEvidenceIntake(
    platform: string,
    input?: Record<string, unknown>,
    intakeOptions?: MeetingPlatformRealEvidenceIntakeOptions,
  ): MeetingPlatformRealEvidenceIntakeReport;
  platformRealEvidenceIntakeMatrix(
    input?: Record<string, unknown>,
    intakeOptions?: MeetingPlatformRealEvidenceIntakeOptions,
  ): MeetingPlatformRealEvidenceIntakeMatrix;
  runPlatformRealEvidenceIntake(
    platform: string,
    input?: Record<string, unknown>,
    intakeOptions?: MeetingPlatformRealEvidenceIntakeOptions,
  ): Promise<MeetingPlatformRealEvidenceIntakeReport>;
  runPlatformRealEvidenceIntakeMatrix(
    input?: Record<string, unknown>,
    intakeOptions?: MeetingPlatformRealEvidenceIntakeOptions,
  ): Promise<MeetingPlatformRealEvidenceIntakeMatrix>;
  assertPlatformRealEvidenceIntake(
    platform: string,
    input?: Record<string, unknown>,
    intakeOptions?: MeetingPlatformRealEvidenceIntakeOptions,
  ): MeetingPlatformRealEvidenceIntakeReport;
  assertPlatformRealEvidenceIntakeMatrix(
    input?: Record<string, unknown>,
    intakeOptions?: MeetingPlatformRealEvidenceIntakeOptions,
  ): MeetingPlatformRealEvidenceIntakeMatrix;
  platformFieldCapturePlan(platform: string, captureOptions?: MeetingPlatformFieldCaptureOptions): MeetingPlatformFieldCapturePlan;
  platformFieldCaptureMatrix(captureOptions?: MeetingPlatformFieldCaptureOptions): MeetingPlatformFieldCaptureMatrix;
  platformFieldCaptureManifest(platform: string, manifestOptions?: MeetingPlatformFieldCaptureOptions): MeetingPlatformFieldCaptureManifest;
  platformFieldCaptureManifestMatrix(manifestOptions?: MeetingPlatformFieldCaptureOptions): MeetingPlatformFieldCaptureManifestMatrix;
  platformFieldCollectorConfig(platform: string, collectorOptions?: MeetingPlatformFieldCaptureOptions): MeetingPlatformFieldCollectorConfig;
  platformFieldCollectorConfigMatrix(collectorOptions?: MeetingPlatformFieldCaptureOptions): MeetingPlatformFieldCollectorConfigMatrix;
  platformFieldEvidenceBundle(platform: string, input?: MeetingPlatformEvidencePackage | MeetingPlatformEvidencePackageOptions, bundleOptions?: MeetingPlatformFieldCaptureOptions): MeetingPlatformFieldEvidenceBundle;
  platformFieldEvidenceMatrix(matrixOptions?: MeetingPlatformFieldCaptureOptions): MeetingPlatformFieldEvidenceMatrix;
  platformFieldIntakePlan(platform: string, intakeOptions?: MeetingPlatformFieldIntakeOptions): MeetingPlatformFieldIntakePlan;
  platformFieldIntakeMatrix(intakeOptions?: MeetingPlatformFieldIntakeOptions): MeetingPlatformFieldIntakeMatrix;
  platformHandoffReadiness(platform: string, input?: MeetingPlatformHandoffReadinessOptions, readinessOptions?: MeetingPlatformHandoffReadinessOptions): MeetingPlatformHandoffReadiness;
  platformHandoffReadinessMatrix(input?: MeetingPlatformHandoffReadinessOptions, readinessOptions?: MeetingPlatformHandoffReadinessOptions): MeetingPlatformHandoffReadinessMatrix;
  runPlatformHandoffReadiness(platform: string, input?: MeetingPlatformHandoffReadinessOptions, readinessOptions?: MeetingPlatformHandoffReadinessOptions): Promise<MeetingPlatformHandoffReadiness>;
  runPlatformHandoffReadinessMatrix(input?: MeetingPlatformHandoffReadinessOptions, readinessOptions?: MeetingPlatformHandoffReadinessOptions): Promise<MeetingPlatformHandoffReadinessMatrix>;
  assertPlatformHandoffReadiness(platform: string, input?: MeetingPlatformHandoffReadinessOptions, readinessOptions?: MeetingPlatformHandoffReadinessOptions): MeetingPlatformHandoffReadiness;
  assertPlatformHandoffReadinessMatrix(input?: MeetingPlatformHandoffReadinessOptions, readinessOptions?: MeetingPlatformHandoffReadinessOptions): MeetingPlatformHandoffReadinessMatrix;
  meetingAppSnapshotRecorder(recorderOptions?: MeetingAppSnapshotRecordOptions): MeetingAppSnapshotRecorder;
  meetingAppGateInputFromRecords(records?: MeetingAppSnapshotRecord[] | MeetingAppSnapshotRecordSet, recordOptions?: MeetingAppSnapshotRecordOptions): Record<string, unknown>;
  meetingAppIntegrationProfile(platform: string, profileOptions?: MeetingAppIntegrationProfileOptions): MeetingAppIntegrationProfile;
  allMeetingAppIntegrationProfiles(profileOptions?: MeetingAppIntegrationProfileOptions): Partial<Record<string, MeetingAppIntegrationProfile>>;
  meetingAppIntegrationMatrix(profileOptions?: MeetingAppIntegrationProfileOptions): MeetingAppIntegrationMatrix;
  meetingAppDeploymentManifest(platform: string, manifestOptions?: MeetingAppIntegrationProfileOptions): MeetingAppDeploymentManifest;
  allMeetingAppDeploymentManifests(manifestOptions?: MeetingAppIntegrationProfileOptions): Partial<Record<string, MeetingAppDeploymentManifest>>;
  meetingAppDeploymentManifestAcceptance(manifestOrPlatform?: MeetingAppDeploymentManifest | MeetingAppRuntimeAdapterConfig | string | MeetingAppIntegrationProfileOptions, acceptanceOptions?: MeetingAppIntegrationProfileOptions): MeetingAppDeploymentManifestAcceptanceReport;
  allMeetingAppDeploymentManifestAcceptanceReports(acceptanceOptions?: MeetingAppIntegrationProfileOptions): Partial<Record<string, MeetingAppDeploymentManifestAcceptanceReport>>;
  meetingAppDeploymentManifestAcceptanceSummary(acceptanceOptions?: MeetingAppIntegrationProfileOptions): MeetingAppDeploymentManifestAcceptanceSummary;
  assertMeetingAppDeploymentManifest(manifestOrPlatform?: MeetingAppDeploymentManifest | MeetingAppRuntimeAdapterConfig | string | MeetingAppIntegrationProfileOptions, acceptanceOptions?: MeetingAppIntegrationProfileOptions): MeetingAppDeploymentManifestAcceptanceReport;
  meetingAppLiveEvidencePackage(input?: Record<string, unknown> | unknown[], evidenceOptions?: MeetingAppIntegrationProfileOptions): MeetingAppLiveEvidencePackage;
  meetingAppLiveEvidencePackageSummary(input?: MeetingAppLiveEvidencePackage | Record<string, unknown> | unknown[], evidenceOptions?: MeetingAppIntegrationProfileOptions): MeetingAppLiveEvidencePackageSummary;
  meetingAppDomAdaptationDiagnosis(platformOrInput?: string | MeetingAppIntegrationProfileOptions | Record<string, unknown> | unknown[], diagnosisOptions?: MeetingAppIntegrationProfileOptions): MeetingAppDomAdaptationDiagnosis;
  allMeetingAppDomAdaptationDiagnoses(diagnosisOptions?: MeetingAppIntegrationProfileOptions): Partial<Record<string, MeetingAppDomAdaptationDiagnosis>>;
  meetingAppDomAdaptationDiagnosisMatrix(diagnosisOptions?: MeetingAppIntegrationProfileOptions): MeetingAppDomAdaptationDiagnosisMatrix;
  meetingAppRuntimeAdapterConfig(platform: string, configOptions?: MeetingAppIntegrationProfileOptions): MeetingAppRuntimeAdapterConfig;
  meetingAppRuntimeAdapterProfile(input?: string | Record<string, unknown>, profileOptions?: MeetingAppIntegrationProfileOptions): MeetingAppRuntimeAdapterProfileResolution;
  meetingAppRuntimeAdapterProfileMatrix(profileOptions?: MeetingAppIntegrationProfileOptions): MeetingAppRuntimeAdapterProfileMatrix;
  meetingAppRuntimeObserverPlan(platformOrInput?: string | Record<string, unknown>, planOptions?: MeetingAppIntegrationProfileOptions): MeetingAppRuntimeObserverPlan;
  meetingAppRuntimeObserverPlanMatrix(planOptions?: MeetingAppIntegrationProfileOptions): MeetingAppRuntimeObserverPlanMatrix;
  meetingAppObserverSchedulerConfig(planOrPlatform?: string | Record<string, unknown> | MeetingAppRuntimeObserverPlan, schedulerOptions?: MeetingAppObserverSchedulerOptions): MeetingAppObserverSchedulerConfig;
  meetingAppObserverSchedulerConfigMatrix(schedulerOptions?: MeetingAppObserverSchedulerOptions): MeetingAppObserverSchedulerConfigMatrix;
  selectMeetingAppRuntimeAdapter(input?: string | Record<string, unknown>, selectionOptions?: MeetingAppIntegrationProfileOptions): MeetingAppRuntimeAdapterSelection;
  meetingAppRuntimeAdapterHandoff(selectionOrInput?: string | Record<string, unknown> | MeetingAppRuntimeAdapterSelection, handoffOptions?: MeetingAppIntegrationProfileOptions): MeetingAppRuntimeAdapterHandoff;
  meetingAppRuntimeAdapterHandoffMatrix(handoffOptions?: MeetingAppIntegrationProfileOptions): MeetingAppRuntimeAdapterHandoffMatrix;
  meetingAppRuntimeAdapterHandoffAcceptance(handoffOrInput?: string | Record<string, unknown> | MeetingAppRuntimeAdapterSelection | MeetingAppRuntimeAdapterHandoff, acceptanceOptions?: MeetingAppIntegrationProfileOptions): MeetingAppRuntimeAdapterHandoffAcceptanceReport;
  assertMeetingAppRuntimeAdapterHandoff(handoffOrInput?: string | Record<string, unknown> | MeetingAppRuntimeAdapterSelection | MeetingAppRuntimeAdapterHandoff, acceptanceOptions?: MeetingAppIntegrationProfileOptions): MeetingAppRuntimeAdapterHandoffAcceptanceReport;
  meetingAppRuntimeAdapterHandoffMatrixAcceptance(matrixOrOptions?: Record<string, unknown> | MeetingAppRuntimeAdapterHandoff | MeetingAppRuntimeAdapterHandoffMatrix, acceptanceOptions?: MeetingAppIntegrationProfileOptions): MeetingAppRuntimeAdapterHandoffMatrixAcceptanceReport;
  assertMeetingAppRuntimeAdapterHandoffMatrix(matrixOrOptions?: Record<string, unknown> | MeetingAppRuntimeAdapterHandoff | MeetingAppRuntimeAdapterHandoffMatrix, acceptanceOptions?: MeetingAppIntegrationProfileOptions): MeetingAppRuntimeAdapterHandoffMatrixAcceptanceReport;
  meetingAppRuntimeAdapterHostPackage(hostPackageOptions?: MeetingAppIntegrationProfileOptions): MeetingAppRuntimeAdapterHostPackage;
  meetingAppLiveSnapshotCapturePlan(platform: string, planOptions?: MeetingAppIntegrationProfileOptions): MeetingAppLiveSnapshotCapturePlan;
  allMeetingAppLiveSnapshotCapturePlans(planOptions?: MeetingAppIntegrationProfileOptions): Partial<Record<string, MeetingAppLiveSnapshotCapturePlan>>;
  allMeetingAppRuntimeAdapterConfigs(configOptions?: MeetingAppIntegrationProfileOptions): Partial<Record<string, MeetingAppRuntimeAdapterConfig>>;
  meetingAppRuntimeAdapterAcceptance(configOrPlatform?: MeetingAppRuntimeAdapterConfig | string | MeetingAppIntegrationProfileOptions, acceptanceOptions?: MeetingAppIntegrationProfileOptions): MeetingAppRuntimeAdapterAcceptanceReport;
  allMeetingAppRuntimeAdapterAcceptanceReports(acceptanceOptions?: MeetingAppIntegrationProfileOptions): Partial<Record<string, MeetingAppRuntimeAdapterAcceptanceReport>>;
  assertMeetingAppRuntimeAdapterConfig(configOrPlatform?: MeetingAppRuntimeAdapterConfig | string | MeetingAppIntegrationProfileOptions, acceptanceOptions?: MeetingAppIntegrationProfileOptions): MeetingAppRuntimeAdapterAcceptanceReport;
  meetingAppRuntimeAdapterValidation(configOrPlatform?: MeetingAppRuntimeAdapterConfig | string | MeetingAppIntegrationProfileOptions, validationReportOptions?: MeetingAppIntegrationProfileOptions): MeetingAppRuntimeAdapterValidationReport;
  allMeetingAppRuntimeAdapterValidationReports(validationReportOptions?: MeetingAppIntegrationProfileOptions): Partial<Record<string, MeetingAppRuntimeAdapterValidationReport>>;
  assertMeetingAppRuntimeAdapterValidation(configOrPlatform?: MeetingAppRuntimeAdapterConfig | string | MeetingAppIntegrationProfileOptions, validationReportOptions?: MeetingAppIntegrationProfileOptions): MeetingAppRuntimeAdapterValidationReport;
  meetingAppExtensionInstallPlan(extensionOptions?: MeetingAppExtensionOptions): MeetingAppExtensionInstallPlan;
  meetingAppContentScriptManifest(extensionOptions?: MeetingAppExtensionOptions): Record<string, unknown>;
  meetingAppExtensionMatches(extensionOptions?: MeetingAppExtensionOptions): MeetingAppExtensionMatchPatterns;
  meetingAppExtensionScaffold(extensionOptions?: MeetingAppExtensionOptions): MeetingAppExtensionScaffold;
  meetingAppExtensionAcceptance(extensionOptions?: MeetingAppExtensionOptions): MeetingAppExtensionScaffoldAcceptanceReport;
  assertMeetingAppExtensionScaffold(extensionOptions?: MeetingAppExtensionOptions): MeetingAppExtensionScaffoldAcceptanceReport;
  meetingAppExtensionAttachedMessage(input?: MeetingAppExtensionMessageOptions | string, messageOptions?: MeetingAppExtensionMessageOptions): Record<string, unknown>;
  meetingAppExtensionStatusMessage(input?: MeetingAppExtensionMessageOptions, messageOptions?: MeetingAppExtensionMessageOptions): Record<string, unknown>;
  meetingAppExtensionObserveCandidatesMessage(input?: MeetingAppExtensionMessageOptions, messageOptions?: MeetingAppExtensionMessageOptions): Record<string, unknown>;
  meetingAppExtensionClientCallMessage(
    methodOrInput: MeetingAppExtensionClientCallMethod | string | MeetingAppExtensionClientCallMessageInput,
    input?: Record<string, unknown>,
    messageOptions?: MeetingAppExtensionMessageOptions,
  ): Record<string, unknown>;
  meetingAppExtensionTimelineEndpoint(method: MeetingAppExtensionClientCallMethod | string): string;
  normalizeMeetingAppExtensionMessageType(messageType: MeetingAppExtensionMessageType | string): MeetingAppExtensionMessageType;
  assertMeetingAppLaunchGate(platform: string, gateOptions?: MeetingAppLaunchGateOptions): MeetingAppLaunchGate;
  assertAllMeetingAppLaunchGates(gateOptions?: MeetingAppLaunchGateOptions): MeetingAppLaunchGateSummary;
  diagnose(platform: string, payload?: unknown, diagnosticOptions?: Record<string, unknown>): PlatformEventDiagnosticResult;
  platform(platform: string, platformOptions?: MeetingPlatformTimelineKitOptions): Record<string, unknown>;
  allPlatforms(platformOptions?: MeetingPlatformTimelineKitOptions): Record<string, unknown>[];
  capability(platform: string, platformOptions?: MeetingPlatformTimelineKitOptions): Record<string, unknown>;
  routeTable(routeOptions?: MeetingPlatformTimelineKitOptions): Record<string, unknown>[];
  routerStatus(routeOptions?: MeetingPlatformTimelineKitOptions): Record<string, unknown>;
  routerSetup(routeOptions?: MeetingPlatformTimelineKitOptions): Record<string, unknown>;
  acceptance(platform: string, reportOptions?: MeetingPlatformTimelineKitOptions): Record<string, unknown>;
  fixtureAcceptance(platform: string, reportOptions?: MeetingPlatformTimelineKitOptions): Record<string, unknown>;
  fixtureInput(fixtureOptions?: MeetingPlatformTimelineKitOptions): Record<string, unknown>;
  meetingAppFixture(platform: string, fixtureOptions?: MeetingAppFixtureOptions): Record<string, unknown>;
  allMeetingAppFixtures(fixtureOptions?: MeetingAppFixtureOptions): Record<MeetingAppFixturePlatform, Record<string, unknown>>;
  diagnoseMeetingAppFixture(platform: string, diagnosticOptions?: MeetingAppFixtureOptions): MeetingAppFixtureDiagnosis;
  diagnoseMeetingAppFixtureLifecycle(platform: string, diagnosticOptions?: MeetingAppFixtureOptions): MeetingAppFixtureLifecycleDiagnosis;
  meetingAppFixtureAcceptance(reportOptions?: MeetingAppFixtureOptions): Record<string, unknown>;
  onboarding(platform: string, onboardingOptions?: MeetingPlatformTimelineKitOptions): Record<string, unknown>;
  report(reportOptions?: MeetingPlatformTimelineKitOptions): Record<string, unknown>;
  getState(): Record<string, unknown>;
  reset(nextState?: Record<string, unknown>): Record<string, unknown>;
}

export function buildMeetingPlatformKitReport(options?: MeetingPlatformTimelineKitOptions): Record<string, unknown>;

export function createMeetingPlatformTimelineKit(
  clientOrOptions: MeetingTimelineClient | MeetingTimelineClientOptions,
  options?: MeetingPlatformTimelineKitOptions,
): MeetingPlatformTimelineKit;
