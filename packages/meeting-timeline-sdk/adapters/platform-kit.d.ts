import type { MeetingTimelineClient, MeetingTimelineClientOptions, TimelineMarkInput, TranscriptImportInput } from '../index.mjs';
import type { MeetingTimelineBridge, MeetingTimelineBridgeOptions } from './timeline-bridge.mjs';
import type {
  PlatformWebhookRouterOptions,
  PlatformWebhookRouterRequest,
} from './platform-webhook-router.mjs';
import type { MeetingPlatformFetchHandler, PlatformWebhookHttpOptions } from './platform-http.mjs';
import type { MeetingPlatformNodeHandler, PlatformWebhookNodeOptions } from './platform-node.mjs';
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
  MeetingPlatformRuntimeProfile,
  MeetingPlatformRuntimeProfileMatrix,
  MeetingPlatformRuntimeProfileOptions,
} from './platform-runtime-profile.mjs';
import type {
  MeetingPlatformAdapterContract,
  MeetingPlatformAdapterContractAcceptanceMatrix,
  MeetingPlatformAdapterContractAcceptanceReport,
  MeetingPlatformAdapterContractMatrix,
  MeetingPlatformAdapterContractOptions,
} from './platform-adapter-contract.mjs';
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
import type { MeetingAppLaunchGate, MeetingAppLaunchGateOptions, MeetingAppLaunchGateSummary } from './meeting-app-gate.mjs';
import type { MeetingAppSnapshotRecorder, MeetingAppSnapshotRecord, MeetingAppSnapshotRecordOptions, MeetingAppSnapshotRecordSet } from './meeting-app-snapshot-recorder.mjs';
import type {
  MeetingAppDeploymentManifest,
  MeetingAppDeploymentManifestAcceptanceReport,
  MeetingAppDeploymentManifestAcceptanceSummary,
  MeetingAppDomAdaptationDiagnosis,
  MeetingAppIntegrationMatrix,
  MeetingAppIntegrationProfile,
  MeetingAppIntegrationProfileOptions,
  MeetingAppLiveEvidencePackage,
  MeetingAppLiveEvidencePackageSummary,
  MeetingAppLiveSnapshotCapturePlan,
  MeetingAppRuntimeAdapterAcceptanceReport,
  MeetingAppRuntimeAdapterConfig,
  MeetingAppRuntimeAdapterValidationReport,
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
import type { PlatformEventDiagnosticResult, PlatformEventIngestInput, PlatformEventIngestOptions, ReconciledPlatformEventIngestResult } from './platform-ingest.mjs';
import type {
  MeetingAppFixtureDiagnosis,
  MeetingAppFixtureLifecycleDiagnosis,
  MeetingAppFixtureOptions,
  MeetingAppFixturePlatform,
} from './meeting-app-fixtures.mjs';

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
  platformRuntimeProfile(platform: string, profileOptions?: MeetingPlatformRuntimeProfileOptions): MeetingPlatformRuntimeProfile;
  platformRuntimeProfileMatrix(profileOptions?: MeetingPlatformRuntimeProfileOptions): MeetingPlatformRuntimeProfileMatrix;
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
  platformFieldCapturePlan(platform: string, captureOptions?: MeetingPlatformFieldCaptureOptions): MeetingPlatformFieldCapturePlan;
  platformFieldCaptureMatrix(captureOptions?: MeetingPlatformFieldCaptureOptions): MeetingPlatformFieldCaptureMatrix;
  platformFieldCaptureManifest(platform: string, manifestOptions?: MeetingPlatformFieldCaptureOptions): MeetingPlatformFieldCaptureManifest;
  platformFieldCaptureManifestMatrix(manifestOptions?: MeetingPlatformFieldCaptureOptions): MeetingPlatformFieldCaptureManifestMatrix;
  platformFieldCollectorConfig(platform: string, collectorOptions?: MeetingPlatformFieldCaptureOptions): MeetingPlatformFieldCollectorConfig;
  platformFieldCollectorConfigMatrix(collectorOptions?: MeetingPlatformFieldCaptureOptions): MeetingPlatformFieldCollectorConfigMatrix;
  platformFieldEvidenceBundle(platform: string, input?: MeetingPlatformEvidencePackage | MeetingPlatformEvidencePackageOptions, bundleOptions?: MeetingPlatformFieldCaptureOptions): MeetingPlatformFieldEvidenceBundle;
  platformFieldEvidenceMatrix(matrixOptions?: MeetingPlatformFieldCaptureOptions): MeetingPlatformFieldEvidenceMatrix;
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
  meetingAppRuntimeAdapterConfig(platform: string, configOptions?: MeetingAppIntegrationProfileOptions): MeetingAppRuntimeAdapterConfig;
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
