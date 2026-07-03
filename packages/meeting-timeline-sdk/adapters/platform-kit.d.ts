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
import type { MeetingAppLaunchGate, MeetingAppLaunchGateOptions, MeetingAppLaunchGateSummary } from './meeting-app-gate.mjs';
import type { MeetingAppSnapshotRecorder, MeetingAppSnapshotRecord, MeetingAppSnapshotRecordOptions, MeetingAppSnapshotRecordSet } from './meeting-app-snapshot-recorder.mjs';
import type {
  MeetingAppIntegrationMatrix,
  MeetingAppIntegrationProfile,
  MeetingAppIntegrationProfileOptions,
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
  meetingAppSnapshotRecorder(recorderOptions?: MeetingAppSnapshotRecordOptions): MeetingAppSnapshotRecorder;
  meetingAppGateInputFromRecords(records?: MeetingAppSnapshotRecord[] | MeetingAppSnapshotRecordSet, recordOptions?: MeetingAppSnapshotRecordOptions): Record<string, unknown>;
  meetingAppIntegrationProfile(platform: string, profileOptions?: MeetingAppIntegrationProfileOptions): MeetingAppIntegrationProfile;
  allMeetingAppIntegrationProfiles(profileOptions?: MeetingAppIntegrationProfileOptions): Partial<Record<string, MeetingAppIntegrationProfile>>;
  meetingAppIntegrationMatrix(profileOptions?: MeetingAppIntegrationProfileOptions): MeetingAppIntegrationMatrix;
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
