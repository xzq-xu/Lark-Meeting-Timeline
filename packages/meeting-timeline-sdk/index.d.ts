export const SDK_VERSION: string;
export const DEFAULT_MEETING_APP_TIMELINE_SDK_PLATFORMS: readonly string[];

export interface MeetingStartInput {
  platform?: string;
  meeting_id?: string;
  meetingId?: string;
  external_meeting_id?: string;
  externalMeetingId?: string;
  meeting_no?: string;
  meetingNo?: string;
  meeting_url?: string;
  meetingUrl?: string;
  url?: string;
  title?: string;
  topic?: string;
  name?: string;
  timezone?: string;
  start_time_ms?: number | string | Date;
  startTimeMs?: number | string | Date;
  start_time?: number | string | Date;
  startTime?: number | string | Date;
  detector_source?: string;
  detectorSource?: string;
  force?: boolean;
  [key: string]: unknown;
}

export interface MeetingEndInput {
  meeting_id?: string;
  meetingId?: string;
  end_time_ms?: number | string | Date;
  endTimeMs?: number | string | Date;
  end_time?: number | string | Date;
  endTime?: number | string | Date;
  time_ms?: number;
  timeMs?: number;
  detector_source?: string;
  detectorSource?: string;
  [key: string]: unknown;
}

export interface TranscriptSegmentInput {
  id?: string;
  segment_id?: string;
  segmentId?: string;
  sentence_id?: string;
  sentenceId?: string;
  start_ms?: number | string;
  startMs?: number | string;
  start_time_ms?: number | string | Date;
  startTimeMs?: number | string | Date;
  start_time?: number | string | Date;
  startTime?: number | string | Date;
  start?: number | string | Date;
  end_ms?: number | string;
  endMs?: number | string;
  end_time_ms?: number | string | Date;
  endTimeMs?: number | string | Date;
  end_time?: number | string | Date;
  endTime?: number | string | Date;
  end?: number | string | Date;
  speaker_id?: string;
  speakerId?: string;
  speaker_name?: string;
  speakerName?: string;
  participant_id?: string;
  participantId?: string;
  participant_name?: string;
  participantName?: string;
  user_id?: string;
  userId?: string;
  user_name?: string;
  userName?: string;
  text?: string;
  content?: string;
  sentence?: string;
  transcript?: string;
  language?: string;
  language_code?: string;
  languageCode?: string;
  source?: string;
  raw?: unknown;
  [key: string]: unknown;
}

export interface TranscriptImportInput {
  meeting?: MeetingStartInput;
  meetingSession?: MeetingStartInput;
  session?: MeetingStartInput;
  meeting_id?: string;
  meetingId?: string;
  platform?: string;
  source?: string;
  transcript?: TranscriptSegmentInput[];
  segments?: TranscriptSegmentInput[];
  entries?: TranscriptSegmentInput[];
  items?: TranscriptSegmentInput[];
  artifact?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface TimelineMarkInput {
  id?: string;
  annotation_id?: string;
  annotationId?: string;
  source?: string;
  device_id?: string;
  deviceId?: string;
  captured_at_ms?: number | string | Date;
  capturedAtMs?: number | string | Date;
  captured_at?: number | string | Date;
  capturedAt?: number | string | Date;
  time_ms?: number;
  timeMs?: number;
  kind?: string;
  type?: string;
  label?: string;
  text?: string;
  reading?: string;
  text_candidates?: string[];
  textCandidates?: string[];
  intent?: string;
  mark?: Record<string, unknown>;
  target?: Record<string, unknown>;
  target_region?: Record<string, unknown>;
  strokes?: unknown[];
  stroke_points?: unknown[];
  device?: Record<string, unknown>;
  payload?: Record<string, unknown>;
  realtime?: boolean;
  live?: boolean;
  meeting_session?: MeetingStartInput;
  meetingSession?: MeetingStartInput;
  start_meeting_session?: boolean;
  startMeetingSession?: boolean;
  force_meeting_session?: boolean;
  forceMeetingSession?: boolean;
  [key: string]: unknown;
}

export interface MeetingTimelineClientOptions {
  baseUrl: string;
  fetch?: typeof fetch;
  fetchImpl?: typeof fetch;
  headers?: Record<string, string>;
  authToken?: string;
  token?: string;
  source?: string;
  deviceId?: string;
  device_id?: string;
  detectorSource?: string;
  detector_source?: string;
  requireCapturedAt?: boolean;
  timeoutMs?: number;
}

export class MeetingTimelineSdkError extends Error {
  details: Record<string, unknown>;
}

export class MeetingTimelineApiError extends MeetingTimelineSdkError {
  status: number | null;
  body: unknown;
}

export function compactObject<T>(value: T): T;
export function normalizeAbsoluteMs(value: number | string | Date, fieldName?: string): number;
export function buildMeetingStartPayload(input?: MeetingStartInput, defaults?: Record<string, unknown>): Record<string, unknown>;
export function buildMeetingEndPayload(input?: MeetingEndInput, defaults?: Record<string, unknown>): Record<string, unknown>;
export function buildTranscriptImportPayload(input?: TranscriptImportInput, defaults?: Record<string, unknown>): Record<string, unknown>;
export function buildTimelineMark(input?: TimelineMarkInput, defaults?: Record<string, unknown>): Record<string, unknown>;

export class MeetingTimelineClient {
  constructor(options: MeetingTimelineClientOptions);
  request(path: string, options?: Record<string, unknown>): Promise<unknown>;
  getState(): Promise<unknown>;
  getIngestInfo(): Promise<unknown>;
  getMeetingSessionStatus(): Promise<unknown>;
  startMeeting(input?: MeetingStartInput): Promise<unknown>;
  endMeeting(input?: MeetingEndInput): Promise<unknown>;
  insertMark(input?: TimelineMarkInput, options?: Record<string, unknown>): Promise<unknown>;
  addAnnotation(input?: TimelineMarkInput, options?: Record<string, unknown>): Promise<unknown>;
  insertAnnotation(input?: TimelineMarkInput, options?: Record<string, unknown>): Promise<unknown>;
  insertMarks(inputs?: TimelineMarkInput[] | { annotations?: TimelineMarkInput[]; items?: TimelineMarkInput[] }, options?: Record<string, unknown>): Promise<unknown>;
  addAnnotations(inputs?: TimelineMarkInput[] | { annotations?: TimelineMarkInput[]; items?: TimelineMarkInput[] }, options?: Record<string, unknown>): Promise<unknown>;
  getAnnotationStatus(id: string): Promise<unknown>;
  importTranscript(input?: TranscriptImportInput, options?: Record<string, unknown>): Promise<unknown>;
  importMeetingTranscript(input?: TranscriptImportInput, options?: Record<string, unknown>): Promise<unknown>;
  subscribeState(options?: Record<string, unknown>): { stream: unknown; close(): void };
}

export function createMeetingTimelineClient(options: MeetingTimelineClientOptions): MeetingTimelineClient;
export const createMeetingTimelineSdk: typeof createMeetingTimelineClient;

export interface MeetingAppTimelineSdkOptions extends Partial<MeetingTimelineClientOptions> {
  client?: MeetingTimelineClient;
  timelineClient?: MeetingTimelineClient;
  timeline_client?: MeetingTimelineClient;
  platforms?: Iterable<string> | string[];
  platform_keys?: Iterable<string> | string[];
  platformKeys?: Iterable<string> | string[];
  runtimeEventClientOptions?: Record<string, unknown>;
  runtime_event_client_options?: Record<string, unknown>;
  runtimeEventFetch?: typeof fetch;
  runtime_event_fetch?: typeof fetch;
  runtimeEventFetchImpl?: typeof fetch;
  runtime_event_fetch_impl?: typeof fetch;
  mode?: string;
  delivery?: string;
  route?: string;
  remote?: boolean;
  runtimeEvent?: boolean;
  runtime_event?: boolean;
  [key: string]: unknown;
}

export interface MeetingAppTimelineConnectorPackage {
  type: 'meeting_app_timeline_connector_package';
  schema: 'meeting_app_timeline_connector_package';
  schema_version: 1;
  id: string;
  base_url?: string;
  accepted: boolean;
  platforms: string[];
  surfaces: string[];
  platform_count: number;
  surface_count: number;
  handoff_count: number;
  ready_count: number;
  host_package: import('./adapters/meeting-app-profile.mjs').MeetingAppRuntimeAdapterHostPackage;
  handoff_matrix: import('./adapters/meeting-app-profile.mjs').MeetingAppRuntimeAdapterHandoffMatrix;
  handoff_acceptance: import('./adapters/meeting-app-profile.mjs').MeetingAppRuntimeAdapterHandoffMatrixAcceptanceReport;
  observer_plan_by_surface: Record<string, import('./adapters/meeting-app-profile.mjs').MeetingAppRuntimeObserverPlanMatrix>;
  scheduler_config_by_surface: Record<string, import('./adapters/meeting-app-observer-scheduler.mjs').MeetingAppObserverSchedulerConfigMatrix>;
  extension?: {
    scaffold?: import('./adapters/meeting-app-extension.mjs').MeetingAppExtensionScaffold;
    acceptance?: import('./adapters/meeting-app-extension.mjs').MeetingAppExtensionScaffoldAcceptanceReport;
    install_plan?: import('./adapters/meeting-app-extension.mjs').MeetingAppExtensionInstallPlan;
    manifest?: Record<string, unknown>;
    bundle?: Record<string, unknown>;
    file_count?: number;
  };
  runtime_events: {
    endpoint: string;
    plan_matrix: import('./adapters/platform-runtime-event.mjs').MeetingPlatformRuntimeEventPlanMatrix;
    action_count: number;
  };
  entrypoints: Array<Record<string, unknown>>;
  contracts: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingAppTimelineSdk {
  type: 'meeting_app_timeline_sdk';
  schema: 'meeting_app_timeline_sdk';
  schema_version: 1;
  platforms: string[];
  client: MeetingTimelineClient;
  kit: import('./adapters/platform-kit.mjs').MeetingPlatformTimelineKit;
  runtime: import('./adapters/platform-integration-runtime.mjs').MeetingPlatformIntegrationRuntime;
  integrationRuntime: import('./adapters/platform-integration-runtime.mjs').MeetingPlatformIntegrationRuntime;
  integration_runtime: import('./adapters/platform-integration-runtime.mjs').MeetingPlatformIntegrationRuntime;
  runtimeEvents: import('./adapters/platform-runtime-event.mjs').MeetingPlatformRuntimeEventClient;
  runtime_events: import('./adapters/platform-runtime-event.mjs').MeetingPlatformRuntimeEventClient;
  detect(input?: Record<string, unknown>, options?: Record<string, unknown>): Record<string, unknown>;
  resolve(input?: Record<string, unknown>, options?: Record<string, unknown>): Record<string, unknown>;
  resolveCandidates(input?: Record<string, unknown>, options?: Record<string, unknown>): Record<string, unknown>;
  package(
    platformOrOptions?: string | Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/meeting-app-adapter-integration-package.mjs').MeetingAppAdapterIntegrationPackage;
  integrationPackage: MeetingAppTimelineSdk['package'];
  packageMatrix(options?: Record<string, unknown>): import('./adapters/meeting-app-adapter-integration-package.mjs').MeetingAppAdapterIntegrationPackageMatrix;
  integrationPackageMatrix: MeetingAppTimelineSdk['packageMatrix'];
  integrationProfile(
    platformOrOptions?: string | Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/meeting-app-profile.mjs').MeetingAppIntegrationProfile;
  integrationMatrix(options?: Record<string, unknown>): import('./adapters/meeting-app-profile.mjs').MeetingAppIntegrationMatrix;
  runtimeAdapterProfile(
    input?: string | Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/meeting-app-profile.mjs').MeetingAppRuntimeAdapterProfileResolution;
  adapterProfile: MeetingAppTimelineSdk['runtimeAdapterProfile'];
  runtimeAdapterProfileMatrix(options?: Record<string, unknown>): import('./adapters/meeting-app-profile.mjs').MeetingAppRuntimeAdapterProfileMatrix;
  adapterProfileMatrix: MeetingAppTimelineSdk['runtimeAdapterProfileMatrix'];
  observerPlan(
    platformOrInput?: string | Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/meeting-app-profile.mjs').MeetingAppRuntimeObserverPlan;
  runtimeObserverPlan: MeetingAppTimelineSdk['observerPlan'];
  observerPlanMatrix(options?: Record<string, unknown>): import('./adapters/meeting-app-profile.mjs').MeetingAppRuntimeObserverPlanMatrix;
  runtimeObserverPlanMatrix: MeetingAppTimelineSdk['observerPlanMatrix'];
  selectAdapter(
    input?: string | Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/meeting-app-profile.mjs').MeetingAppRuntimeAdapterSelection;
  selectRuntimeAdapter: MeetingAppTimelineSdk['selectAdapter'];
  runtimeAdapterHandoff(
    selectionOrInput?: string | Record<string, unknown> | import('./adapters/meeting-app-profile.mjs').MeetingAppRuntimeAdapterSelection,
    options?: Record<string, unknown>,
  ): import('./adapters/meeting-app-profile.mjs').MeetingAppRuntimeAdapterHandoff;
  adapterHandoff: MeetingAppTimelineSdk['runtimeAdapterHandoff'];
  handoff: MeetingAppTimelineSdk['runtimeAdapterHandoff'];
  runtimeAdapterHandoffMatrix(options?: Record<string, unknown>): import('./adapters/meeting-app-profile.mjs').MeetingAppRuntimeAdapterHandoffMatrix;
  adapterHandoffMatrix: MeetingAppTimelineSdk['runtimeAdapterHandoffMatrix'];
  handoffMatrix: MeetingAppTimelineSdk['runtimeAdapterHandoffMatrix'];
  runtimeAdapterHandoffAcceptance(
    handoffOrInput?: string | Record<string, unknown> | import('./adapters/meeting-app-profile.mjs').MeetingAppRuntimeAdapterSelection | import('./adapters/meeting-app-profile.mjs').MeetingAppRuntimeAdapterHandoff,
    options?: Record<string, unknown>,
  ): import('./adapters/meeting-app-profile.mjs').MeetingAppRuntimeAdapterHandoffAcceptanceReport;
  handoffAcceptance: MeetingAppTimelineSdk['runtimeAdapterHandoffAcceptance'];
  runtimeAdapterHandoffMatrixAcceptance(
    matrixOrOptions?: Record<string, unknown> | import('./adapters/meeting-app-profile.mjs').MeetingAppRuntimeAdapterHandoff | import('./adapters/meeting-app-profile.mjs').MeetingAppRuntimeAdapterHandoffMatrix,
    options?: Record<string, unknown>,
  ): import('./adapters/meeting-app-profile.mjs').MeetingAppRuntimeAdapterHandoffMatrixAcceptanceReport;
  handoffMatrixAcceptance: MeetingAppTimelineSdk['runtimeAdapterHandoffMatrixAcceptance'];
  runtimeAdapterHostPackage(options?: Record<string, unknown>): import('./adapters/meeting-app-profile.mjs').MeetingAppRuntimeAdapterHostPackage;
  hostPackage: MeetingAppTimelineSdk['runtimeAdapterHostPackage'];
  connectorPackage(options?: Record<string, unknown>): MeetingAppTimelineConnectorPackage;
  runtimeConnectorPackage: MeetingAppTimelineSdk['connectorPackage'];
  platformAdaptationPackage(
    platformOrOptions?: string | Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/platform-adaptation-package.mjs').MeetingPlatformAdaptationPackage;
  adaptationPackage: MeetingAppTimelineSdk['platformAdaptationPackage'];
  platformAdaptationPackageMatrix(options?: Record<string, unknown>): import('./adapters/platform-adaptation-package.mjs').MeetingPlatformAdaptationPackageMatrix;
  adaptationPackageMatrix: MeetingAppTimelineSdk['platformAdaptationPackageMatrix'];
  platformConsumerHandoff(options?: Record<string, unknown>): import('./adapters/platform-consumer-handoff.mjs').MeetingPlatformConsumerHandoff;
  consumerHandoff: MeetingAppTimelineSdk['platformConsumerHandoff'];
  assertPlatformConsumerHandoff(options?: Record<string, unknown>): import('./adapters/platform-consumer-handoff.mjs').MeetingPlatformConsumerHandoff;
  assertConsumerHandoff: MeetingAppTimelineSdk['assertPlatformConsumerHandoff'];
  platformImplementationHandoff(
    platformOrOptions?: string | Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/platform-implementation-handoff.mjs').MeetingPlatformImplementationHandoff;
  implementationHandoff: MeetingAppTimelineSdk['platformImplementationHandoff'];
  platformImplementationHandoffMatrix(options?: Record<string, unknown>): import('./adapters/platform-implementation-handoff.mjs').MeetingPlatformImplementationHandoffMatrix;
  implementationHandoffMatrix: MeetingAppTimelineSdk['platformImplementationHandoffMatrix'];
  platformAdapterAuthoringPlan(
    platformOrOptions?: string | Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/platform-adapter-authoring.mjs').MeetingPlatformAdapterAuthoringPlan;
  adapterAuthoringPlan: MeetingAppTimelineSdk['platformAdapterAuthoringPlan'];
  platformAdapterAuthoringMatrix(options?: Record<string, unknown>): import('./adapters/platform-adapter-authoring.mjs').MeetingPlatformAdapterAuthoringMatrix;
  adapterAuthoringMatrix: MeetingAppTimelineSdk['platformAdapterAuthoringMatrix'];
  platformAdapterPortfolioItem(
    platformOrOptions?: string | Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/platform-adapter-portfolio.mjs').MeetingPlatformAdapterPortfolioItem;
  adapterPortfolioItem: MeetingAppTimelineSdk['platformAdapterPortfolioItem'];
  platformAdapterPortfolio(options?: Record<string, unknown>): import('./adapters/platform-adapter-portfolio.mjs').MeetingPlatformAdapterPortfolio;
  adapterPortfolio: MeetingAppTimelineSdk['platformAdapterPortfolio'];
  platformAdapterAcceptanceChecklist(
    platformOrOptions?: string | Record<string, unknown>,
    input?: Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/platform-adapter-acceptance-checklist.mjs').MeetingPlatformAdapterAcceptanceChecklist;
  adapterAcceptanceChecklist: MeetingAppTimelineSdk['platformAdapterAcceptanceChecklist'];
  platformAdapterAcceptanceChecklistMatrix(
    input?: Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/platform-adapter-acceptance-checklist.mjs').MeetingPlatformAdapterAcceptanceChecklistMatrix;
  adapterAcceptanceChecklistMatrix: MeetingAppTimelineSdk['platformAdapterAcceptanceChecklistMatrix'];
  platformAdapterExportPackage(
    platformOrOptions?: string | Record<string, unknown>,
    input?: Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/platform-adapter-export-package.mjs').MeetingPlatformAdapterExportPackage;
  adapterExportPackage: MeetingAppTimelineSdk['platformAdapterExportPackage'];
  platformAdapterExportPackageMatrix(
    input?: Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/platform-adapter-export-package.mjs').MeetingPlatformAdapterExportPackageMatrix;
  adapterExportPackageMatrix: MeetingAppTimelineSdk['platformAdapterExportPackageMatrix'];
  platformAdapterImportPlan(
    exportPackage?: import('./adapters/platform-adapter-export-package.mjs').MeetingPlatformAdapterExportPackage | Record<string, unknown>,
    input?: Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/platform-adapter-import-plan.mjs').MeetingPlatformAdapterImportPlan;
  adapterImportPlan: MeetingAppTimelineSdk['platformAdapterImportPlan'];
  assertPlatformAdapterImportPlan(
    planOrPackage?: import('./adapters/platform-adapter-import-plan.mjs').MeetingPlatformAdapterImportPlan | import('./adapters/platform-adapter-export-package.mjs').MeetingPlatformAdapterExportPackage | Record<string, unknown>,
    input?: Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/platform-adapter-import-plan.mjs').MeetingPlatformAdapterImportPlan;
  assertAdapterImportPlan: MeetingAppTimelineSdk['assertPlatformAdapterImportPlan'];
  platformAdapterImportPlanMatrix(
    packagesOrInput?: import('./adapters/platform-adapter-export-package.mjs').MeetingPlatformAdapterExportPackage[] | Record<string, unknown>,
    input?: Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/platform-adapter-import-plan.mjs').MeetingPlatformAdapterImportPlanMatrix;
  adapterImportPlanMatrix: MeetingAppTimelineSdk['platformAdapterImportPlanMatrix'];
  platformAdapterInstallManifest(
    plansOrInput?: import('./adapters/platform-adapter-import-plan.mjs').MeetingPlatformAdapterImportPlan[] | import('./adapters/platform-adapter-export-package.mjs').MeetingPlatformAdapterExportPackage[] | Record<string, unknown>,
    input?: Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/platform-adapter-install-manifest.mjs').MeetingPlatformAdapterInstallManifest;
  adapterInstallManifest: MeetingAppTimelineSdk['platformAdapterInstallManifest'];
  assertPlatformAdapterInstallManifest(
    manifestOrInput?: import('./adapters/platform-adapter-install-manifest.mjs').MeetingPlatformAdapterInstallManifest | import('./adapters/platform-adapter-import-plan.mjs').MeetingPlatformAdapterImportPlan[] | import('./adapters/platform-adapter-export-package.mjs').MeetingPlatformAdapterExportPackage[] | Record<string, unknown>,
    input?: Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/platform-adapter-install-manifest.mjs').MeetingPlatformAdapterInstallManifest;
  assertAdapterInstallManifest: MeetingAppTimelineSdk['assertPlatformAdapterInstallManifest'];
  platformAdapterLaunchPlan(
    manifestOrInput?: import('./adapters/platform-adapter-install-manifest.mjs').MeetingPlatformAdapterInstallManifest | Record<string, unknown>,
    input?: Record<string, unknown> | string,
    options?: Record<string, unknown>,
  ): import('./adapters/platform-adapter-launch-plan.mjs').MeetingPlatformAdapterLaunchPlan;
  adapterLaunchPlan: MeetingAppTimelineSdk['platformAdapterLaunchPlan'];
  assertPlatformAdapterLaunchPlan(
    planOrInput?: import('./adapters/platform-adapter-launch-plan.mjs').MeetingPlatformAdapterLaunchPlan | import('./adapters/platform-adapter-install-manifest.mjs').MeetingPlatformAdapterInstallManifest | Record<string, unknown>,
    input?: Record<string, unknown> | string,
    options?: Record<string, unknown>,
  ): import('./adapters/platform-adapter-launch-plan.mjs').MeetingPlatformAdapterLaunchPlan;
  assertAdapterLaunchPlan: MeetingAppTimelineSdk['assertPlatformAdapterLaunchPlan'];
  platformAdapterSession(
    launchPlanOrInput?: import('./adapters/platform-adapter-launch-plan.mjs').MeetingPlatformAdapterLaunchPlan | Record<string, unknown>,
    clientOrOptions?: import('./adapters/platform-adapter-session.mjs').MeetingPlatformAdapterSessionClient | import('./adapters/platform-adapter-session.mjs').MeetingPlatformAdapterSessionOptions,
    options?: import('./adapters/platform-adapter-session.mjs').MeetingPlatformAdapterSessionOptions,
  ): import('./adapters/platform-adapter-session.mjs').MeetingPlatformAdapterSession;
  adapterSession: MeetingAppTimelineSdk['platformAdapterSession'];
  platformAdapterSessionHandoff(
    launchPlanOrInput?: import('./adapters/platform-adapter-launch-plan.mjs').MeetingPlatformAdapterLaunchPlan | Record<string, unknown>,
    options?: import('./adapters/platform-adapter-session.mjs').MeetingPlatformAdapterSessionOptions,
  ): import('./adapters/platform-adapter-session.mjs').MeetingPlatformAdapterSessionHandoff;
  adapterSessionHandoff: MeetingAppTimelineSdk['platformAdapterSessionHandoff'];
  platformAdapterRunner(
    manifestOrInput?: import('./adapters/platform-adapter-install-manifest.mjs').MeetingPlatformAdapterInstallManifest | import('./adapters/platform-adapter-launch-plan.mjs').MeetingPlatformAdapterLaunchPlan | Record<string, unknown>,
    clientOrOptions?: import('./adapters/platform-adapter-session.mjs').MeetingPlatformAdapterSessionClient | import('./adapters/platform-adapter-runner.mjs').MeetingPlatformAdapterRunnerOptions,
    options?: import('./adapters/platform-adapter-runner.mjs').MeetingPlatformAdapterRunnerOptions,
  ): import('./adapters/platform-adapter-runner.mjs').MeetingPlatformAdapterRunner;
  adapterRunner: MeetingAppTimelineSdk['platformAdapterRunner'];
  openPlatformAdapterSession(
    manifestOrInput?: import('./adapters/platform-adapter-install-manifest.mjs').MeetingPlatformAdapterInstallManifest | import('./adapters/platform-adapter-launch-plan.mjs').MeetingPlatformAdapterLaunchPlan | Record<string, unknown>,
    launchInput?: Record<string, unknown> | string,
    options?: import('./adapters/platform-adapter-runner.mjs').MeetingPlatformAdapterRunnerOptions,
  ): Promise<import('./adapters/platform-adapter-runner.mjs').MeetingPlatformAdapterOpenSessionEvent>;
  openAdapterSession: MeetingAppTimelineSdk['openPlatformAdapterSession'];
  platformAdapterRunnerHandoff(
    manifestOrInput?: import('./adapters/platform-adapter-install-manifest.mjs').MeetingPlatformAdapterInstallManifest | import('./adapters/platform-adapter-launch-plan.mjs').MeetingPlatformAdapterLaunchPlan | Record<string, unknown>,
    options?: import('./adapters/platform-adapter-runner.mjs').MeetingPlatformAdapterRunnerOptions,
  ): import('./adapters/platform-adapter-runner.mjs').MeetingPlatformAdapterRunnerHandoff;
  adapterRunnerHandoff: MeetingAppTimelineSdk['platformAdapterRunnerHandoff'];
  platformRuntimeBundle(
    platformOrOptions?: string | Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/platform-runtime-bundle.mjs').MeetingPlatformRuntimeBundle;
  runtimeBundle: MeetingAppTimelineSdk['platformRuntimeBundle'];
  platformRuntimeBundleMatrix(options?: Record<string, unknown>): import('./adapters/platform-runtime-bundle.mjs').MeetingPlatformRuntimeBundleMatrix;
  runtimeBundleMatrix: MeetingAppTimelineSdk['platformRuntimeBundleMatrix'];
  platformAdapterRoute(
    platformOrOptions?: string | Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/platform-adapter-route.mjs').MeetingPlatformAdapterRoute;
  adapterRoute: MeetingAppTimelineSdk['platformAdapterRoute'];
  platformAdapterRouteMatrix(options?: Record<string, unknown>): import('./adapters/platform-adapter-route.mjs').MeetingPlatformAdapterRouteMatrix;
  adapterRouteMatrix: MeetingAppTimelineSdk['platformAdapterRouteMatrix'];
  platformAdaptationStrategy(
    platformOrOptions?: string | Record<string, unknown>,
    options?: Record<string, unknown>,
  ): import('./adapters/platform-strategy.mjs').MeetingPlatformAdaptationStrategy;
  adaptationStrategy: MeetingAppTimelineSdk['platformAdaptationStrategy'];
  platformAdaptationStrategyMatrix(options?: Record<string, unknown>): import('./adapters/platform-strategy.mjs').MeetingPlatformAdaptationStrategyMatrix;
  adaptationStrategyMatrix: MeetingAppTimelineSdk['platformAdaptationStrategyMatrix'];
  platformConnectorHub(options?: Record<string, unknown>): import('./adapters/meeting-platform-connector.mjs').MeetingPlatformConnectorHub;
  connectorHub: MeetingAppTimelineSdk['platformConnectorHub'];
  manifest(options?: Record<string, unknown>): Record<string, unknown>;
  readiness(options?: Record<string, unknown>): Record<string, unknown>;
  handoffReadiness(options?: Record<string, unknown>): Record<string, unknown>;
  observePlatformCandidates(input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown> | unknown;
  observeMeetingApp(
    platformOrInput: string | Record<string, unknown>,
    inputOrOptions?: Record<string, unknown>,
    options?: Record<string, unknown>,
  ): Promise<unknown> | unknown;
  observeApp: MeetingAppTimelineSdk['observeMeetingApp'];
  insertAnnotation(
    platformOrInput: string | TimelineMarkInput,
    inputOrOptions?: TimelineMarkInput | Record<string, unknown>,
    options?: Record<string, unknown>,
  ): Promise<unknown> | unknown;
  insertMark: MeetingAppTimelineSdk['insertAnnotation'];
  ingestProvider(
    platformOrInput: string | Record<string, unknown>,
    payloadOrOptions?: Record<string, unknown>,
    optionsOrPayload?: Record<string, unknown>,
    options?: Record<string, unknown>,
  ): Promise<unknown> | unknown;
  speakerTrack(
    platformOrInput: string | Record<string, unknown>,
    inputOrOptions?: Record<string, unknown>,
    options?: Record<string, unknown>,
  ): Promise<unknown> | unknown;
  participantTrack: MeetingAppTimelineSdk['speakerTrack'];
  timelineView: MeetingAppTimelineSdk['speakerTrack'];
  sendRuntimeEvent(input?: Record<string, unknown>, options?: Record<string, unknown>): Promise<unknown>;
  handleRuntimeEvent(input?: Record<string, unknown>, payload?: unknown, options?: Record<string, unknown>): Promise<unknown> | unknown;
}

export function createMeetingAppTimelineSdk(options: MeetingAppTimelineSdkOptions): MeetingAppTimelineSdk;

export * from './adapters/platform-kit.mjs';
export * from './adapters/meeting-app-adapter-integration-package.mjs';
export * from './adapters/meeting-app-connector-package.mjs';
export * from './adapters/meeting-platform-connector.mjs';
export * from './adapters/platform-integration-runtime.mjs';
export * from './adapters/platform-runtime-event.mjs';
export * from './adapters/platform-adaptation-package.mjs';
export * from './adapters/platform-consumer-handoff.mjs';
export * from './adapters/platform-implementation-handoff.mjs';
export * from './adapters/platform-adapter-authoring.mjs';
export * from './adapters/platform-adapter-portfolio.mjs';
export * from './adapters/platform-adapter-acceptance-checklist.mjs';
export * from './adapters/platform-adapter-export-package.mjs';
export * from './adapters/platform-adapter-import-plan.mjs';
export * from './adapters/platform-adapter-install-manifest.mjs';
export * from './adapters/platform-adapter-launch-plan.mjs';
export * from './adapters/platform-adapter-session.mjs';
export * from './adapters/platform-adapter-runner.mjs';
export * from './adapters/platform-runtime-bundle.mjs';
export * from './adapters/platform-adapter-route.mjs';
export * from './adapters/platform-strategy.mjs';
