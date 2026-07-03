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
  buildAllMeetingAppDeploymentManifestAcceptanceReports,
  buildAllMeetingAppDeploymentManifests,
  buildAllMeetingAppIntegrationProfiles,
  buildAllMeetingAppLiveSnapshotCapturePlans,
  buildAllMeetingAppRuntimeAdapterAcceptanceReports,
  buildAllMeetingAppRuntimeAdapterConfigs,
  buildAllMeetingAppRuntimeAdapterValidationReports,
  buildMeetingAppDeploymentManifestAcceptanceReport,
  buildMeetingAppDeploymentManifestAcceptanceSummary,
  buildMeetingAppDeploymentManifest,
  buildMeetingAppIntegrationMatrix,
  buildMeetingAppIntegrationProfile,
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
    meeting_app_runtime_adapter_configs: buildAllMeetingAppRuntimeAdapterConfigs(options),
    meeting_app_runtime_adapter_acceptance: buildAllMeetingAppRuntimeAdapterAcceptanceReports(options),
    meeting_app_live_snapshot_capture_plans: buildAllMeetingAppLiveSnapshotCapturePlans(options),
    meeting_app_runtime_adapter_validation: buildAllMeetingAppRuntimeAdapterValidationReports(options),
    meeting_app_fixture_acceptance: meetingAppFixtureAcceptance,
    meeting_app_launch_gate: meetingAppLaunchGateSummary,
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
