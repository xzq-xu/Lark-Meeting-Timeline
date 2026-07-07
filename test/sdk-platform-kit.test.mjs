import assert from 'node:assert/strict';

import {
  buildMeetingPlatformKitReport,
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import {
  buildPlatformFixtureEnv,
  buildPlatformFixtureEvent,
} from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';
import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import { platformCapabilityContract } from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';

const baseUrl = 'https://timeline.example.com';
const basePath = '/hooks/meeting-events';
const env = buildPlatformFixtureEnv({ baseUrl });
const calls = [];
const client = {
  async startMeeting(input) {
    calls.push({ method: 'startMeeting', input });
    return { ok: true, input };
  },
  async endMeeting(input) {
    calls.push({ method: 'endMeeting', input });
    return { ok: true, input };
  },
  async insertMark(input) {
    calls.push({ method: 'insertMark', input });
    return { ok: true, input };
  },
  async insertMarks(input) {
    calls.push({ method: 'insertMarks', input });
    return { ok: true, input };
  },
  async importTranscript(input) {
    calls.push({ method: 'importTranscript', input });
    return { ok: true, input };
  },
};

function fakeBrowserWindow(url = 'https://meet.google.com/abc-defg-hij') {
  const document = {
    nodeType: 9,
    title: 'Kit connector browser runtime',
    hidden: false,
    location: { href: url },
    body: { nodeType: 1 },
    documentElement: { nodeType: 1 },
    querySelectorAll() {
      return [];
    },
  };
  return {
    document,
    location: { href: url, origin: new URL(url).origin },
    navigator: { userAgent: 'Chrome fixture' },
    addEventListener() {},
    removeEventListener() {},
  };
}

assert.equal(
  platformCapabilityContract('google-meet', { baseUrl }).sdk_modules.platform_kit,
  '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit',
);

const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl,
  basePath,
  env,
  verify: false,
  reconcile: true,
  applyOptions: {
    participantAsAnnotation: true,
  },
});

assert.equal(kit.client, client);
assert.equal(kit.platforms.includes('google_meet'), true);
assert.equal(kit.routeTable().find((item) => item.platform === 'google_meet').path, `${basePath}/google-meet`);
assert.equal(kit.routerStatus().platforms.find((item) => item.platform === 'zoom').ready, true);

const googleOverview = kit.platform('google-meet');
assert.equal(googleOverview.platform, 'google_meet');
assert.equal(googleOverview.webhook_route, `${basePath}/google-meet`);
assert.equal(googleOverview.fixture_samples.some((item) => item.label === 'google_meet:meeting_start'), true);
assert.equal(googleOverview.meeting_app_fixture.coverage.meeting_started, true);
assert.equal(googleOverview.meeting_app_fixture.coverage.speaker_started, true);
assert.equal(googleOverview.integration_plan.recommended_mode, 'hybrid_local_observer_first');

const googleRegistryEntry = kit.platformRegistryEntry('google-meet');
assert.equal(googleRegistryEntry.platform, 'google_meet');
assert.equal(googleRegistryEntry.annotations.timestamp_field, 'captured_at_ms');
assert.equal(googleRegistryEntry.runtime.browser_matches.includes('https://meet.google.com/*'), true);
const registryManifest = kit.platformRegistryManifest({ platforms: ['google-meet', 'zoom'] });
assert.equal(registryManifest.platform_count, 2);
assert.equal(registryManifest.provider_required_for_realtime_count, 0);
const googleConnector = kit.platformConnector('google-meet');
assert.equal(googleConnector.schema, 'meeting_platform_connector');
assert.equal(googleConnector.platform, 'google_meet');
assert.equal(googleConnector.runtime_events.endpoint, `${baseUrl}/api/meeting-platform/runtime-events`);
assert.equal(googleConnector.runtime_events.supported_actions.includes('insert_annotation'), true);
assert.equal(googleConnector.readiness.realtime_annotation_ready, true);
assert.equal(kit.platformConnectorAcceptance(googleConnector).accepted, true);
assert.equal(kit.assertPlatformConnector(googleConnector).accepted, true);
const connectorMatrix = kit.platformConnectorMatrix({ platforms: ['google-meet', 'teams', 'zoom'] });
assert.equal(connectorMatrix.platform_count, 3);
assert.equal(connectorMatrix.accepted_count, 3);
assert.equal(connectorMatrix.realtime_ready_count, 3);
assert.equal(connectorMatrix.rows.find((row) => row.platform === 'microsoft_teams').browser_observer_enabled, true);
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_connector_matrix.platform_count, 1);
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_connector_hub.accepted, true);
const connectorHub = kit.platformConnectorHub({ platforms: ['google-meet', 'teams', 'zoom'] });
assert.equal(connectorHub.schema, 'meeting_platform_connector_hub');
assert.equal(connectorHub.platform_count, 3);
assert.equal(connectorHub.accepted, true);
assert.equal(kit.resolvePlatformConnector({
  tabs: [{ active: true, url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample' }],
}).platform, 'microsoft_teams');
const connectorRuntimeCalls = [];
const googleConnectorRuntime = kit.createPlatformConnectorRuntime(googleConnector, {
  fetch: async (url, init) => {
    const body = JSON.parse(init.body);
    connectorRuntimeCalls.push({ url, body });
    return new Response(JSON.stringify({ ok: true, accepted: body }), {
      headers: { 'content-type': 'application/json' },
    });
  },
  now: () => 1_782_614_400_000,
});
assert.equal(googleConnectorRuntime.platform, 'google_meet');
assert.equal(googleConnectorRuntime.supports('insert_annotation'), true);
assert.equal(googleConnectorRuntime.normalizeProviderEvent({
  id: 'kit-google-event-001',
  type: 'google.workspace.meet.conference.v2.started',
  data: { conferenceRecord: { name: 'conferenceRecords/kit-google-record-001' } },
}).at(0).type, 'meeting_started');
await googleConnectorRuntime.insertAnnotation({
  id: 'kit-mark-001',
  label: 'why?',
  captured_at_ms: 1_782_614_401_000,
});
assert.equal(connectorRuntimeCalls.at(-1).url, `${baseUrl}/api/meeting-platform/runtime-events`);
assert.equal(connectorRuntimeCalls.at(-1).body.action, 'insert_annotation');
assert.equal(connectorRuntimeCalls.at(-1).body.platform, 'google_meet');
const connectorHubRuntime = kit.createPlatformConnectorHub({
  fetch: async (url, init) => {
    const body = JSON.parse(init.body);
    connectorRuntimeCalls.push({ url, body });
    return new Response(JSON.stringify({ ok: true, accepted: body }), {
      headers: { 'content-type': 'application/json' },
    });
  },
  now: () => 1_782_614_400_000,
});
assert.equal(connectorHubRuntime.connectorFor({ url: 'https://zoom.us/j/987654321' }).platform, 'zoom');
await connectorHubRuntime.insertAnnotation({
  url: 'https://meet.google.com/abc-defg-hij',
}, {
  id: 'kit-hub-mark-001',
  label: 'follow-up',
  captured_at_ms: 1_782_614_401_000,
});
assert.equal(connectorRuntimeCalls.at(-1).body.action, 'insert_annotation');
assert.equal(connectorRuntimeCalls.at(-1).body.platform, 'google_meet');
const kitConnectorWindow = fakeBrowserWindow();
const kitConnectorBrowserRuntime = kit.createPlatformConnectorBrowserRuntime({
  fetch: async (url, init) => {
    const body = JSON.parse(init.body);
    connectorRuntimeCalls.push({ url, body });
    return new Response(JSON.stringify({ ok: true, accepted: body }), {
      headers: { 'content-type': 'application/json' },
    });
  },
  window: kitConnectorWindow,
  now: () => 1_782_614_400_000,
});
assert.equal(kitConnectorBrowserRuntime.schema, 'meeting_platform_connector_browser_runtime');
assert.equal(kitConnectorBrowserRuntime.resolvePlatform().platform, 'google_meet');
await kitConnectorBrowserRuntime.insertAnnotation({
  id: 'kit-browser-mark-001',
  label: 'browser follow-up',
  captured_at_ms: 1_782_614_402_000,
});
assert.equal(connectorRuntimeCalls.at(-1).body.action, 'insert_annotation');
assert.equal(connectorRuntimeCalls.at(-1).body.platform, 'google_meet');
assert.equal(connectorRuntimeCalls.at(-1).body.annotation.id, 'kit-browser-mark-001');
const kitConnectorBridge = kit.createPlatformConnectorContentScriptBridge({
  fetch: async (url, init) => {
    const body = JSON.parse(init.body);
    connectorRuntimeCalls.push({ url, body });
    return new Response(JSON.stringify({ ok: true, accepted: body }), {
      headers: { 'content-type': 'application/json' },
    });
  },
  window: kitConnectorWindow,
  now: () => 1_782_614_400_000,
});
assert.equal(kitConnectorBridge.schema, 'meeting_platform_connector_content_script_bridge');
const kitBridgeResponse = await kitConnectorBridge.dispatchMessage({
  type: 'meeting_timeline.insert_mark',
  payload: {
    mark: {
      id: 'kit-bridge-mark-001',
      label: 'bridge follow-up',
      captured_at_ms: 1_782_614_403_000,
    },
  },
});
assert.equal(kitBridgeResponse.handled, true);
assert.equal(connectorRuntimeCalls.at(-1).body.platform, 'google_meet');
assert.equal(connectorRuntimeCalls.at(-1).body.annotation.id, 'kit-bridge-mark-001');
assert.equal(kit.platformConsumerHandoff({ platforms: ['google-meet'] }).schema, 'meeting_platform_consumer_handoff');
assert.equal(kit.assertPlatformConsumerHandoff({ platforms: ['google-meet'] }).accepted, true);
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_consumer_handoff.consumer_ready_count, 1);

const fixtureAcceptance = kit.fixtureAcceptance('google-meet', {
  requiredCoverage: ['meeting_start', 'meeting_end', 'participant_track', 'artifact_ready', 'subscription_lifecycle'],
});
assert.equal(fixtureAcceptance.accepted, true);
assert.equal(fixtureAcceptance.coverage.subscription_lifecycle, true);

const diagnostic = kit.diagnose('zoom', buildPlatformFixtureEvent('zoom', 'meeting_start'));
assert.equal(diagnostic.signal_types.includes('meeting_started'), true);

const webexAppFixture = kit.meetingAppFixture('webex');
assert.equal(webexAppFixture.platform, 'webex');
assert.equal(webexAppFixture.meeting_id, 'meet-sdk-fixture');

const zoomAppDiagnosis = kit.diagnoseMeetingAppFixture('zoom');
assert.deepEqual(zoomAppDiagnosis.signal_types, ['meeting_started', 'speaker_started']);
assert.equal(zoomAppDiagnosis.coverage.active_speaker, true);

const zoomAppLifecycle = kit.diagnoseMeetingAppFixtureLifecycle('zoom');
assert.deepEqual(zoomAppLifecycle.signal_types, ['meeting_started', 'speaker_started', 'meeting_ended']);
assert.equal(zoomAppLifecycle.coverage.meeting_ended, true);

const appFixtures = kit.allMeetingAppFixtures();
assert.equal(Object.keys(appFixtures).length, 5);
assert.equal(appFixtures.lark.platform, 'lark');

const appFixtureAcceptance = kit.meetingAppFixtureAcceptance();
assert.equal(appFixtureAcceptance.accepted, true);
assert.equal(appFixtureAcceptance.accepted_count, 5);
assert.equal(appFixtureAcceptance.coverage_by_platform.google_meet.meeting_started, true);
assert.equal(appFixtureAcceptance.coverage_by_platform.google_meet.meeting_ended, true);
const appFit = kit.meetingAppAdapterFit({
  tabs: [{
    active: true,
    audible: true,
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Google Meet',
    page: {
      buttons: [{ ariaLabel: 'Leave call' }],
      tiles: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
    },
  }],
}, {
  platform: 'google-meet',
});
assert.equal(appFit.accepted, true);
assert.equal(appFit.ready_for_realtime_axis, true);
assert.equal(appFit.ready_for_speaker_track, true);
assert.equal(appFit.recommended_surface, 'browser_extension_or_webview');
const appFitMatrix = kit.meetingAppAdapterFitMatrix({
  platforms: ['google-meet', 'zoom'],
  inputs: {
    google_meet: {
      url: 'https://meet.google.com/abc-defg-hij',
      title: 'Google Meet',
      page: {
        controls: [{ label: 'Leave call' }],
        participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace speaking' }],
      },
    },
    zoom: {
      app: { name: 'Zoom Workplace' },
      window: { title: 'Zoom Meeting', controls: [{ label: 'Leave Meeting' }] },
      meeting_id: 'zoom-local',
      tiles: [{ id: 'mira', ariaLabel: 'Mira Patel is speaking' }],
    },
  },
});
assert.equal(appFitMatrix.platform_count, 2);
assert.equal(appFitMatrix.realtime_axis_ready_count, 2);
const appCapability = kit.meetingAppAdapterCapability('google-meet', {
  tabs: [{
    active: true,
    audible: true,
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Google Meet',
    page: {
      controls: [{ label: 'Leave call' }],
      participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace speaking' }],
    },
  }],
});
assert.equal(appCapability.schema, 'meeting_app_adapter_capability_report');
assert.equal(appCapability.platform, 'google_meet');
assert.equal(appCapability.pilot_ready, true);
assert.equal(appCapability.recommended_mode, 'hybrid_local_observer_first');
const appCapabilityMatrix = kit.meetingAppAdapterCapabilityMatrix({
  platforms: ['google-meet', 'zoom'],
  inputs: {
    google_meet: {
      url: 'https://meet.google.com/abc-defg-hij',
      title: 'Google Meet',
      page: {
        controls: [{ label: 'Leave call' }],
        participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace speaking' }],
      },
    },
    zoom: {
      app: { name: 'Zoom Workplace' },
      window: { title: 'Zoom Meeting', controls: [{ label: 'Leave Meeting' }] },
      meeting_id: 'zoom-local',
      tiles: [{ id: 'mira', ariaLabel: 'Mira Patel is speaking' }],
    },
  },
});
assert.equal(appCapabilityMatrix.platform_count, 2);
assert.equal(appCapabilityMatrix.pilot_ready_count, 2);
assert.equal(kit.report({ platforms: ['google-meet'] }).meeting_app_adapter_capability_matrix.platform_count, 1);
const appExecutionPlan = kit.meetingAppAdapterExecutionPlan(appCapability);
assert.equal(appExecutionPlan.schema, 'meeting_app_adapter_execution_plan');
assert.equal(appExecutionPlan.realtime_ready, true);
assert.equal(appExecutionPlan.steps.find((step) => step.id === 'configure_provider_reconcile').status, 'ready');
const appExecutionPlanMatrix = kit.meetingAppAdapterExecutionPlanMatrix({
  platforms: ['google-meet', 'zoom'],
  inputs: {
    google_meet: {
      url: 'https://meet.google.com/abc-defg-hij',
      title: 'Google Meet',
      page: {
        controls: [{ label: 'Leave call' }],
        participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace speaking' }],
      },
    },
    zoom: {
      app: { name: 'Zoom Workplace' },
      window: { title: 'Zoom Meeting', controls: [{ label: 'Leave Meeting' }] },
      meeting_id: 'zoom-local',
      tiles: [{ id: 'mira', ariaLabel: 'Mira Patel is speaking' }],
    },
  },
});
assert.equal(appExecutionPlanMatrix.platform_count, 2);
assert.equal(appExecutionPlanMatrix.realtime_ready_count, 2);
assert.equal(kit.report({ platforms: ['google-meet'] }).meeting_app_adapter_execution_plan_matrix.platform_count, 1);
const appIntegrationPackage = kit.meetingAppAdapterIntegrationPackage(appCapability);
assert.equal(appIntegrationPackage.schema, 'meeting_app_adapter_integration_package');
assert.equal(appIntegrationPackage.platform, 'google_meet');
assert.equal(appIntegrationPackage.pilot_ready, true);
assert.equal(appIntegrationPackage.realtime_ready, true);
assert.equal(appIntegrationPackage.entrypoints.platform_kit, '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit');
assert.equal(appIntegrationPackage.runtime_delivery.adapter_route.local_observer_first, true);
assert.equal(appIntegrationPackage.runtime_delivery.messaging.runtime_event_endpoint.endsWith('/api/meeting-platform/runtime-events'), true);
assert.equal(kit.assertMeetingAppAdapterIntegrationPackage(appIntegrationPackage).platform, 'google_meet');
const appIntegrationPackageMatrix = kit.meetingAppAdapterIntegrationPackageMatrix({
  platforms: ['google-meet', 'zoom'],
  inputs: {
    google_meet: {
      url: 'https://meet.google.com/abc-defg-hij',
      title: 'Google Meet',
      page: {
        controls: [{ label: 'Leave call' }],
        participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace speaking' }],
      },
    },
    zoom: {
      app: { name: 'Zoom Workplace' },
      window: { title: 'Zoom Meeting', controls: [{ label: 'Leave Meeting' }] },
      meeting_id: 'zoom-local',
      tiles: [{ id: 'mira', ariaLabel: 'Mira Patel is speaking' }],
    },
  },
});
assert.equal(appIntegrationPackageMatrix.platform_count, 2);
assert.equal(appIntegrationPackageMatrix.realtime_ready_count, 2);
assert.equal(kit.assertMeetingAppAdapterIntegrationPackageMatrix(appIntegrationPackageMatrix).accepted, true);
assert.throws(
  () => kit.assertMeetingAppAdapterIntegrationPackageMatrix(appIntegrationPackageMatrix, { target: 'production' }),
  /Meeting app adapter integration package matrix acceptance failed/,
);
assert.equal(kit.report({ platforms: ['google-meet'] }).meeting_app_adapter_integration_package_matrix.platform_count, 1);
const appObserverPlan = kit.meetingAppRuntimeObserverPlan({
  url: 'https://meet.google.com/abc-defg-hij',
  page: {
    controls: [{ label: 'Leave call' }],
    participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
  },
}, {
  platform: 'google-meet',
});
assert.equal(appObserverPlan.accepted, true);
assert.equal(appObserverPlan.preflight_status, 'accepted');
assert.equal(appObserverPlan.observer_runtime.factory, 'createMeetingAppBrowserRuntime');
assert.equal(appObserverPlan.signal_contract.timestamp_field, 'captured_at_ms');
const appObserverPlanMatrix = kit.meetingAppRuntimeObserverPlanMatrix({
  platforms: ['google-meet', 'zoom'],
  inputs: {
    google_meet: {
      url: 'https://meet.google.com/abc-defg-hij',
      page: {
        controls: [{ label: 'Leave call' }],
        participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace speaking' }],
      },
    },
    zoom: {
      app: { name: 'Zoom Workplace' },
      window: { title: 'Zoom Meeting', controls: [{ label: 'Leave Meeting' }] },
      meeting_id: 'zoom-local',
      tiles: [{ id: 'mira', ariaLabel: 'Mira Patel is speaking' }],
    },
  },
});
assert.equal(appObserverPlanMatrix.platform_count, 2);
assert.equal(appObserverPlanMatrix.preflight_accepted_count, 2);
assert.equal(kit.report({ platforms: ['google-meet'] }).meeting_app_runtime_observer_plan_matrix.platform_count, 1);
const platformHostConfig = kit.platformRuntimeHostConfig('google-meet');
assert.equal(platformHostConfig.schema, 'meeting_platform_runtime_host_config');
assert.equal(platformHostConfig.readiness.host_ready, true);
assert.equal(platformHostConfig.driver.change_observer.enabled, true);
const platformHostConfigMatrix = kit.platformRuntimeHostConfigMatrix({ platforms: ['google-meet', 'zoom'] });
assert.equal(platformHostConfigMatrix.platform_count, 2);
assert.equal(platformHostConfigMatrix.host_ready_count, 2);
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_runtime_host_config_matrix.host_ready_count, 1);
const platformHostHandoff = kit.platformRuntimeHostHandoff('google-meet');
assert.equal(platformHostHandoff.schema, 'meeting_platform_runtime_host_handoff');
assert.equal(platformHostHandoff.acceptance.accepted, true);
const platformHostHandoffMatrix = kit.platformRuntimeHostHandoffMatrix({ platforms: ['google-meet', 'zoom'] });
assert.equal(platformHostHandoffMatrix.accepted_count, 2);
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_runtime_host_handoff_matrix.accepted_count, 1);
const kitRuntimeHost = kit.createPlatformRuntimeHost({
  inputProvider() {
    return { platform: 'google_meet', url: 'https://meet.google.com/abc-defg-hij' };
  },
  sample(options = {}) {
    return { ok: true, trigger: options.trigger };
  },
  stop() {
    return { stopped: true };
  },
}, 'google-meet', {
  setInterval() {
    return {};
  },
  clearInterval() {},
});
assert.equal(kitRuntimeHost.config.platform, 'google_meet');

const appGate = kit.meetingAppLaunchGate('google-meet', {
  allowFixtureProduction: true,
  requireProductionReady: false,
});
assert.equal(appGate.platform, 'google_meet');
assert.equal(appGate.passed, true);
assert.equal(appGate.runtime_ready, true);
assert.equal(appGate.evidence_level, 'fixture_dom');

const appGateSummary = kit.meetingAppLaunchGateSummary({
  allowFixtureProduction: true,
  requireProductionReady: false,
});
assert.equal(appGateSummary.ok, true);
assert.equal(appGateSummary.gates.length, 5);

const kitRecorder = kit.meetingAppSnapshotRecorder({ captureProfile: 'google_meet' });
kitRecorder.add(kit.meetingAppFixture('google-meet'), { phase: 'active', label: 'kit-active' });
assert.equal(kitRecorder.getState().record_count, 1);
assert.equal(kit.meetingAppGateInputFromRecords(kitRecorder.exportRecords()).snapshots.google_meet.length, 1);

const integrationProfile = kit.meetingAppIntegrationProfile('google-meet');
assert.equal(integrationProfile.platform, 'google_meet');
assert.equal(integrationProfile.extension.matches.includes('https://meet.google.com/*'), true);
assert.equal(integrationProfile.runtime.runtimePreset, 'google_meet');
assert.equal(integrationProfile.event_model.realtime_axis.primary, 'browser_extension_local_observer');
const integrationProfiles = kit.allMeetingAppIntegrationProfiles({ platforms: ['zoom'] });
assert.deepEqual(Object.keys(integrationProfiles), ['zoom']);
const integrationMatrix = kit.meetingAppIntegrationMatrix({ platforms: ['google-meet', 'teams'] });
assert.deepEqual(integrationMatrix.platforms, ['google_meet', 'microsoft_teams']);
assert.equal(integrationMatrix.rows.every((row) => row.sdk_wiring_ready === true), true);
const deploymentManifest = kit.meetingAppDeploymentManifest('google-meet');
assert.equal(deploymentManifest.platform, 'google_meet');
assert.equal(deploymentManifest.runtime_contract.timestamp_field, 'captured_at_ms');
assert.equal(deploymentManifest.extension_install_plan.platforms.includes('google_meet'), true);
const deploymentManifestAcceptance = kit.meetingAppDeploymentManifestAcceptance(deploymentManifest);
assert.equal(deploymentManifestAcceptance.accepted, true);
assert.equal(deploymentManifestAcceptance.production_ready, false);
assert.equal(kit.assertMeetingAppDeploymentManifest(deploymentManifest).accepted, true);
const deploymentManifests = kit.allMeetingAppDeploymentManifests({ platforms: ['teams'] });
assert.deepEqual(Object.keys(deploymentManifests), ['microsoft_teams']);
assert.equal(deploymentManifests.microsoft_teams.integration_targets.some((item) => item.surface === 'chrome_or_edge_extension'), true);
const deploymentManifestAcceptances = kit.allMeetingAppDeploymentManifestAcceptanceReports({ platforms: ['teams'] });
assert.deepEqual(Object.keys(deploymentManifestAcceptances), ['microsoft_teams']);
assert.equal(deploymentManifestAcceptances.microsoft_teams.accepted, true);
const deploymentManifestSummary = kit.meetingAppDeploymentManifestAcceptanceSummary({ platforms: ['google-meet', 'teams'] });
assert.equal(deploymentManifestSummary.accepted, true);
assert.equal(deploymentManifestSummary.production_ready_count, 0);
const runtimeAdapterConfig = kit.meetingAppRuntimeAdapterConfig('google-meet');
assert.equal(runtimeAdapterConfig.platform, 'google_meet');
assert.equal(runtimeAdapterConfig.bridge_options.browser_runtime_preset, 'google_meet');
assert.equal(runtimeAdapterConfig.capture_options.captureProfile, 'google_meet');
const runtimeAdapterProfile = kit.meetingAppRuntimeAdapterProfile('https://meet.google.com/abc-defg-hij');
assert.equal(runtimeAdapterProfile.detected, true);
assert.equal(runtimeAdapterProfile.platform, 'google_meet');
assert.equal(runtimeAdapterProfile.runtime_config.capture_options.captureProfile, 'google_meet');
assert.equal(runtimeAdapterProfile.tracks.output_intents.includes('speaker_track'), true);
assert.equal(kit.meetingAppRuntimeAdapterProfile('https://example.com/not-a-meeting').detected, false);
const runtimeAdapterProfileMatrix = kit.meetingAppRuntimeAdapterProfileMatrix({
  platforms: ['google-meet', 'teams'],
});
assert.equal(runtimeAdapterProfileMatrix.platform_count, 2);
assert.equal(runtimeAdapterProfileMatrix.detected_count, 2);
assert.deepEqual(runtimeAdapterProfileMatrix.platforms, ['google_meet', 'microsoft_teams']);
assert.equal(runtimeAdapterProfileMatrix.rows.every((row) => row.track_output_intents.includes('speaker_track')), true);
const runtimeAdapterSelection = kit.selectMeetingAppRuntimeAdapter({
  tabs: [
    { active: false, url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting', title: 'Teams call' },
    { active: true, audible: true, url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet' },
  ],
}, {
  platforms: ['google-meet', 'teams'],
  observedAtMs: 1_783_356_000_000,
});
assert.equal(runtimeAdapterSelection.selected, true);
assert.equal(runtimeAdapterSelection.platform, 'google_meet');
assert.equal(runtimeAdapterSelection.launch.runtime_options.runtimePreset, 'google_meet');
assert.equal(runtimeAdapterSelection.supported_candidate_count, 2);
const runtimeAdapterHandoff = kit.meetingAppRuntimeAdapterHandoff(runtimeAdapterSelection, {
  surface: 'native-detector',
});
assert.equal(runtimeAdapterHandoff.selected, true);
assert.equal(runtimeAdapterHandoff.surface, 'native_detector');
assert.equal(runtimeAdapterHandoff.install.required_capabilities.includes('capture_window_snapshot'), true);
assert.equal(runtimeAdapterHandoff.tracks.options.platform, 'google_meet');
assert.equal(kit.meetingAppRuntimeAdapterHandoffAcceptance(runtimeAdapterHandoff).accepted, true);
assert.equal(kit.assertMeetingAppRuntimeAdapterHandoff(runtimeAdapterHandoff).accepted, true);
const runtimeAdapterHandoffMatrix = kit.meetingAppRuntimeAdapterHandoffMatrix({
  platforms: ['google-meet', 'teams'],
  surfaces: ['browser-extension', 'electron-webview'],
});
assert.equal(runtimeAdapterHandoffMatrix.handoff_count, 4);
assert.equal(runtimeAdapterHandoffMatrix.ready_count, 4);
assert.equal(runtimeAdapterHandoffMatrix.rows.find((row) => row.surface === 'electron_webview').start_mode, 'host_injected_runtime');
assert.equal(kit.meetingAppRuntimeAdapterHandoffMatrixAcceptance(runtimeAdapterHandoffMatrix).accepted, true);
assert.equal(kit.assertMeetingAppRuntimeAdapterHandoffMatrix(runtimeAdapterHandoffMatrix).accepted, true);
const runtimeAdapterHostPackage = kit.meetingAppRuntimeAdapterHostPackage({
  platforms: ['google-meet', 'teams'],
  surfaces: ['browser-extension'],
});
assert.equal(runtimeAdapterHostPackage.accepted, true);
assert.equal(runtimeAdapterHostPackage.handoff_count, 2);
assert.equal(runtimeAdapterHostPackage.sdk.required_methods.includes('meetingAppRuntimeAdapterHandoffMatrixAcceptance'), true);
const capturePlan = kit.meetingAppLiveSnapshotCapturePlan('google-meet');
assert.equal(capturePlan.platform, 'google_meet');
assert.equal(capturePlan.required_snapshots.some((item) => item.id === 'active_speaker'), true);
const capturePlans = kit.allMeetingAppLiveSnapshotCapturePlans({ platforms: ['zoom'] });
assert.deepEqual(Object.keys(capturePlans), ['zoom']);
assert.equal(capturePlans.zoom.minimum_record_count, 2);
assert.equal(kit.meetingAppRuntimeAdapterAcceptance(runtimeAdapterConfig).accepted, true);
assert.equal(kit.assertMeetingAppRuntimeAdapterConfig(runtimeAdapterConfig).accepted, true);
assert.equal(kit.meetingAppRuntimeAdapterValidation(runtimeAdapterConfig).accepted, false);
const kitLiveSnapshots = [
  buildMeetingAppFixtureSnapshot('google-meet', { state: 'active', observedAtMs: 1_783_356_000_000 }),
  buildMeetingAppFixtureSnapshot('google-meet', { state: 'prejoin', observedAtMs: 1_783_356_600_000 }),
];
assert.equal(kit.meetingAppRuntimeAdapterValidation(runtimeAdapterConfig, { snapshots: kitLiveSnapshots }).production_ready, true);
assert.equal(kit.assertMeetingAppRuntimeAdapterValidation(runtimeAdapterConfig, { snapshots: kitLiveSnapshots }).accepted, true);
const kitEvidencePackage = kit.meetingAppLiveEvidencePackage({
  packageId: 'kit-google-evidence',
  platforms: ['google-meet'],
  snapshots: kitLiveSnapshots,
});
assert.equal(kitEvidencePackage.id, 'kit-google-evidence');
assert.equal(kitEvidencePackage.production_ready, true);
assert.equal(kitEvidencePackage.manifest_acceptance.google_meet.production_ready, true);
const kitEvidenceSummary = kit.meetingAppLiveEvidencePackageSummary(kitEvidencePackage);
assert.equal(kitEvidenceSummary.production_ready_count, 1);
assert.equal(kitEvidenceSummary.rows[0].record_count, 2);
const kitDomDiagnosis = kit.meetingAppDomAdaptationDiagnosis('google-meet', {
  snapshots: kitLiveSnapshots,
});
assert.equal(kitDomDiagnosis.platform, 'google_meet');
assert.equal(kitDomDiagnosis.production_ready, true);
assert.equal(kitDomDiagnosis.selector_probe.matched.active_speaker, true);
assert.equal(kitDomDiagnosis.observer_probe.signal_types.includes('meeting_ended'), true);
const kitDomDiagnoses = kit.allMeetingAppDomAdaptationDiagnoses({
  platforms: ['zoom'],
  snapshots: {
    zoom: [
      kit.meetingAppFixture('zoom', { state: 'active', observedAtMs: 1_783_356_000_000 }),
      kit.meetingAppFixture('zoom', { state: 'prejoin', observedAtMs: 1_783_356_600_000 }),
    ],
  },
});
assert.deepEqual(Object.keys(kitDomDiagnoses), ['zoom']);
assert.equal(kitDomDiagnoses.zoom.accepted, true);
const kitDomDiagnosisMatrix = kit.meetingAppDomAdaptationDiagnosisMatrix({
  platforms: ['google-meet'],
  snapshots: {
    'google-meet': kitLiveSnapshots,
  },
});
assert.equal(kitDomDiagnosisMatrix.platform_count, 1);
assert.equal(kitDomDiagnosisMatrix.accepted_count, 1);
assert.equal(kitDomDiagnosisMatrix.rows[0].active_speaker_matched, true);
const runtimeAdapterConfigs = kit.allMeetingAppRuntimeAdapterConfigs({ platforms: ['zoom'] });
assert.deepEqual(Object.keys(runtimeAdapterConfigs), ['zoom']);
assert.equal(runtimeAdapterConfigs.zoom.runtime_options.runtimePreset, 'zoom');
const runtimeAdapterAcceptance = kit.allMeetingAppRuntimeAdapterAcceptanceReports({ platforms: ['zoom'] });
assert.equal(runtimeAdapterAcceptance.zoom.accepted, true);
const runtimeAdapterValidation = kit.allMeetingAppRuntimeAdapterValidationReports({ platforms: ['zoom'] });
assert.equal(runtimeAdapterValidation.zoom.accepted, false);

const googleContract = kit.platformAdapterContract('google-meet');
assert.equal(googleContract.platform, 'google_meet');
assert.equal(googleContract.realtime_axis.rules.includes('use_provider_events_only_for_reconcile_and_backfill'), true);
assert.equal(googleContract.provider_observer.required_for_realtime, false);
assert.equal(googleContract.annotations.endpoints.insertMark, `${baseUrl}/api/annotations`);
const contractMatrix = kit.platformAdapterContractMatrix({ platforms: ['google-meet', 'teams'] });
assert.equal(contractMatrix.platform_count, 2);
assert.equal(contractMatrix.browser_observer_count, 2);
assert.equal(contractMatrix.rows.find((row) => row.platform === 'microsoft_teams').provider_start_event_count, 1);
assert.equal(kit.platformAdapterContractAcceptance('google-meet').accepted, true);
assert.equal(kit.platformAdapterContractAcceptanceMatrix({ platforms: ['google-meet', 'teams'] }).accepted_count, 2);
assert.equal(kit.assertPlatformAdapterContract('google-meet').accepted, true);

const extensionMatches = kit.meetingAppExtensionMatches({ platforms: ['google-meet'] });
assert.deepEqual(extensionMatches.matches, ['https://meet.google.com/*']);
const extensionManifest = kit.meetingAppContentScriptManifest({
  platforms: ['google-meet'],
  js: ['content.js'],
});
assert.equal(extensionManifest.manifest_version, 3);
assert.deepEqual(extensionManifest.content_scripts[0].matches, ['https://meet.google.com/*']);
assert.deepEqual(extensionManifest.content_scripts[0].js, ['content.js']);
const extensionPlan = kit.meetingAppExtensionInstallPlan();
assert.equal(extensionPlan.content_script_adapter, '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime');
assert.equal(extensionPlan.meeting_app_content_script_adapter, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-content-script');
assert.equal(extensionPlan.platform_integration_runtime_adapter, '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime');
assert.equal(extensionPlan.platforms.length, 5);
assert.equal(extensionPlan.runtime_contract.message_types.client_call, 'meeting_timeline.client_call');
assert.equal(extensionPlan.runtime_contract.message_types.observe_candidates, 'meeting_timeline.observe_candidates');
assert.equal(extensionPlan.runtime_contract.message_types.preflight_current_window, 'meeting_timeline.preflight_current_window');
assert.equal(extensionPlan.runtime_contract.message_types.preflight_candidates, 'meeting_timeline.preflight_candidates');
assert.equal(kit.normalizeMeetingAppExtensionMessageType('attached'), 'meeting_timeline.extension_attached');
assert.equal(kit.normalizeMeetingAppExtensionMessageType('observe-platform-candidates'), 'meeting_timeline.observe_candidates');
assert.equal(kit.normalizeMeetingAppExtensionMessageType('current-window-preflight'), 'meeting_timeline.preflight_current_window');
assert.equal(kit.normalizeMeetingAppExtensionMessageType('preflight-platform-candidates'), 'meeting_timeline.preflight_candidates');
assert.equal(kit.meetingAppExtensionTimelineEndpoint('endMeeting'), '/api/meeting-session/end');
assert.deepEqual(kit.meetingAppExtensionAttachedMessage({
  platform: 'google-meet',
  capturedAtMs: 234,
  url: 'https://meet.google.com/abc-defg-hij',
}), {
  type: 'meeting_timeline.extension_attached',
  platform: 'google_meet',
  captured_at_ms: 234,
  url: 'https://meet.google.com/abc-defg-hij',
});
assert.deepEqual(kit.meetingAppExtensionObserveCandidatesMessage({
  capturedAtMs: 236,
  tabs: [{ url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet', active: true }],
}), {
  type: 'meeting_timeline.observe_candidates',
  captured_at_ms: 236,
  tabs: [{ url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet', active: true }],
});
assert.deepEqual(kit.meetingAppExtensionClientCallMessage('insertMark', {
  label: 'why?',
}, {
  platform: 'google-meet',
  capturedAtMs: 235,
}), {
  type: 'meeting_timeline.client_call',
  method: 'insertMark',
  platform: 'google_meet',
  captured_at_ms: 235,
  input: {
    label: 'why?',
  },
});
const extensionScaffold = kit.meetingAppExtensionScaffold({
  platforms: ['google-meet'],
  baseUrl,
});
assert.equal(extensionScaffold.type, 'meeting_app_extension_scaffold');
assert.equal(extensionScaffold.manifest.host_permissions.includes(`${baseUrl}/*`), true);
assert.equal(extensionScaffold.files.some((file) => file.path === 'src/content-script.entry.mjs'), true);
const extensionAcceptance = kit.meetingAppExtensionAcceptance({
  platforms: ['google-meet'],
  baseUrl,
});
assert.equal(extensionAcceptance.accepted, true);
assert.equal(extensionAcceptance.accepted_platform_count, 1);
assert.equal(kit.assertMeetingAppExtensionScaffold({
  platforms: ['google-meet'],
  baseUrl,
}).accepted, true);

assert.equal(kit.assertMeetingAppLaunchGate('webex', {
  allowFixtureProduction: true,
  requireProductionReady: false,
}).passed, true);
assert.equal(kit.assertAllMeetingAppLaunchGates({
  allowFixtureProduction: true,
  requireProductionReady: false,
}).ok, true);

const googleStart = await kit.handleWebhook({
  method: 'POST',
  url: `${baseUrl}${basePath}/google-meet`,
  body: buildPlatformFixtureEvent('google-meet', 'meeting_start'),
});
assert.equal(googleStart.status, 200);
assert.equal(googleStart.body.ok, true);
assert.equal(calls.at(-1).method, 'startMeeting');
assert.equal(calls.at(-1).input.platform, 'google_meet');
assert.equal(calls.at(-1).input.meeting_id, 'google-fixture-001');

await kit.insertMark({
  id: 'kit-mark-1',
  capturedAtMs: 1_782_442_810_000,
  label: 'why?',
});
assert.equal(calls.at(-1).method, 'insertMark');
assert.equal(calls.at(-1).input.id, 'kit-mark-1');

assert.equal(kit.platformParticipantTrackPlan('google-meet').schema, 'meeting_platform_participant_track_plan');
assert.equal(kit.platformParticipantTrackMatrix({ platforms: ['google-meet', 'zoom'] }).platform_count, 2);
assert.equal(kit.platformParticipantTrack('google-meet', {
  signals: [
    {
      type: 'participant_joined',
      meeting: { platform: 'google_meet', meeting_id: 'abc-defg-hij' },
      occurred_at_ms: 1_782_614_400_000,
      participant_id: 'ada',
      participant_name: 'Ada',
    },
  ],
}).marks[0].intent, 'participant_track');
assert.equal(kit.meetingAppTrackPipeline([
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1_782_614_400_000,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1_782_614_400_400,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
], {
  speakerTrackOptions: {
    minStableMs: 250,
    minSegmentMs: 0,
    closeOpenSegmentsAtMs: 1_782_614_401_000,
  },
}).marks[0].intent, 'speaker_track');
assert.equal(kit.createMeetingAppTrackPipeline().getSnapshots().length, 0);
const appTrackRuntime = kit.meetingAppTrackRuntime({
  speakerTrackOptions: {
    minStableMs: 250,
    minSegmentMs: 0,
    closeOpenSegmentsAtMs: 1_782_614_401_000,
  },
});
const appTrackRuntimeObserved = await appTrackRuntime.observe([
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1_782_614_400_000,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1_782_614_400_400,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
]);
assert.equal(appTrackRuntimeObserved.new_mark_count, 1);
assert.equal(calls.at(-1).method, 'insertMarks');
assert.equal(calls.at(-1).input[0].intent, 'speaker_track');
assert.equal(kit.platformTimelineViewPlan('google-meet').schema, 'meeting_platform_timeline_view_plan');
assert.equal(kit.platformTimelineViewMatrix({ platforms: ['google-meet', 'zoom'] }).platform_count, 2);
assert.equal(kit.platformTimelineView('google-meet', {
  meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    start_time_ms: 1_782_614_400_000,
  },
  annotations: [
    {
      id: 'view-note-1',
      label: 'why?',
      captured_at_ms: 1_782_614_401_000,
    },
  ],
}).visible_markers[0].id, 'view-note-1');
assert.equal(kit.zoomPlatformTimelineViewport({
  start_ms: 0,
  duration_ms: 600_000,
  full_duration_ms: 600_000,
}, 2).duration_ms, 300_000);
assert.equal(kit.platformAnnotationIntakePlan('google-meet').schema, 'meeting_platform_annotation_intake_plan');
assert.equal(kit.platformAnnotationIntakeMatrix({ platforms: ['google-meet', 'zoom'] }).platform_count, 2);
assert.equal(kit.platformAnnotationIntake('google-meet', {
  current_meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    start_time_ms: 1_782_614_400_000,
  },
  annotation: {
    id: 'intake-note-1',
    label: 'why?',
    captured_at_ms: 1_782_614_402_000,
  },
}).status, 'ready_to_insert_current_axis');
assert.equal(kit.platformAnnotationIntake('lark', {
  annotation: {
    id: 'intake-note-2',
    label: 'follow up',
    captured_at_ms: 1_782_614_403_000,
  },
}).status, 'start_open_session_then_insert');
assert.equal(kit.platformClockSyncPlan('google-meet').schema, 'meeting_platform_clock_sync_plan');
assert.equal(kit.platformClockSyncMatrix({ platforms: ['google-meet', 'zoom'] }).platform_count, 2);
assert.equal(kit.platformClockSync('google-meet', {
  samples: [
    {
      client_send_at_ms: 1_782_614_400_000,
      server_time_ms: 1_782_614_400_060,
      client_receive_at_ms: 1_782_614_400_100,
    },
  ],
}).status, 'clock_sync_ready');
assert.equal(kit.platformSessionBindingPlan('google-meet').schema, 'meeting_platform_session_binding_plan');
assert.equal(kit.platformSessionBindingMatrix({ platforms: ['google-meet', 'zoom'] }).platform_count, 2);
assert.equal(kit.platformSessionBinding('google-meet', {
  current_meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    start_time_ms: 1_782_614_400_000,
  },
  local_observer: {
    url: 'https://meet.google.com/abc-defg-hij',
    observed_at_ms: 1_782_614_401_000,
  },
}).status, 'bound_to_current_axis');
assert.equal(kit.platformRealtimeAnnotationPlan('google-meet').schema, 'meeting_platform_realtime_annotation_plan');
assert.equal(kit.platformRealtimeAnnotationMatrix({ platforms: ['google-meet', 'zoom'] }).platform_count, 2);
assert.equal(kit.platformRealtimeAnnotation('google-meet', {
  clock_sync: {
    offset_ms: 0,
    rtt_ms: 40,
  },
  current_meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    start_time_ms: 1_782_614_400_000,
  },
  local_observer: {
    url: 'https://meet.google.com/abc-defg-hij',
    observed_at_ms: 1_782_614_401_000,
  },
  annotation: {
    id: 'realtime-note-1',
    label: 'why?',
    captured_at_ms: 1_782_614_402_000,
  },
}).status, 'ready_to_insert');

const kitState = kit.getState();
assert.equal(kitState.bridge.signal_reconciler.active_meetings.length, 0);
assert.equal(kitState.webhook_router.active_meetings.length, 1);
assert.equal(kit.reset().webhook_router.seen.length, 0);

const report = buildMeetingPlatformKitReport({
  baseUrl,
  basePath,
  env,
});
assert.equal(report.supported_platforms.length, 6);
assert.equal(report.supported_meeting_app_platforms.length, 5);
assert.equal(report.webhook_router.base_path, basePath);
assert.equal(report.fixture_acceptance.accepted_count, 6);
assert.equal(report.meeting_app_extension_install_plan.platforms.length, 5);
assert.equal(report.meeting_app_extension_install_plan.matches.includes('https://meet.google.com/*'), true);
assert.equal(report.meeting_app_extension_acceptance.accepted, true);
assert.equal(report.meeting_app_extension_acceptance.accepted_platform_count, 5);
assert.equal(report.meeting_app_integration_matrix.platform_count, 5);
assert.equal(report.meeting_app_integration_matrix.rows.some((row) => row.platform === 'google_meet'), true);
assert.equal(report.meeting_app_deployment_manifests.google_meet.production_gate.requires_captured_dom, true);
assert.equal(report.meeting_app_deployment_manifest_acceptance.google_meet.accepted, true);
assert.equal(report.meeting_app_deployment_manifest_acceptance_summary.accepted_count, 5);
assert.equal(report.meeting_app_dom_adaptation_diagnosis.google_meet.accepted, false);
assert.equal(report.meeting_app_dom_adaptation_diagnosis_matrix.platform_count, 5);
assert.equal(report.meeting_app_dom_adaptation_diagnosis_matrix.rows.find((row) => row.platform === 'google_meet').accepted, false);
assert.equal(report.meeting_app_runtime_adapter_configs.google_meet.bridge_options.browser_runtime_preset, 'google_meet');
assert.equal(report.meeting_app_runtime_adapter_acceptance.google_meet.accepted, true);
assert.equal(report.meeting_app_live_snapshot_capture_plans.google_meet.minimum_record_count, 2);
assert.equal(report.meeting_app_runtime_adapter_validation.google_meet.accepted, false);
assert.equal(report.meeting_app_fixture_acceptance.accepted, true);
assert.equal(report.meeting_app_fixture_acceptance.accepted_count, 5);
assert.equal(report.meeting_app_launch_gate.ok, false);
assert.equal(report.meeting_app_launch_gate.gates.length, 5);
assert.equal(report.platform_live_adapter_matrix.platform_count, 6);
assert.equal(report.platform_live_adapter_matrix.rows.some((row) => row.platform === 'google_meet'), true);
assert.equal(report.platform_live_adapter_readiness_matrix.platform_count, 6);
assert.equal(report.platform_live_adapter_readiness_matrix.rows.some((row) => row.platform === 'google_meet'), true);
assert.equal(report.platform_live_adapter_handoff_bundle.platform_count, 6);
assert.equal(report.platform_live_adapter_handoff_bundle.sdk.factory, 'createMeetingPlatformLiveAdapter');
assert.equal(report.platform_live_adapter_handoff_bundle.host_contract.annotation_timestamp_field, 'captured_at_ms');
assert.equal(report.platform_live_adapter_handoff_bundle.handoffs.some((handoff) => handoff.platform === 'google_meet'), true);
assert.equal(report.platform_host_integration.runtime_contract.annotation_timestamp_field, 'captured_at_ms');
assert.equal(report.platform_host_integration_acceptance.accepted, true);
assert.equal(report.platform_host_integration_acceptance.scaffold.files.some((file) => file.path === 'src/meeting-platform-host.mjs'), true);
assert.equal(report.platform_provider_connection_matrix.platform_count, 6);
assert.equal(report.platform_provider_connection_matrix.rows.some((row) => row.platform === 'google_meet'), true);
assert.equal(report.platform_provider_connection_matrix.rows.find((row) => row.platform === 'google_meet').docs.some((url) => url.includes('developers.google.com')), true);
assert.equal(report.platform_participant_track_matrix.platform_count, 6);
assert.equal(report.platform_participant_track_matrix.provider_blocking_count, 0);
assert.equal(report.platform_participant_track_matrix.rows.some((row) => row.platform === 'google_meet'), true);
assert.equal(report.platform_timeline_view_matrix.platform_count, 6);
assert.equal(report.platform_timeline_view_matrix.provider_blocking_count, 0);
assert.equal(report.platform_timeline_view_matrix.rows.some((row) => row.platform === 'google_meet'), true);
assert.equal(report.platform_annotation_intake_matrix.platform_count, 6);
assert.equal(report.platform_annotation_intake_matrix.provider_blocking_count, 0);
assert.equal(report.platform_annotation_intake_matrix.rows.some((row) => row.platform === 'google_meet'), true);
assert.equal(report.platform_clock_sync_matrix.platform_count, 6);
assert.equal(report.platform_clock_sync_matrix.provider_blocking_count, 0);
assert.equal(report.platform_clock_sync_matrix.rows.some((row) => row.platform === 'google_meet'), true);
assert.equal(report.platform_session_binding_matrix.platform_count, 6);
assert.equal(report.platform_session_binding_matrix.provider_blocking_count, 0);
assert.equal(report.platform_session_binding_matrix.rows.some((row) => row.platform === 'google_meet'), true);
assert.equal(report.platform_realtime_annotation_matrix.platform_count, 6);
assert.equal(report.platform_realtime_annotation_matrix.provider_blocking_count, 0);
assert.equal(report.platform_realtime_annotation_matrix.rows.some((row) => row.platform === 'google_meet'), true);
assert.equal(report.platform_adapter_contract_matrix.platform_count, 6);
assert.equal(report.platform_adapter_contract_matrix.rows.some((row) => row.platform === 'google_meet'), true);
assert.equal(report.platform_adapter_contract_matrix.rows.find((row) => row.platform === 'google_meet').browser_observer, true);
assert.equal(report.platform_adapter_contract_acceptance_matrix.platform_count, 6);
assert.equal(report.platform_adapter_contract_acceptance_matrix.accepted_count, 6);

console.log('ok meeting platform timeline kit');
