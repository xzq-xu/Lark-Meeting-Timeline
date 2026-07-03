import assert from 'node:assert/strict';

import {
  buildMeetingPlatformKitReport,
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import {
  buildPlatformFixtureEnv,
  buildPlatformFixtureEvent,
} from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';
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
const runtimeAdapterConfig = kit.meetingAppRuntimeAdapterConfig('google-meet');
assert.equal(runtimeAdapterConfig.platform, 'google_meet');
assert.equal(runtimeAdapterConfig.bridge_options.browser_runtime_preset, 'google_meet');
assert.equal(runtimeAdapterConfig.capture_options.captureProfile, 'google_meet');
const runtimeAdapterConfigs = kit.allMeetingAppRuntimeAdapterConfigs({ platforms: ['zoom'] });
assert.deepEqual(Object.keys(runtimeAdapterConfigs), ['zoom']);
assert.equal(runtimeAdapterConfigs.zoom.runtime_options.runtimePreset, 'zoom');

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
assert.equal(extensionPlan.content_script_adapter, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-content-script');
assert.equal(extensionPlan.platforms.length, 5);
assert.equal(extensionPlan.runtime_contract.message_types.client_call, 'meeting_timeline.client_call');
assert.equal(kit.normalizeMeetingAppExtensionMessageType('attached'), 'meeting_timeline.extension_attached');
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
assert.equal(report.meeting_app_runtime_adapter_configs.google_meet.bridge_options.browser_runtime_preset, 'google_meet');
assert.equal(report.meeting_app_fixture_acceptance.accepted, true);
assert.equal(report.meeting_app_fixture_acceptance.accepted_count, 5);
assert.equal(report.meeting_app_launch_gate.ok, false);
assert.equal(report.meeting_app_launch_gate.gates.length, 5);

console.log('ok meeting platform timeline kit');
