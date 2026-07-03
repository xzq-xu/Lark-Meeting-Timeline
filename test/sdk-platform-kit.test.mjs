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
