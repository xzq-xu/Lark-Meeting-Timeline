import assert from 'node:assert/strict';

import {
  DEFAULT_MEETING_APP_TIMELINE_SDK_PLATFORMS,
  createMeetingAppTimelineSdk,
} from '../packages/meeting-timeline-sdk/index.mjs';

const calls = [];
const baseUrl = 'https://timeline.example.com';
const fetchImpl = async (url, init = {}) => {
  calls.push({
    url: String(url),
    method: init.method ?? 'GET',
    body: init.body ? JSON.parse(String(init.body)) : null,
  });
  return new Response(JSON.stringify({
    ok: true,
    url: String(url),
    body: init.body ? JSON.parse(String(init.body)) : null,
  }), {
    headers: { 'content-type': 'application/json' },
  });
};

assert.deepEqual(DEFAULT_MEETING_APP_TIMELINE_SDK_PLATFORMS, [
  'google-meet',
  'teams',
  'zoom',
  'webex',
  'lark',
]);

const sdk = createMeetingAppTimelineSdk({
  baseUrl,
  fetch: fetchImpl,
  platforms: ['google-meet', 'zoom'],
});

assert.equal(sdk.schema, 'meeting_app_timeline_sdk');
assert.deepEqual(sdk.platforms, ['google_meet', 'zoom']);
assert.equal(sdk.client.baseUrl, baseUrl);
assert.equal(sdk.runtime.schema, 'meeting_platform_integration_runtime');
assert.equal(sdk.runtimeEvents.endpoint, `${baseUrl}/api/meeting-platform/runtime-events`);
assert.equal(sdk.kit.platformRegistryManifest({ platforms: ['google-meet'] }).platform_count, 1);

const detected = sdk.detect({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Google Meet',
});
assert.equal(detected.detected, true);
assert.equal(detected.platform, 'google_meet');

const packageMatrix = sdk.integrationPackageMatrix({
  inputs: {
    google_meet: {
      url: 'https://meet.google.com/abc-defg-hij',
      page: {
        controls: [{ label: 'Leave call' }],
        participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
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
assert.equal(packageMatrix.schema, 'meeting_app_adapter_integration_package_matrix');
assert.equal(packageMatrix.platform_count, 2);
assert.equal(packageMatrix.realtime_ready_count, 2);

const integrationProfile = sdk.integrationProfile('google-meet');
assert.equal(integrationProfile.schema, 'meeting_app_integration_profile');
assert.equal(integrationProfile.platform, 'google_meet');
assert.equal(integrationProfile.event_model.realtime_axis.timestamp_invariant, 'annotations_use_absolute_captured_at_ms');

const adapterProfile = sdk.adapterProfile('https://meet.google.com/abc-defg-hij');
assert.equal(adapterProfile.schema, 'meeting_app_runtime_adapter_profile_resolution');
assert.equal(adapterProfile.detected, true);
assert.equal(adapterProfile.platform, 'google_meet');
assert.equal(adapterProfile.tracks.output_intents.includes('speaker_track'), true);

const observerPlan = sdk.observerPlan({
  url: 'https://meet.google.com/abc-defg-hij',
  page: {
    controls: [{ label: 'Leave call' }],
    participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
  },
}, {
  surface: 'browser-extension',
});
assert.equal(observerPlan.schema, 'meeting_app_runtime_observer_plan');
assert.equal(observerPlan.platform, 'google_meet');
assert.equal(observerPlan.accepted, true);
assert.equal(observerPlan.signal_contract.timestamp_field, 'captured_at_ms');
assert.equal(observerPlan.track_runtime.output_intents.includes('participant_track'), true);

const handoff = sdk.handoff('https://meet.google.com/abc-defg-hij', {
  surface: 'browser-extension',
});
assert.equal(handoff.schema, 'meeting_app_runtime_adapter_handoff');
assert.equal(handoff.platform, 'google_meet');
assert.equal(handoff.readiness.ready_to_start, true);
assert.equal(handoff.annotations.provider_events_block_realtime, false);
assert.equal(sdk.handoffAcceptance(handoff).accepted, true);

const handoffMatrix = sdk.handoffMatrix({
  surfaces: ['browser-extension', 'native-detector'],
});
assert.equal(handoffMatrix.schema, 'meeting_app_runtime_adapter_handoff_matrix');
assert.equal(handoffMatrix.platform_count, 2);
assert.equal(handoffMatrix.handoff_count, 4);
assert.equal(handoffMatrix.ready_count, 4);

const hostPackage = sdk.hostPackage({
  surfaces: ['browser-extension', 'native-detector'],
});
assert.equal(hostPackage.schema, 'meeting_app_runtime_adapter_host_package');
assert.equal(hostPackage.accepted, true);
assert.equal(hostPackage.runtime_contract.annotation_timestamp_field, 'captured_at_ms');
assert.equal(hostPackage.ci_gates.includes('require_provider_and_transcript_non_blocking'), true);

const connectorPackage = sdk.connectorPackage({
  surfaces: ['browser-extension', 'native-detector'],
  observeTracks: true,
});
assert.equal(connectorPackage.schema, 'meeting_app_timeline_connector_package');
assert.equal(connectorPackage.accepted, true);
assert.deepEqual(connectorPackage.surfaces, ['browser_extension', 'native_detector']);
assert.equal(connectorPackage.host_package.schema, 'meeting_app_runtime_adapter_host_package');
assert.equal(connectorPackage.observer_plan_by_surface.browser_extension.platform_count, 2);
assert.equal(connectorPackage.scheduler_config_by_surface.native_detector.platform_count, 2);
assert.equal(connectorPackage.scheduler_config_by_surface.browser_extension.track_enabled_count, 2);
assert.equal(connectorPackage.extension.acceptance.accepted, true);
assert.equal(connectorPackage.extension.manifest.permissions.includes('tabs'), true);
assert.equal(connectorPackage.runtime_events.plan_matrix.platform_count, 2);
assert.equal(connectorPackage.contracts.provider_events_block_realtime, false);
assert.equal(connectorPackage.entrypoints.some((entry) => entry.id === 'speaker-participant-track'), true);
const connectorReleaseGate = sdk.connectorReleaseGate(connectorPackage);
assert.equal(connectorReleaseGate.schema, 'meeting_app_timeline_connector_release_gate');
assert.equal(connectorReleaseGate.accepted, true);
assert.equal(connectorReleaseGate.target, 'pilot');
assert.equal(connectorReleaseGate.pilot_ready_count, 2);
assert.equal(sdk.assertConnectorReleaseGate(connectorPackage).accepted, true);
const productionConnectorReleaseGate = sdk.connectorReleaseGate(connectorPackage, {
  target: 'production',
});
assert.equal(productionConnectorReleaseGate.accepted, false);
assert.equal(productionConnectorReleaseGate.issues.some((issue) => issue.code.includes('production_evidence_accepted')), true);
const connectorPlatformRoadmap = sdk.connectorPlatformRoadmap(connectorPackage);
assert.equal(connectorPlatformRoadmap.schema, 'meeting_app_timeline_connector_platform_roadmap');
assert.equal(connectorPlatformRoadmap.accepted, true);
assert.equal(connectorPlatformRoadmap.rows.find((row) => row.platform === 'google_meet').recommended_first_surface, 'browser_extension');
assert.equal(connectorPlatformRoadmap.rows.find((row) => row.platform === 'zoom').recommended_first_surface, 'native_detector');
assert.equal(sdk.assertConnectorPlatformRoadmap(connectorPackage).accepted, true);
const connectorAdapterMatrix = sdk.connectorAdapterMatrix(connectorPackage);
assert.equal(connectorAdapterMatrix.schema, 'meeting_app_timeline_connector_adapter_matrix');
assert.equal(connectorAdapterMatrix.accepted, true);
assert.equal(connectorAdapterMatrix.rows.find((row) => row.platform === 'google_meet').adapter_mode, 'browser_content_script');
assert.equal(connectorAdapterMatrix.rows.find((row) => row.platform === 'zoom').adapter_mode, 'native_or_desktop_observer');
assert.equal(connectorAdapterMatrix.rows.every((row) => row.runtime_sequence[0].action === 'observe_platform_candidates'), true);
assert.equal(sdk.assertConnectorAdapterMatrix(connectorPackage).accepted, true);
const resolvedConnectorHostAdapter = sdk.resolveConnectorHostAdapterConfig({
  tabs: [
    { url: 'https://zoom.us/j/987654321', active: false, in_meeting: true },
    { url: 'https://meet.google.com/abc-defg-hij', active: true, audible: true, in_meeting: true },
  ],
}, connectorPackage);
assert.equal(resolvedConnectorHostAdapter.schema, 'meeting_app_timeline_host_adapter_config_resolution');
assert.equal(resolvedConnectorHostAdapter.accepted, true);
assert.equal(resolvedConnectorHostAdapter.platform, 'google_meet');
assert.equal(resolvedConnectorHostAdapter.host_config.selected_surface, 'browser_extension');
assert.equal(sdk.assertResolvedConnectorHostAdapterConfig('https://zoom.us/j/987654321', connectorPackage).platform, 'zoom');
const connectorHostAdapterBootstrapPlan = sdk.connectorHostAdapterBootstrapPlan({
  tab: { url: 'https://meet.google.com/abc-defg-hij', active: true, in_meeting: true },
}, connectorPackage);
assert.equal(connectorHostAdapterBootstrapPlan.schema, 'meeting_app_timeline_host_adapter_bootstrap_plan');
assert.equal(connectorHostAdapterBootstrapPlan.accepted, true);
assert.equal(connectorHostAdapterBootstrapPlan.startup_order[3], 'runtime_observe_platform_candidates');
assert.equal(connectorHostAdapterBootstrapPlan.required_runtime_actions.includes('insert_annotation'), true);
assert.equal(sdk.assertConnectorHostAdapterBootstrapPlan('https://zoom.us/j/987654321', connectorPackage).platform, 'zoom');
const connectorHostAdapterBootstrapPlanMatrix = sdk.connectorHostAdapterBootstrapPlanMatrix(connectorPackage);
assert.equal(connectorHostAdapterBootstrapPlanMatrix.schema, 'meeting_app_timeline_host_adapter_bootstrap_plan_matrix');
assert.equal(connectorHostAdapterBootstrapPlanMatrix.accepted, true);
assert.equal(connectorHostAdapterBootstrapPlanMatrix.accepted_count, 2);
const connectorHostAdapterBootstrapPlanMatrixAcceptance = sdk.connectorHostAdapterBootstrapPlanMatrixAcceptanceReport(connectorHostAdapterBootstrapPlanMatrix);
assert.equal(connectorHostAdapterBootstrapPlanMatrixAcceptance.schema, 'meeting_app_timeline_host_adapter_bootstrap_plan_matrix_acceptance_report');
assert.equal(connectorHostAdapterBootstrapPlanMatrixAcceptance.accepted, true);
assert.equal(connectorHostAdapterBootstrapPlanMatrixAcceptance.rows.every((row) => row.startup_order_ready === true), true);
assert.equal(sdk.assertConnectorHostAdapterBootstrapPlanMatrix(connectorPackage).accepted, true);
const unsupportedConnectorHostAdapter = sdk.resolveConnectorHostAdapterConfig(
  'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
  connectorPackage,
);
assert.equal(unsupportedConnectorHostAdapter.accepted, false);
assert.equal(unsupportedConnectorHostAdapter.issues.some((issue) => issue.code === 'meeting_platform_not_supported'), true);
const providerReplayReport = sdk.providerReplayReport('google-meet', {
  baseReceivedAtMs: 1_782_614_400_000,
});
assert.equal(providerReplayReport.schema, 'meeting_platform_provider_replay_report');
assert.equal(providerReplayReport.accepted, true);
assert.equal(providerReplayReport.coverage.meeting_start, true);
assert.equal(providerReplayReport.runtime_contract.provider_events_block_realtime, false);
assert.equal(sdk.assertProviderReplayReport(providerReplayReport).accepted, true);
const providerReplayMatrix = sdk.providerReplayMatrix({
  platforms: ['google-meet', 'zoom'],
});
assert.equal(providerReplayMatrix.schema, 'meeting_platform_provider_replay_matrix');
assert.equal(providerReplayMatrix.accepted, true);
assert.equal(providerReplayMatrix.accepted_count, 2);
assert.equal(providerReplayMatrix.rows.every((row) => row.provider_events_block_realtime === false), true);
assert.equal(sdk.assertProviderReplayMatrix(providerReplayMatrix).accepted, true);

const adapterCapability = sdk.meetingAppAdapterCapability('google-meet', {
  url: 'https://meet.google.com/abc-defg-hij',
  page: {
    controls: [{ label: 'Leave call' }],
    participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
  },
});
assert.equal(adapterCapability.schema, 'meeting_app_adapter_capability_report');
assert.equal(adapterCapability.platform, 'google_meet');
assert.equal(adapterCapability.accepted, true);
assert.equal(adapterCapability.timeline_capabilities.realtime_axis.status, 'local_ready');
assert.equal(adapterCapability.timeline_capabilities.speaker_track.status, 'local_ready');
assert.equal(adapterCapability.timeline_capabilities.participant_track.status, 'available');
assert.equal(adapterCapability.timeline_capabilities.post_meeting_transcript.provider_declared, true);
assert.equal(sdk.adapterCapability('zoom').platform, 'zoom');

const adapterCapabilityMatrix = sdk.meetingAppAdapterCapabilityMatrix();
assert.equal(adapterCapabilityMatrix.schema, 'meeting_app_adapter_capability_matrix');
assert.equal(adapterCapabilityMatrix.platform_count, 2);
assert.equal(adapterCapabilityMatrix.provider_axis_declared_count, 2);
assert.equal(sdk.adapterCapabilityMatrix().static_ready_count, 2);

const adapterExecutionPlan = sdk.meetingAppAdapterExecutionPlan(adapterCapability);
assert.equal(adapterExecutionPlan.schema, 'meeting_app_adapter_execution_plan');
assert.equal(adapterExecutionPlan.platform, 'google_meet');
assert.equal(adapterExecutionPlan.realtime_ready, true);
assert.equal(adapterExecutionPlan.steps.some((step) => step.id === 'configure_provider_reconcile' && step.blocks_realtime_if_missing === false), true);
assert.equal(sdk.adapterExecutionPlan('zoom').platform, 'zoom');
assert.equal(sdk.meetingAppAdapterExecutionPlanMatrix().platform_count, 2);
assert.equal(sdk.adapterExecutionPlanMatrix().accepted_count, 2);

const adaptationPackage = sdk.platformAdaptationPackage('google-meet');
assert.equal(adaptationPackage.schema, 'meeting_platform_adaptation_package');
assert.equal(adaptationPackage.platform, 'google_meet');
assert.equal(adaptationPackage.adaptation_playbook.integration_path.path, 'google_workspace_events_pubsub');
assert.equal(sdk.adaptationPackage('zoom').platform, 'zoom');

const adaptationPackageMatrix = sdk.platformAdaptationPackageMatrix();
assert.equal(adaptationPackageMatrix.platform_count, 2);
assert.equal(adaptationPackageMatrix.rows.find((row) => row.platform === 'google_meet').provider_path, 'google_workspace_events_pubsub');
assert.equal(sdk.adaptationPackageMatrix().candidate_observer_count, 2);

const consumerHandoff = sdk.platformConsumerHandoff();
assert.equal(consumerHandoff.schema, 'meeting_platform_consumer_handoff');
assert.equal(consumerHandoff.platform_count, 2);
assert.equal(consumerHandoff.hard_contracts.timestamp_field, 'captured_at_ms');
assert.equal(sdk.consumerHandoff().lightweight_connector_ready, true);
assert.equal(sdk.assertConsumerHandoff().accepted, true);

const runtimeBundle = sdk.platformRuntimeBundle('google-meet');
assert.equal(runtimeBundle.schema, 'meeting_platform_runtime_bundle');
assert.equal(runtimeBundle.runtime.lightweight_connector_bridge.install_function, 'installMeetingPlatformConnectorContentScriptBridge');
assert.equal(runtimeBundle.messaging.runtime_bridge_bootstrap.target_message_type, 'meeting_timeline.runtime_target');
assert.equal(sdk.runtimeBundle('zoom').platform, 'zoom');
assert.equal(sdk.platformRuntimeBundleMatrix().platform_count, 2);
assert.equal(sdk.platformRuntimeBundleMatrix().runtime_target_message_count, 2);
assert.equal(sdk.runtimeBundleMatrix().lightweight_connector_ready_count, 2);

const adapterRoute = sdk.platformAdapterRoute('google-meet');
assert.equal(adapterRoute.schema, 'meeting_platform_adapter_route');
assert.equal(adapterRoute.platform, 'google_meet');
assert.equal(sdk.adapterRoute('zoom').platform, 'zoom');
assert.equal(sdk.platformAdapterRouteMatrix().platform_count, 2);

const adapterDecision = sdk.platformAdapterDecision({
  url: 'https://meet.google.com/abc-defg-hij',
});
assert.equal(adapterDecision.schema, 'meeting_platform_adapter_decision');
assert.equal(adapterDecision.platform, 'google_meet');
assert.equal(adapterDecision.selected_surface, 'browser_extension');
assert.equal(adapterDecision.contracts.provider_events_block_realtime, false);
assert.equal(sdk.adapterDecision({
  url: 'https://zoom.us/j/987654321',
}).platform, 'zoom');
assert.equal(sdk.platformAdapterDecisionMatrix({}, {
  platforms: ['google-meet', 'zoom'],
}).accepted_count, 2);

const adapterStartupPlan = sdk.platformAdapterStartupPlan({
  url: 'https://meet.google.com/abc-defg-hij',
});
assert.equal(adapterStartupPlan.schema, 'meeting_platform_adapter_startup_plan');
assert.equal(adapterStartupPlan.platform, 'google_meet');
assert.equal(adapterStartupPlan.selected_surface, 'browser_extension');
assert.equal(adapterStartupPlan.realtime_startup_ready, true);
assert.equal(adapterStartupPlan.message_contract.insert_annotation, 'meeting_timeline.insert_mark');
assert.equal(sdk.adapterStartupPlan({
  url: 'https://zoom.us/j/987654321',
}).platform, 'zoom');
assert.equal(sdk.platformAdapterStartupPlanMatrix({}, {
  platforms: ['google-meet', 'zoom'],
}).realtime_startup_ready_count, 2);

const adaptationStrategy = sdk.platformAdaptationStrategy('google-meet');
assert.equal(adaptationStrategy.schema, 'meeting_platform_adaptation_strategy');
assert.equal(adaptationStrategy.adaptation_playbook.integration_path.path, 'google_workspace_events_pubsub');
assert.equal(sdk.adaptationStrategy('zoom').platform, 'zoom');
assert.equal(sdk.platformAdaptationStrategyMatrix().provider_reconcile_required_count, 2);

const connectorHub = sdk.platformConnectorHub();
assert.equal(connectorHub.schema, 'meeting_platform_connector_hub');
assert.equal(connectorHub.accepted, true);
assert.equal(sdk.connectorHub().platform_count, 2);

const allPlatformsHostPackage = createMeetingAppTimelineSdk({
  baseUrl,
  fetch: fetchImpl,
}).hostPackage({
  surfaces: ['browser-extension'],
});
assert.deepEqual(allPlatformsHostPackage.platforms, ['google_meet', 'microsoft_teams', 'zoom', 'webex', 'lark']);
assert.equal(allPlatformsHostPackage.handoff_count, 5);
assert.equal(allPlatformsHostPackage.ready_count, 5);

const remoteObserved = await sdk.observeMeetingApp({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Google Meet',
  observed_at_ms: 1_782_614_401_000,
}, {
  remote: true,
});
assert.equal(remoteObserved.ok, true);
assert.equal(calls.at(-1).url, `${baseUrl}/api/meeting-platform/runtime-events`);
assert.equal(calls.at(-1).body.action, 'observe_meeting_app');
assert.equal(calls.at(-1).body.platform, 'google_meet');

const remoteInserted = await sdk.insertAnnotation('zoom', {
  id: 'mark-1',
  captured_at_ms: 1_782_614_402_000,
  text: 'why?',
}, {
  remote: true,
});
assert.equal(remoteInserted.ok, true);
assert.equal(calls.at(-1).body.action, 'insert_annotation');
assert.equal(calls.at(-1).body.platform, 'zoom');
assert.equal(calls.at(-1).body.annotation.text, 'why?');
assert.equal(calls.at(-1).body.payload.text, 'why?');

const candidateObserved = await sdk.observePlatformCandidates({
  tabs: [{
    active: true,
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Google Meet',
  }],
}, {
  mode: 'runtime_event',
});
assert.equal(candidateObserved.ok, true);
assert.equal(calls.at(-1).body.action, 'observe_platform_candidates');

console.log('ok meeting app timeline SDK facade');
