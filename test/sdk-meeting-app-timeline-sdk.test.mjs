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
