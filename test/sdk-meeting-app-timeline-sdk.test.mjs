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
