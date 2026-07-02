import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

import {
  MEETING_PLATFORM_ALIASES,
  MEETING_PLATFORM_KEYS,
  allPlatformCapabilityContracts,
  allPlatformSetupManifests,
  normalizeMeetingPlatform,
  platformCapabilityContract,
  platformEventEndpoint,
  platformSetupManifest,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';
import * as transcriptAdapters from '../packages/meeting-timeline-sdk/adapters/transcript.mjs';

const baseUrl = 'https://timeline.example.com';
const expectedPlatforms = ['google_meet', 'microsoft_teams', 'zoom', 'webex'];
const eventAdapterModules = {
  google_meet: '../packages/meeting-timeline-sdk/adapters/google-meet.mjs',
  microsoft_teams: '../packages/meeting-timeline-sdk/adapters/microsoft-teams.mjs',
  zoom: '../packages/meeting-timeline-sdk/adapters/zoom.mjs',
  webex: '../packages/meeting-timeline-sdk/adapters/webex.mjs',
};
const eventAdapterExports = {
  google_meet: 'normalizeGoogleMeetEvent',
  microsoft_teams: 'normalizeMicrosoftTeamsEvent',
  zoom: 'normalizeZoomEvent',
  webex: 'normalizeWebexEvent',
};

assert.deepEqual([...MEETING_PLATFORM_KEYS], expectedPlatforms);
assert.equal(MEETING_PLATFORM_ALIASES['google-meet'], 'google_meet');
assert.equal(MEETING_PLATFORM_ALIASES['microsoft-teams'], 'microsoft_teams');
assert.equal(normalizeMeetingPlatform('meet'), 'google_meet');
assert.equal(normalizeMeetingPlatform('teams'), 'microsoft_teams');
assert.equal(normalizeMeetingPlatform('cisco-webex'), 'webex');
assert.throws(() => normalizeMeetingPlatform('unknown-meeting-platform'), /Unsupported meeting platform/);

const setupRows = allPlatformSetupManifests({ baseUrl });
const capabilityRows = allPlatformCapabilityContracts({ baseUrl });
assert.deepEqual(setupRows.map((item) => item.platform), expectedPlatforms);
assert.deepEqual(capabilityRows.map((item) => item.platform), expectedPlatforms);

const envExample = readFileSync(new URL('../.env.example', import.meta.url), 'utf8');

for (const platform of MEETING_PLATFORM_KEYS) {
  const manifest = platformSetupManifest(platform, { baseUrl });
  const capability = platformCapabilityContract(platform, { baseUrl });
  assert.equal(manifest.platform, platform);
  assert.equal(manifest.capabilities.platform, platform);
  assert.equal(capability.platform, platform);
  assert.equal(capability.endpoints.platform_events, platformEventEndpoint(baseUrl, platform));
  assert.equal(capability.endpoints.transcript_import, `${baseUrl}/api/import/transcript`);
  assert.equal(typeof capability.sdk_modules.events, 'string');
  assert.equal(typeof capability.sdk_modules.transcript, 'string');
  assert.equal(typeof capability.sdk_modules.setup, 'string');
  assert.equal(typeof capability.sdk_modules.security, 'string');
  assert.equal(Array.isArray(capability.realtime_axis.signal_types), true);
  assert.equal(capability.realtime_axis.signal_types.includes('meeting_started'), true);
  assert.equal(capability.realtime_axis.signal_types.includes('meeting_ended'), true);

  for (const envName of manifest.required_security_env ?? []) {
    assert.match(envExample, new RegExp(`^${envName}=`, 'm'), `${platform} missing ${envName} in .env.example`);
  }

  const transcriptNormalizer = capability.post_meeting_transcript?.sdk_normalizer;
  if (transcriptNormalizer) {
    assert.equal(typeof transcriptAdapters[transcriptNormalizer], 'function', `${platform} transcript normalizer missing`);
  }

  const eventModule = await import(eventAdapterModules[platform]);
  assert.equal(typeof eventModule[eventAdapterExports[platform]], 'function', `${platform} event normalizer missing`);
}

console.log('ok meeting platform SDK conformance');
