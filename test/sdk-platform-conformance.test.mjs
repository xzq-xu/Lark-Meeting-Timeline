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
import {
  MEETING_PLATFORM_EVENT_ADAPTERS,
  meetingPlatformEventAdapterFor,
} from '../packages/meeting-timeline-sdk/adapters/platform-registry.mjs';
import * as transcriptAdapters from '../packages/meeting-timeline-sdk/adapters/transcript.mjs';

const baseUrl = 'https://timeline.example.com';
const expectedPlatforms = ['local_detector', 'lark', 'google_meet', 'microsoft_teams', 'zoom', 'webex'];

assert.deepEqual([...MEETING_PLATFORM_KEYS], expectedPlatforms);
assert.equal(MEETING_PLATFORM_ALIASES['local-detector'], 'local_detector');
assert.equal(MEETING_PLATFORM_ALIASES.feishu, 'lark');
assert.equal(MEETING_PLATFORM_ALIASES['google-meet'], 'google_meet');
assert.equal(MEETING_PLATFORM_ALIASES['microsoft-teams'], 'microsoft_teams');
assert.equal(normalizeMeetingPlatform('meet'), 'google_meet');
assert.equal(normalizeMeetingPlatform('desktop-observer'), 'local_detector');
assert.equal(normalizeMeetingPlatform('teams'), 'microsoft_teams');
assert.equal(normalizeMeetingPlatform('cisco-webex'), 'webex');
assert.equal(normalizeMeetingPlatform('lark-suite'), 'lark');
assert.throws(() => normalizeMeetingPlatform('unknown-meeting-platform'), /Unsupported meeting platform/);
assert.equal(meetingPlatformEventAdapterFor('unknown-meeting-platform'), null);

const setupRows = allPlatformSetupManifests({ baseUrl });
const capabilityRows = allPlatformCapabilityContracts({ baseUrl });
assert.deepEqual(setupRows.map((item) => item.platform), expectedPlatforms);
assert.deepEqual(capabilityRows.map((item) => item.platform), expectedPlatforms);
assert.deepEqual(MEETING_PLATFORM_EVENT_ADAPTERS.map((item) => item.key), expectedPlatforms);

const envExample = readFileSync(new URL('../.env.example', import.meta.url), 'utf8');

for (const platform of MEETING_PLATFORM_KEYS) {
  const eventAdapter = meetingPlatformEventAdapterFor(platform);
  const manifest = platformSetupManifest(platform, { baseUrl });
  const capability = platformCapabilityContract(platform, { baseUrl });
  assert.equal(eventAdapter.key, platform);
  assert.equal(eventAdapter.source, platform === 'local_detector' ? 'local_detector' : `${platform}_webhook`);
  assert.equal(typeof eventAdapter.normalize, 'function');
  assert.equal(eventAdapter.aliases.includes(platform), true);
  for (const alias of eventAdapter.aliases) {
    assert.equal(meetingPlatformEventAdapterFor(alias).key, platform);
  }
  assert.equal(manifest.platform, platform);
  assert.equal(manifest.capabilities.platform, platform);
  assert.equal(capability.platform, platform);
  assert.equal(capability.endpoints.platform_events, platformEventEndpoint(baseUrl, platform));
  assert.equal(capability.endpoints.transcript_import, `${baseUrl}/api/import/transcript`);
  assert.equal(typeof capability.sdk_modules.events, 'string');
  assert.equal(capability.sdk_modules.ingest, '@ai-annotation/meeting-timeline-sdk/adapters/platform-ingest');
  assert.equal(capability.sdk_modules.webhook_handler, '@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-handler');
  assert.equal(typeof capability.sdk_modules.transcript, 'string');
  assert.equal(typeof capability.sdk_modules.setup, 'string');
  assert.equal(typeof capability.sdk_modules.security, 'string');
  assert.equal(Array.isArray(capability.realtime_axis.signal_types), true);
  assert.equal(capability.realtime_axis.signal_types.includes('meeting_started'), true);
  assert.equal(capability.realtime_axis.signal_types.includes('meeting_ended'), true);
  assert.equal(Array.isArray(capability.speaker_activity.signal_types), true);
  assert.equal(capability.speaker_activity.signal_types.includes('speaker_started'), true);
  assert.equal(capability.speaker_activity.signal_types.includes('speaker_ended'), true);

  for (const envName of manifest.required_security_env ?? []) {
    assert.match(envExample, new RegExp(`^${envName}=`, 'm'), `${platform} missing ${envName} in .env.example`);
  }

  const transcriptNormalizer = capability.post_meeting_transcript?.sdk_normalizer;
  if (transcriptNormalizer) {
    assert.equal(typeof transcriptAdapters[transcriptNormalizer], 'function', `${platform} transcript normalizer missing`);
  }
}

console.log('ok meeting platform SDK conformance');
