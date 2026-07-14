import assert from 'node:assert/strict';

import { buildPlatformAcceptanceReport } from '../packages/meeting-timeline-sdk/adapters/platform-acceptance.mjs';
import {
  DEFAULT_PLATFORM_FIXTURE_SIGNAL_TYPES,
  PLATFORM_FIXTURE_SIGNAL_TYPES,
  buildAllPlatformFixtureSamples,
  buildPlatformFixtureAcceptanceInput,
  buildPlatformFixtureEnv,
  buildPlatformFixtureEvent,
  buildPlatformFixtureSamples,
} from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';
import { diagnosePlatformEvent } from '../packages/meeting-timeline-sdk/adapters/platform-ingest.mjs';
import {
  MEETING_PLATFORM_KEYS,
  platformCapabilityContract,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';

const baseUrl = 'https://timeline.example.com';

assert.equal(PLATFORM_FIXTURE_SIGNAL_TYPES.includes('meeting_start'), true);
assert.equal(DEFAULT_PLATFORM_FIXTURE_SIGNAL_TYPES.google_meet.includes('subscription_lifecycle'), true);
assert.equal(
  platformCapabilityContract('google-meet', { baseUrl }).sdk_modules.fixtures,
  '@ai-annotation/meeting-timeline-sdk/adapters/platform-fixtures',
);

const fixtureEnv = buildPlatformFixtureEnv({ baseUrl });
assert.equal(fixtureEnv.GOOGLE_PUBSUB_OIDC_AUDIENCE, 'https://timeline.example.com/api/platform-events/google-meet');
assert.equal(typeof fixtureEnv.MICROSOFT_GRAPH_CLIENT_STATE, 'string');
assert.equal(typeof fixtureEnv.ZOOM_WEBHOOK_SECRET_TOKEN, 'string');
assert.equal(typeof fixtureEnv.WEBEX_WEBHOOK_SECRET, 'string');

const googleStart = buildPlatformFixtureEvent('google-meet', 'meeting_start', { baseUrl });
assert.equal(googleStart.type, 'google.workspace.meet.conference.v2.started');
assert.equal(googleStart.data.meetingUri, 'https://meet.google.com/abc-defg-hij');

const zoomTranscript = buildPlatformFixtureEvent('zoom', 'transcript_ready', { baseUrl });
assert.equal(zoomTranscript.event, 'recording.transcript_completed');
assert.equal(zoomTranscript.payload.object.transcript.download_url, 'https://zoom.example/transcript.vtt');

const localSpeaker = diagnosePlatformEvent(
  'local-detector',
  buildPlatformFixtureEvent('local_detector', 'speaker_activity', { baseUrl }),
);
assert.equal(localSpeaker.coverage.speaker_activity, true);
assert.equal(localSpeaker.signal_types.includes('speaker_started'), true);

const fixtureInput = buildPlatformFixtureAcceptanceInput({ baseUrl });
assert.deepEqual(Object.keys(fixtureInput.samples), MEETING_PLATFORM_KEYS);
assert.equal(fixtureInput.baseUrl, baseUrl);
assert.equal(fixtureInput.env.GOOGLE_PUBSUB_OIDC_AUDIENCE, 'https://timeline.example.com/api/platform-events/google-meet');

for (const platform of MEETING_PLATFORM_KEYS) {
  const samples = buildPlatformFixtureSamples(platform, { baseUrl });
  assert.equal(samples.length, DEFAULT_PLATFORM_FIXTURE_SIGNAL_TYPES[platform].length);
  assert.equal(samples.every((item) => item.platform === platform), true);

  const report = buildPlatformAcceptanceReport(platform, {
    ...fixtureInput,
    requireEndEvent: true,
    requiredCoverage: platform === 'local_detector'
      ? ['meeting_start', 'meeting_end', 'participant_track', 'speaker_activity']
      : ['meeting_start', 'meeting_end', 'participant_track', 'artifact_ready'],
  });
  assert.equal(report.accepted, true, `${platform} fixture report should be accepted`);
  assert.equal(report.coverage.meeting_start, true, `${platform} should cover meeting start`);
  assert.equal(report.coverage.meeting_end, true, `${platform} should cover meeting end`);
  assert.equal(report.coverage.participant_track, true, `${platform} should cover participant track`);
  if (platform === 'local_detector') {
    assert.equal(report.coverage.speaker_activity, true, 'local detector should cover active speaker fixtures');
  } else {
    assert.equal(report.coverage.artifact_ready, true, `${platform} should cover artifact fixtures`);
  }
}

const allSamples = buildAllPlatformFixtureSamples({ baseUrl });
assert.equal(Object.keys(allSamples).length, MEETING_PLATFORM_KEYS.length);
assert.equal(allSamples.google_meet.some((item) => item.label === 'google_meet:meeting_start'), true);

assert.throws(
  () => buildPlatformFixtureEvent('google-meet', 'speaker_activity'),
  /not supported/,
);

console.log('ok meeting platform fixture samples');
