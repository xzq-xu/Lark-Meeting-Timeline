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

const appFixtures = kit.allMeetingAppFixtures();
assert.equal(Object.keys(appFixtures).length, 5);
assert.equal(appFixtures.lark.platform, 'lark');

const appFixtureAcceptance = kit.meetingAppFixtureAcceptance();
assert.equal(appFixtureAcceptance.accepted, true);
assert.equal(appFixtureAcceptance.accepted_count, 5);
assert.equal(appFixtureAcceptance.coverage_by_platform.google_meet.meeting_started, true);

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
assert.equal(report.meeting_app_fixture_acceptance.accepted, true);
assert.equal(report.meeting_app_fixture_acceptance.accepted_count, 5);

console.log('ok meeting platform timeline kit');
