import assert from 'node:assert/strict';

import {
  MEETING_PLATFORM_ADAPTER_SAMPLE_PLATFORMS,
  assertMeetingPlatformAdapterSample,
  assertMeetingPlatformAdapterSampleMatrix,
  buildMeetingPlatformAdapterSamplePlan,
  runMeetingPlatformAdapterSample,
  runMeetingPlatformAdapterSampleMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-sample.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';

assert.deepEqual(MEETING_PLATFORM_ADAPTER_SAMPLE_PLATFORMS, [
  'google_meet',
  'microsoft_teams',
  'zoom',
  'webex',
  'lark',
]);

const googlePlan = buildMeetingPlatformAdapterSamplePlan('google-meet', { baseUrl });
assert.equal(googlePlan.schema, 'meeting_platform_adapter_sample_plan');
assert.equal(googlePlan.platform, 'google_meet');
assert.equal(googlePlan.endpoint, `${baseUrl}/api/platform-events/google-meet`);
assert.equal(googlePlan.provider_signal_types.includes('meeting_start'), true);
assert.equal(googlePlan.provider_signal_types.includes('meeting_end'), true);
assert.equal(googlePlan.annotation.capturedAtMs, googlePlan.clock.annotationMs);
assert.equal(googlePlan.contract_acceptance.accepted, true);

const googleSample = await runMeetingPlatformAdapterSample('google-meet', {
  baseUrl,
  includeEvidencePackage: true,
});
assert.equal(googleSample.schema, 'meeting_platform_adapter_sample');
assert.equal(googleSample.platform, 'google_meet');
assert.equal(googleSample.accepted, true);
assert.equal(googleSample.readiness.status, 'ready');
assert.equal(googleSample.verification.passed, true);
assert.equal(googleSample.summary.production_ready, true);
assert.equal(googleSample.timeline_method_counts.insertMark, 1);
assert.equal(googleSample.timeline_method_counts.startMeeting > 0, true);
assert.equal(googleSample.timeline_method_counts.endMeeting > 0, true);
assert.equal(googleSample.provider_event_count, 5);
assert.equal(googleSample.evidence_package.schema, 'meeting_platform_evidence_package');

const teamsSample = await runMeetingPlatformAdapterSample('teams', {
  baseUrl,
  providerSignalTypes: ['meeting_start', 'meeting_end'],
});
assert.equal(teamsSample.platform, 'microsoft_teams');
assert.equal(teamsSample.accepted, true);
assert.equal(teamsSample.provider_event_count, 2);

const sampleMatrix = await runMeetingPlatformAdapterSampleMatrix({
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(sampleMatrix.schema, 'meeting_platform_adapter_sample_matrix');
assert.equal(sampleMatrix.platform_count, 5);
assert.equal(sampleMatrix.accepted_count, 5);
assert.equal(sampleMatrix.rejected_count, 0);
assert.equal(sampleMatrix.rows.every((row) => row.production_ready === true), true);
assert.deepEqual(sampleMatrix.platforms, ['google_meet', 'microsoft_teams', 'zoom', 'webex', 'lark']);

assert.equal((await assertMeetingPlatformAdapterSample('zoom', { baseUrl })).accepted, true);
assert.equal((await assertMeetingPlatformAdapterSampleMatrix({
  baseUrl,
  platforms: ['webex', 'lark'],
})).accepted_count, 2);

const client = {
  async startMeeting(input) { return { ok: true, input }; },
  async endMeeting(input) { return { ok: true, input }; },
  async insertMark(input) { return { ok: true, input }; },
  async insertMarks(input) { return { ok: true, input }; },
};
const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl,
  verify: false,
});

assert.equal(kit.platformAdapterSamplePlan('google-meet').platform, 'google_meet');
assert.equal((await kit.runPlatformAdapterSample('google-meet')).accepted, true);
assert.equal((await kit.runPlatformAdapterSampleMatrix({
  platforms: ['zoom'],
})).accepted_count, 1);
assert.equal((await kit.assertPlatformAdapterSample('webex')).accepted, true);
assert.equal((await kit.assertPlatformAdapterSampleMatrix({
  platforms: ['lark'],
})).accepted_count, 1);

console.log('ok meeting platform adapter sample');
