import assert from 'node:assert/strict';

import {
  assertMeetingPlatformRuntimeHostVerification,
  assertMeetingPlatformRuntimeHostVerificationMatrix,
  createMeetingPlatformRuntimeHostFixtureEnvironment,
  createMeetingPlatformRuntimeHostVerificationClient,
  runMeetingPlatformRuntimeHostVerification,
  runMeetingPlatformRuntimeHostVerificationMatrix,
} from '../packages/meeting-timeline-sdk/adapters/meeting-platform-runtime-host-verifier.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const startMs = 1_783_356_000_000;
const report = await runMeetingPlatformRuntimeHostVerification('google-meet', {
  startMs,
  endOffsetMs: 2_000,
});
assert.equal(report.schema, 'meeting_platform_runtime_host_verification');
assert.equal(report.platform, 'google_meet');
assert.equal(report.accepted, true);
assert.equal(report.start_at_ms, startMs);
assert.equal(report.end_at_ms, startMs + 2_000);
assert.deepEqual(report.actions, ['startMeeting', 'insertMark', 'endMeeting']);
assert.deepEqual(report.signal_types, ['meeting_started', 'speaker_started', 'meeting_ended']);
assert.deepEqual(report.result_actions, ['startMeeting', 'insertSpeakerMark', 'endMeeting']);
assert.equal(report.coverage.host_ready, true);
assert.equal(report.coverage.runtime_capture_profile, true);
assert.equal(report.coverage.meeting_start_signal, true);
assert.equal(report.coverage.speaker_start_signal, true);
assert.equal(report.coverage.meeting_end_signal, true);
assert.equal(report.coverage.start_meeting_written, true);
assert.equal(report.coverage.speaker_mark_written, true);
assert.equal(report.coverage.end_meeting_written, true);
assert.equal(report.coverage.start_timestamp_aligned, true);
assert.equal(report.coverage.speaker_timestamp_aligned, true);
assert.equal(report.coverage.end_timestamp_aligned, true);
assert.equal(report.coverage.participant_capture, true);
assert.equal(report.coverage.control_capture, true);
assert.equal(report.capture.start.profile, 'google_meet');
assert.equal(report.capture.end.profile, 'google_meet');
assert.equal(report.capture.start.participant_count > 0, true);
assert.equal(report.calls.find((call) => call.action === 'insertMark').input.kind, 'speaker_started');
assertMeetingPlatformRuntimeHostVerification(report);

const matrix = await runMeetingPlatformRuntimeHostVerificationMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  startMs,
});
assert.equal(matrix.schema, 'meeting_platform_runtime_host_verification_matrix');
assert.equal(matrix.platform_count, 5);
assert.equal(matrix.accepted_count, 5);
assert.deepEqual(matrix.platforms, ['google_meet', 'microsoft_teams', 'zoom', 'webex', 'lark']);
assert.equal(matrix.rows.every((row) => row.accepted === true), true);
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').actions.includes('endMeeting'), true);
assertMeetingPlatformRuntimeHostVerificationMatrix(matrix);

const env = createMeetingPlatformRuntimeHostFixtureEnvironment('zoom', { startMs });
assert.equal(env.platform, 'zoom');
assert.equal(env.snapshot.fixture_state, 'active');
env.setEnded({ observedAtMs: startMs + 1_000 });
assert.equal(env.snapshot.fixture_state, 'prejoin');
assert.equal(env.document.querySelectorAll('button').length > 0, true);

const client = createMeetingPlatformRuntimeHostVerificationClient();
await client.startMeeting({ meeting_id: 'm1' });
await client.insertMark({ kind: 'speaker_started' });
assert.deepEqual(client.getState().actions, ['startMeeting', 'insertMark']);

const kit = createMeetingPlatformTimelineKit({
  async startMeeting(input) { return { ok: true, input }; },
  async endMeeting(input) { return { ok: true, input }; },
  async insertMark(input) { return { ok: true, input }; },
}, {
  baseUrl: 'https://timeline.example.com',
});
const kitVerification = await kit.verifyPlatformRuntimeHost('lark', { startMs });
assert.equal(kitVerification.accepted, true);
assert.equal(kitVerification.platform, 'lark');
const kitMatrix = await kit.verifyPlatformRuntimeHostMatrix({ platforms: ['google-meet', 'zoom'], startMs });
assert.equal(kitMatrix.accepted_count, 2);
assert.equal(kit.createPlatformRuntimeHostVerificationClient().getState().call_count, 0);
assert.equal(kit.createPlatformRuntimeHostFixtureEnvironment('webex').platform, 'webex');

await assert.rejects(
  () => runMeetingPlatformRuntimeHostVerification('local-detector'),
  /Unsupported runtime host verification platform/,
);

assert.throws(() => assertMeetingPlatformRuntimeHostVerification({
  platform: 'google_meet',
  accepted: false,
  missing: ['meeting_end_signal'],
}), /verification failed/);

console.log('ok meeting platform runtime host verifier');
