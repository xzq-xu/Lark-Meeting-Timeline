import assert from 'node:assert/strict';

import {
  assertMeetingPlatformRuntimeHostReplay,
  assertMeetingPlatformRuntimeHostReplayMatrix,
  runMeetingPlatformRuntimeHostReplay,
  runMeetingPlatformRuntimeHostReplayMatrix,
} from '../packages/meeting-timeline-sdk/adapters/meeting-platform-runtime-host-verifier.mjs';
import {
  buildMeetingAppFixtureSnapshot,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import {
  buildMeetingAppSnapshotRecordSet,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';
import {
  buildMeetingPlatformEvidencePackage,
} from '../packages/meeting-timeline-sdk/adapters/platform-evidence-package.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const startMs = 1_783_356_000_000;

function replayRecords(platform, offset = 0) {
  const activeAtMs = startMs + offset;
  const endAtMs = activeAtMs + 2_000;
  return [
    {
      phase: 'active',
      captured_at_ms: activeAtMs,
      snapshot: buildMeetingAppFixtureSnapshot(platform, {
        state: 'active',
        observedAtMs: activeAtMs,
      }),
    },
    {
      phase: 'ended',
      captured_at_ms: endAtMs,
      snapshot: buildMeetingAppFixtureSnapshot(platform, {
        state: 'prejoin',
        observedAtMs: endAtMs,
      }),
    },
  ];
}

const recordSet = buildMeetingAppSnapshotRecordSet(replayRecords('google-meet'), {
  createdAtMs: startMs,
});
const report = await runMeetingPlatformRuntimeHostReplay('google-meet', recordSet);
assert.equal(report.schema, 'meeting_platform_runtime_host_replay');
assert.equal(report.platform, 'google_meet');
assert.equal(report.accepted, true);
assert.equal(report.input_record_count, 2);
assert.equal(report.replay_row_count, 2);
assert.equal(report.active_at_ms, startMs);
assert.equal(report.end_at_ms, startMs + 2_000);
assert.deepEqual(report.actions, ['startMeeting', 'insertMark', 'endMeeting']);
assert.deepEqual(report.signal_types, ['meeting_started', 'speaker_started', 'meeting_ended']);
assert.equal(report.coverage.active_snapshot_available, true);
assert.equal(report.coverage.end_snapshot_available, true);
assert.equal(report.coverage.start_timestamp_aligned, true);
assert.equal(report.coverage.end_timestamp_aligned, true);
assert.equal(report.active_row.has_active_speaker, true);
assertMeetingPlatformRuntimeHostReplay(report);

const evidencePackage = buildMeetingPlatformEvidencePackage('zoom', {
  meetingAppRecords: replayRecords('zoom', 10_000),
}, {
  createdAtMs: startMs,
});
const evidenceReplay = await runMeetingPlatformRuntimeHostReplay('zoom', evidencePackage);
assert.equal(evidenceReplay.accepted, true);
assert.equal(evidenceReplay.platform, 'zoom');
assert.equal(evidenceReplay.active_at_ms, startMs + 10_000);
assert.equal(evidenceReplay.actions.includes('endMeeting'), true);

const matrix = await runMeetingPlatformRuntimeHostReplayMatrix({
  platforms: ['google-meet', 'zoom'],
  inputs: {
    google_meet: recordSet,
    zoom: evidencePackage,
  },
});
assert.equal(matrix.schema, 'meeting_platform_runtime_host_replay_matrix');
assert.equal(matrix.platform_count, 2);
assert.equal(matrix.accepted_count, 2);
assert.deepEqual(matrix.platforms, ['google_meet', 'zoom']);
assert.equal(matrix.rows.every((row) => row.accepted === true), true);
assertMeetingPlatformRuntimeHostReplayMatrix(matrix);

const missingEnd = await runMeetingPlatformRuntimeHostReplay('lark', {
  records: replayRecords('lark').slice(0, 1),
});
assert.equal(missingEnd.accepted, false);
assert.equal(missingEnd.missing.includes('end_snapshot_available'), true);
assert.equal(missingEnd.actions.length, 0);

const kit = createMeetingPlatformTimelineKit({
  async startMeeting(input) { return { ok: true, input }; },
  async endMeeting(input) { return { ok: true, input }; },
  async insertMark(input) { return { ok: true, input }; },
  async insertMarks(input) { return { ok: true, input }; },
}, {
  baseUrl: 'https://timeline.example.com',
});
const kitReplay = await kit.replayPlatformRuntimeHost('webex', {
  records: replayRecords('webex', 20_000),
});
assert.equal(kitReplay.accepted, true);
assert.equal(kitReplay.platform, 'webex');
const kitReplayMatrix = await kit.replayPlatformRuntimeHostMatrix({
  platforms: ['google-meet'],
  input: recordSet,
});
assert.equal(kitReplayMatrix.accepted_count, 1);

await assert.rejects(
  () => runMeetingPlatformRuntimeHostReplay('local-detector', { records: [] }),
  /Unsupported runtime host replay platform/,
);

assert.throws(() => assertMeetingPlatformRuntimeHostReplay({
  platform: 'google_meet',
  accepted: false,
  missing: ['end_snapshot_available'],
}), /replay failed/);

console.log('ok meeting platform runtime host replay');
