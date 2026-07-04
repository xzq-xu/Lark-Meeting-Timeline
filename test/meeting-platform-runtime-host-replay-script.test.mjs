import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

import {
  buildMeetingAppFixtureSnapshot,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import {
  buildMeetingAppSnapshotRecordSet,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-runtime-host-replay-script-'));
const inputFile = join(tmpDir, 'google-meet-records.json');
const reportFile = join(tmpDir, 'runtime-host-replay-report.json');
const startMs = 1_783_356_000_000;

const recordSet = buildMeetingAppSnapshotRecordSet([
  {
    phase: 'active',
    captured_at_ms: startMs,
    snapshot: buildMeetingAppFixtureSnapshot('google-meet', {
      state: 'active',
      observedAtMs: startMs,
    }),
  },
  {
    phase: 'ended',
    captured_at_ms: startMs + 2_000,
    snapshot: buildMeetingAppFixtureSnapshot('google-meet', {
      state: 'prejoin',
      observedAtMs: startMs + 2_000,
    }),
  },
], {
  createdAtMs: startMs,
});
await writeFile(inputFile, `${JSON.stringify(recordSet, null, 2)}\n`, 'utf8');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-runtime-host-replay.mjs',
  `--input=${inputFile}`,
  '--platforms=google-meet',
  `--report-file=${reportFile}`,
], {
  cwd: repoRoot,
});

assert.match(stdout, /meeting_platform_runtime_host_replay_report/);
assert.match(stdout, /accepted=1\/1/);
assert.match(stdout, /google_meet: accepted=yes/);

const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.type, 'meeting_platform_runtime_host_replay_report');
assert.equal(report.ok, true);
assert.equal(report.input_file_count, 1);
assert.equal(report.platform_count, 1);
assert.equal(report.accepted_count, 1);
assert.equal(report.rows[0].actions.includes('endMeeting'), true);
assert.equal(report.rows[0].signal_types.includes('speaker_started'), true);

const { stdout: jsonStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-runtime-host-replay.mjs',
  `--input=${inputFile}`,
  '--platforms=google-meet',
  '--include-reports=true',
  '--json=true',
], {
  cwd: repoRoot,
  maxBuffer: 20 * 1024 * 1024,
});
const jsonReport = JSON.parse(jsonStdout);
assert.equal(jsonReport.ok, true);
assert.equal(jsonReport.matrix.reports.length, 1);
assert.equal(jsonReport.matrix.reports[0].coverage.end_timestamp_aligned, true);

const { stdout: missingStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-runtime-host-replay.mjs',
  '--platforms=zoom',
], {
  cwd: repoRoot,
});
assert.match(missingStdout, /zoom: accepted=no/);
assert.match(missingStdout, /active_snapshot_available/);

console.log('ok meeting platform runtime host replay script');
