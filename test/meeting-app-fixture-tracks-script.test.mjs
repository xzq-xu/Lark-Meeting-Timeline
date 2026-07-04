import assert from 'node:assert/strict';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';

const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-app-fixture-tracks-'));
const reportFile = join(tmpDir, 'report.json');

const result = spawnSync(process.execPath, [
  'scripts/meeting-app-fixture-tracks.mjs',
  '--platforms=google-meet,teams,zoom,webex,lark',
  `--report-file=${reportFile}`,
], {
  cwd: new URL('..', import.meta.url),
  encoding: 'utf8',
});

assert.equal(result.status, 0, result.stderr || result.stdout);
assert.match(result.stdout, /meeting_app_fixture_track_readiness_report/);

const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.schema, 'meeting_app_fixture_track_readiness_report');
assert.equal(report.accepted, true);
assert.equal(report.platform_count, 5);
assert.equal(report.accepted_count, 5);
assert.deepEqual(report.missing, []);
assert.equal(report.coverage_by_platform.google_meet.speaker_track_mark, true);
assert.equal(report.coverage_by_platform.microsoft_teams.participant_track_mark, true);

console.log('ok meeting app fixture track readiness script');
