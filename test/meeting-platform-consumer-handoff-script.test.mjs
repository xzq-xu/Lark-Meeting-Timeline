import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-consumer-handoff-script-'));
const reportFile = join(tmpDir, 'consumer-handoff-report.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-consumer-handoff.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,teams,zoom',
  `--report-file=${reportFile}`,
], {
  cwd: repoRoot,
});

assert.match(stdout, /meeting_platform_consumer_handoff/);
assert.match(stdout, /accepted=yes/);
assert.match(stdout, /consumer_ready=3/);
assert.match(stdout, /production=0/);

const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.schema, 'meeting_platform_consumer_handoff');
assert.equal(report.accepted, true);
assert.equal(report.platform_count, 3);
assert.equal(report.consumer_ready_count, 3);
assert.equal(report.candidate_observer_count, 3);
assert.equal(report.speaker_track_ready_count, 3);
assert.equal(report.participant_track_ready_count, 3);
assert.equal(report.production_ready_count, 0);
assert.equal(report.entrypoints.primary_modules.consumer_handoff, '@ai-annotation/meeting-timeline-sdk/adapters/platform-consumer-handoff');
assert.equal(report.rows.find((row) => row.platform === 'microsoft_teams').consumer_ready, true);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').adapter_first_route, 'local_observer_axis');

const { stdout: jsonStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-consumer-handoff.mjs',
  '--platforms=webex',
  '--json=true',
], {
  cwd: repoRoot,
});
const jsonReport = JSON.parse(jsonStdout);
assert.equal(jsonReport.rows[0].platform, 'webex');
assert.equal(jsonReport.rows[0].consumer_ready, true);

let strictStdout = '';
let strictCode = 0;
try {
  const result = await execFileAsync(process.execPath, [
    'scripts/meeting-platform-consumer-handoff.mjs',
    '--platforms=zoom',
    '--require-production-ready=true',
    '--fail-on-rejected=true',
  ], {
    cwd: repoRoot,
  });
  strictStdout = result.stdout;
} catch (error) {
  strictStdout = error.stdout;
  strictCode = error.code;
}
assert.equal(strictCode, 2);
assert.match(strictStdout, /accepted=no/);

console.log('ok meeting platform consumer handoff script');
