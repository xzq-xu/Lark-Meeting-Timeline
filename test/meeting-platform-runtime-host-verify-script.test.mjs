import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-runtime-host-verify-script-'));
const reportFile = join(tmpDir, 'runtime-host-verify-report.json');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-runtime-host-verify.mjs',
  '--platforms=google-meet,teams,zoom',
  '--start-ms=1783356000000',
  '--end-offset-ms=2000',
  `--report-file=${reportFile}`,
], {
  cwd: repoRoot,
});

assert.match(stdout, /meeting_platform_runtime_host_verify_report/);
assert.match(stdout, /accepted=3\/3/);
assert.match(stdout, /google_meet: accepted=yes/);

const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.type, 'meeting_platform_runtime_host_verify_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.accepted_count, 3);
assert.equal(report.rows.find((row) => row.platform === 'microsoft_teams').actions.includes('insertMark'), true);
assert.equal(report.rows.find((row) => row.platform === 'zoom').signal_types.includes('meeting_ended'), true);

const { stdout: jsonStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-runtime-host-verify.mjs',
  '--platforms=webex,lark',
  '--include-reports=true',
  '--json=true',
], {
  cwd: repoRoot,
  maxBuffer: 20 * 1024 * 1024,
});
const jsonReport = JSON.parse(jsonStdout);
assert.equal(jsonReport.ok, true);
assert.equal(jsonReport.matrix.reports.length, 2);
assert.equal(jsonReport.rows.find((row) => row.platform === 'lark').start_capture_profile, 'lark');

await execFileAsync(process.execPath, [
  'scripts/meeting-platform-runtime-host-verify.mjs',
  '--platforms=google-meet',
  '--fail-on-failure=true',
], {
  cwd: repoRoot,
});

console.log('ok meeting platform runtime host verify script');
