import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-conformance-script-'));
const reportFile = join(tmpDir, 'conformance-report.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-conformance.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,teams,zoom',
  `--report-file=${reportFile}`,
], {
  cwd: repoRoot,
});

assert.match(stdout, /meeting_platform_conformance_report/);
assert.match(stdout, /accepted=yes/);
assert.match(stdout, /routes=3/);
assert.match(stdout, /candidates=3/);

const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.schema, 'meeting_platform_conformance_report');
assert.equal(report.accepted, true);
assert.equal(report.platform_count, 3);
assert.equal(report.accepted_count, 3);
assert.equal(report.adapter_route_ready_count, 3);
assert.equal(report.candidate_observer_count, 3);
assert.equal(report.provider_required_for_realtime_count, 0);
assert.equal(report.transcript_blocking_count, 0);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').adapter_first_route, 'local_observer_axis');
assert.equal(report.rows.find((row) => row.platform === 'microsoft_teams').normalizer_available, true);

const { stdout: jsonStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-conformance.mjs',
  '--platforms=webex',
  '--json=true',
], {
  cwd: repoRoot,
});
const jsonReport = JSON.parse(jsonStdout);
assert.equal(jsonReport.rows[0].platform, 'webex');
assert.equal(jsonReport.rows[0].accepted, true);

console.log('ok meeting platform conformance script');
