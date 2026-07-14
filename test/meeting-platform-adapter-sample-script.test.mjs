import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adapter-sample-script-'));
const outDir = join(tmpDir, 'samples');
const reportFile = join(tmpDir, 'adapter-sample-report.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-sample.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,teams,zoom',
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
], {
  cwd: repoRoot,
});

assert.match(stdout, /meeting_platform_adapter_sample_report/);
assert.match(stdout, /accepted=3\/3/);

const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.type, 'meeting_platform_adapter_sample_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.accepted_count, 3);
assert.equal(report.rejected_count, 0);
assert.equal(report.written_files.length, 3);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').provider_event_count, 5);
assert.equal(report.rows.find((row) => row.platform === 'microsoft_teams').timeline_method_counts.insertMark, 1);

assert.equal(report.matrix.accepted_count, 3);
assert.equal(report.rows.find((row) => row.platform === 'zoom').accepted, true);

const googleSample = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googleSample.schema, 'meeting_platform_adapter_sample');
assert.equal(googleSample.platform, 'google_meet');
assert.equal(googleSample.accepted, true);
assert.equal(googleSample.endpoint, `${baseUrl}/api/platform-events/google-meet`);
assert.equal(googleSample.timeline_method_counts.insertMark, 1);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-sample.mjs',
  '--platforms=webex',
  `--out-dir=${join(tmpDir, 'webex-sample')}`,
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_adapter_sample_report/);
assert.match(textStdout, /accepted=1\/1/);
assert.match(textStdout, /webex: accepted=yes/);

console.log('ok meeting platform adapter sample script');
