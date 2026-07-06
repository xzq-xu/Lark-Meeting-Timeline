import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adapter-acceptance-checklist-script-'));
const outDir = join(tmpDir, 'checklists');
const reportFile = join(tmpDir, 'acceptance-checklist-report.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-acceptance-checklist.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet',
  '--target=static',
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_adapter_acceptance_checklist_report');
assert.equal(report.ok, true);
assert.equal(report.target, 'static');
assert.equal(report.platform_count, 1);
assert.equal(report.accepted_count, 1);
assert.equal(report.blocked_count, 0);
assert.equal(report.written_files.length, 1);
assert.equal(report.rows[0].checklist_file, join(outDir, 'google_meet.json'));
assert.equal(report.rows[0].accepted, true);
assert.equal(report.matrix.checklists, undefined);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 1);
assert.equal(writtenReport.matrix.checklists, undefined);

const googleChecklist = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googleChecklist.schema, 'meeting_platform_adapter_acceptance_checklist');
assert.equal(googleChecklist.target, 'static');
assert.equal(googleChecklist.accepted, true);
assert.equal(googleChecklist.checklist.find((item) => item.id === 'provider_nonblocking_contract').passed, true);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-acceptance-checklist.mjs',
  '--platforms=webex',
  '--target=pilot',
  `--out-dir=${join(tmpDir, 'pilot-checklists')}`,
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_adapter_acceptance_checklist_report/);
assert.match(textStdout, /target=pilot/);
assert.match(textStdout, /webex: accepted=no/);

console.log('ok meeting platform adapter acceptance checklist script');
