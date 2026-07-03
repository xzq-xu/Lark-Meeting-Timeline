import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adapter-contract-script-'));
const outDir = join(tmpDir, 'contracts');
const reportFile = join(tmpDir, 'adapter-contract-report.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-contract.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,teams,local-detector',
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_adapter_contract_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.browser_observer_count, 2);
assert.equal(report.provider_observer_count, 2);
assert.equal(report.acceptance_target, 'contract');
assert.equal(report.acceptance.accepted_count, 3);
assert.equal(report.acceptance.rejected_count, 0);
assert.equal(report.written_files.length, 3);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').browser_matches.includes('https://meet.google.com/*'), true);
assert.equal(report.rows.find((row) => row.platform === 'microsoft_teams').provider_start_events.includes('meetingCallEvents.created'), true);
assert.equal(report.rows.find((row) => row.platform === 'local_detector').provider_observer, false);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').insert_mark_endpoint, `${baseUrl}/api/annotations`);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').transcript_realtime_dependency, false);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 3);
assert.equal(writtenReport.acceptance.rows.find((row) => row.platform === 'google_meet').accepted, true);
const googleContract = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googleContract.schema, 'meeting_platform_adapter_contract');
assert.equal(googleContract.realtime_axis.rules.includes('create_or_bind_axis_from_local_observer_before_provider_event_arrives'), true);
assert.equal(googleContract.provider_observer.required_for_realtime, false);
assert.equal(googleContract.annotations.endpoints.insertMark, `${baseUrl}/api/annotations`);
assert.equal(googleContract.implementation.kit_methods.includes('platformAdapterContract'), true);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-contract.mjs',
  '--platforms=webex',
  `--out-dir=${join(tmpDir, 'webex-contract')}`,
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_adapter_contract_report/);
assert.match(textStdout, /accepted=yes/);
assert.match(textStdout, /webex: mode=/);

await assert.rejects(
  execFileAsync(process.execPath, [
    'scripts/meeting-platform-adapter-contract.mjs',
    '--platforms=zoom',
    '--write-contracts=false',
    '--fail-on-incomplete=true',
  ], {
    cwd: repoRoot,
  }),
  (error) => error.code === 2,
);

await assert.rejects(
  execFileAsync(process.execPath, [
    'scripts/meeting-platform-adapter-contract.mjs',
    '--platforms=zoom',
    '--write-contracts=false',
    '--acceptance-target=production',
    '--fail-on-rejected=true',
  ], {
    cwd: repoRoot,
  }),
  (error) => error.code === 2,
);

console.log('ok meeting platform adapter contract script');
