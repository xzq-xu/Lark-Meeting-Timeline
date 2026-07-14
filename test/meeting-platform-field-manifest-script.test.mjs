import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-field-manifest-script-'));
const outDir = join(tmpDir, 'manifests');
const reportFile = join(tmpDir, 'field-manifest-report.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-field-manifest.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,zoom',
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_field_manifest_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 2);
assert.equal(report.written_files.length, 2);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').accepted_input_schemas.includes('raw_field_evidence'), true);
assert.equal(report.rows.find((row) => row.platform === 'zoom').required_provider_coverage.includes('meeting_start'), true);
assert.match(report.rows.find((row) => row.platform === 'zoom').build_field_evidence_command, /meeting-platform:field-evidence/);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 2);
const googleManifest = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googleManifest.schema, 'meeting_platform_field_capture_manifest');
assert.equal(googleManifest.file_contract.files.field_evidence_input.endsWith('/meeting-platform-field-evidence/google_meet.json'), true);
assert.equal(googleManifest.automation.commands.build_field_evidence.includes(`--base-url=${baseUrl}`), true);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-field-manifest.mjs',
  '--platforms=webex',
  `--out-dir=${join(tmpDir, 'webex-manifest')}`,
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_field_manifest_report/);
assert.match(textStdout, /webex: status=/);

await assert.rejects(
  execFileAsync(process.execPath, [
    'scripts/meeting-platform-field-manifest.mjs',
    '--platforms=zoom',
    '--write-manifests=false',
    '--fail-on-incomplete=true',
  ], {
    cwd: repoRoot,
  }),
  (error) => error.code === 2,
);

console.log('ok meeting platform field manifest script');
