import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adapter-smoke-script-'));
const exportDir = join(tmpDir, 'exports');
const importDir = join(tmpDir, 'imports');
const manifestFile = join(tmpDir, 'adapter-install-manifest.json');
const smokeReportFile = join(tmpDir, 'adapter-smoke.json');
const cliReportFile = join(tmpDir, 'adapter-smoke-report.json');
const baseUrl = 'https://timeline.example.com';

const { stdout: generatedStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-smoke.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,zoom,teams',
  '--target=static',
  '--captured-at-ms=1782900000000',
  '--json=true',
], {
  cwd: repoRoot,
});
const generatedReport = JSON.parse(generatedStdout);
assert.equal(generatedReport.type, 'meeting_platform_adapter_smoke_cli_report');
assert.equal(generatedReport.ok, true);
assert.equal(generatedReport.generated_manifest, true);
assert.equal(generatedReport.platform_count, 3);
assert.equal(generatedReport.accepted_count, 3);
assert.equal(generatedReport.rows.every((row) => row.captured_at_ms_preserved === true), true);

await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-export-package.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,webex,lark',
  '--target=static',
  `--out-dir=${exportDir}`,
  '--json=true',
], {
  cwd: repoRoot,
});

await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-import-plan.mjs',
  `--dir=${exportDir}`,
  '--platforms=google-meet,webex,lark',
  '--target=static',
  `--out-dir=${importDir}`,
  '--json=true',
], {
  cwd: repoRoot,
});

await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-install-manifest.mjs',
  `--dir=${importDir}`,
  '--platforms=google-meet,webex,lark',
  '--target=static',
  `--out-file=${manifestFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-smoke.mjs',
  `--manifest-file=${manifestFile}`,
  '--platforms=google-meet,webex,lark',
  `--out-file=${smokeReportFile}`,
  `--report-file=${cliReportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});
const report = JSON.parse(stdout);
assert.equal(report.ok, true);
assert.equal(report.generated_manifest, false);
assert.deepEqual(report.platforms, ['google_meet', 'webex', 'lark']);
assert.equal(report.rows.find((row) => row.platform === 'webex').provider_reconcile_nonblocking, true);

const writtenCliReport = JSON.parse(await readFile(cliReportFile, 'utf8'));
assert.equal(writtenCliReport.ok, true);

const smokeReport = JSON.parse(await readFile(smokeReportFile, 'utf8'));
assert.equal(smokeReport.schema, 'meeting_platform_adapter_smoke_report');
assert.equal(smokeReport.accepted, true);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-smoke.mjs',
  '--platforms=google-meet,zoom',
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_adapter_smoke_cli_report/);
assert.match(textStdout, /google_meet: accepted=yes/);
assert.match(textStdout, /zoom: accepted=yes/);

const { stdout: missingStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-smoke.mjs',
  `--manifest-file=${join(tmpDir, 'missing-install-manifest.json')}`,
  '--json=true',
], {
  cwd: repoRoot,
});
const missingReport = JSON.parse(missingStdout);
assert.equal(missingReport.ok, false);
assert.match(missingReport.manifest_read_error, /ENOENT/);

console.log('ok meeting platform adapter smoke script');
