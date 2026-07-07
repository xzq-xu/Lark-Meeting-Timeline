import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adapter-blueprint-script-'));
const outDir = join(tmpDir, 'blueprints');
const reportFile = join(tmpDir, 'adapter-blueprint-report.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-blueprint.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,teams,zoom',
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_adapter_blueprint_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.ready_count, 3);
assert.equal(report.provider_non_blocking_count, 3);
assert.equal(report.transcript_non_blocking_count, 3);
assert.equal(report.blocking_count, 0);
assert.equal(report.written_files.length, 3);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').primary_surface, 'browser_extension');
assert.equal(report.rows.find((row) => row.platform === 'microsoft_teams').native_recommended, true);
assert.equal(report.rows.find((row) => row.platform === 'zoom').primary_surface, 'native_detector');
assert.equal(report.rows.find((row) => row.platform === 'zoom').blueprint_file, join(outDir, 'zoom', 'adapter-blueprint.json'));
assert.equal(report.matrix.blueprints, undefined);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 3);
assert.equal(writtenReport.required_platforms.includes('teams'), true);

const googleBlueprint = JSON.parse(await readFile(join(outDir, 'google_meet', 'adapter-blueprint.json'), 'utf8'));
assert.equal(googleBlueprint.schema, 'meeting_platform_adapter_blueprint');
assert.equal(googleBlueprint.platform, 'google_meet');
assert.equal(googleBlueprint.primary_surface, 'browser_extension');
assert.equal(googleBlueprint.surfaces.browser_extension.evidence.required_fields.includes('captured_at_ms'), true);
assert.equal(googleBlueprint.surfaces.provider_reconcile.blocks_realtime, false);
assert.equal(googleBlueprint.runtime_contract.transcript_blocks_realtime, false);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-blueprint.mjs',
  '--platforms=webex',
  `--out-dir=${join(tmpDir, 'webex-blueprints')}`,
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_adapter_blueprint_report/);
assert.match(textStdout, /ready=1/);
assert.match(textStdout, /webex: ready=yes/);

const { stdout: binStdout } = await execFileAsync(process.execPath, [
  'packages/meeting-timeline-sdk/bin/meeting-platform-adapter-blueprint.mjs',
  '--platforms=zoom',
  '--json=true',
  '--write-blueprints=false',
], {
  cwd: repoRoot,
});
const binReport = JSON.parse(binStdout);
assert.equal(binReport.platform_count, 1);
assert.equal(binReport.rows[0].platform, 'zoom');
assert.equal(binReport.rows[0].primary_surface, 'native_detector');
assert.equal(binReport.written_files.length, 0);

console.log('ok meeting platform adapter blueprint script');
