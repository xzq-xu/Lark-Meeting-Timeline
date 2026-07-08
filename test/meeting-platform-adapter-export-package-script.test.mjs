import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adapter-export-package-script-'));
const outDir = join(tmpDir, 'exports');
const reportFile = join(tmpDir, 'adapter-export-package-report.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-export-package.mjs',
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
assert.equal(report.type, 'meeting_platform_adapter_export_package_report');
assert.equal(report.ok, true);
assert.equal(report.target, 'static');
assert.equal(report.platform_count, 1);
assert.equal(report.accepted_count, 1);
assert.equal(report.export_ready_count, 1);
assert.equal(report.adapter_preflight_startup_ready_count, 1);
assert.equal(report.adapter_preflight_realtime_ready_count, 0);
assert.equal(report.written_files.includes(join(outDir, 'google_meet', 'adapter-export-package.json')), true);
assert.equal(report.written_files.includes(join(outDir, 'google_meet', 'adapter-blueprint.json')), true);
assert.equal(report.written_files.includes(join(outDir, 'google_meet', 'runtime-bundle.json')), true);
assert.equal(report.written_files.includes(join(outDir, 'google_meet', 'provider-connection.json')), true);
assert.equal(report.rows[0].package_file, join(outDir, 'google_meet', 'adapter-export-package.json'));
assert.equal(report.rows[0].host_file_count >= 8, true);
assert.equal(report.rows[0].adapter_preflight_status, 'needs_live_page_evidence');
assert.equal(report.rows[0].adapter_preflight_selected_surface, 'browser_extension');
assert.equal(report.rows[0].adapter_preflight_startup_ready, true);
assert.equal(report.rows[0].adapter_preflight_realtime_ready, false);
assert.equal(report.matrix.packages, undefined);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 1);
assert.equal(writtenReport.matrix.packages, undefined);

const exportPackage = JSON.parse(await readFile(join(outDir, 'google_meet', 'adapter-export-package.json'), 'utf8'));
assert.equal(exportPackage.schema, 'meeting_platform_adapter_export_package');
assert.equal(exportPackage.accepted, true);
assert.equal(exportPackage.artifacts, undefined);
assert.equal(exportPackage.host_files.find((file) => file.source === 'adapter_blueprint').path, 'google_meet/adapter-blueprint.json');
assert.equal(exportPackage.host_files.find((file) => file.source === 'adapter_contract').path, 'google_meet/adapter-contract.json');
assert.equal(exportPackage.adapter_preflight.status, 'needs_live_page_evidence');
assert.equal(exportPackage.adapter_preflight.realtime_annotation_ready, false);
assert.equal(exportPackage.surface_entrypoints.provider_reconcile.blocks_realtime_annotation, false);

const adapterBlueprint = JSON.parse(await readFile(join(outDir, 'google_meet', 'adapter-blueprint.json'), 'utf8'));
assert.equal(adapterBlueprint.schema, 'meeting_platform_adapter_blueprint');
assert.equal(adapterBlueprint.platform, 'google_meet');
assert.equal(adapterBlueprint.surfaces.provider_reconcile.blocks_realtime, false);

const runtimeBundle = JSON.parse(await readFile(join(outDir, 'google_meet', 'runtime-bundle.json'), 'utf8'));
assert.equal(runtimeBundle.schema, 'meeting_platform_runtime_bundle');
assert.equal(runtimeBundle.platform, 'google_meet');

const providerConnection = JSON.parse(await readFile(join(outDir, 'google_meet', 'provider-connection.json'), 'utf8'));
assert.equal(providerConnection.schema, 'meeting_platform_provider_connection_pack');
assert.equal(providerConnection.platform, 'google_meet');

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-export-package.mjs',
  '--platforms=webex',
  '--target=static',
  `--out-dir=${join(tmpDir, 'webex-exports')}`,
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_adapter_export_package_report/);
assert.match(textStdout, /webex: accepted=yes/);
assert.match(textStdout, /export_ready=yes/);
assert.match(textStdout, /preflight=needs_live_page_evidence/);
assert.match(textStdout, /realtime=no/);

console.log('ok meeting platform adapter export package script');
