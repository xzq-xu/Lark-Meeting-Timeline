import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adapter-import-plan-script-'));
const exportDir = join(tmpDir, 'exports');
const importDir = join(tmpDir, 'imports');
const reportFile = join(tmpDir, 'adapter-import-plan-report.json');
const baseUrl = 'https://timeline.example.com';

await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-export-package.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet',
  '--target=static',
  `--out-dir=${exportDir}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-import-plan.mjs',
  `--dir=${exportDir}`,
  '--platforms=google-meet',
  '--target=static',
  `--out-dir=${importDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_adapter_import_plan_report');
assert.equal(report.ok, true);
assert.equal(report.target, 'static');
assert.equal(report.package_count, 1);
assert.equal(report.accepted_count, 1);
assert.equal(report.blocked_count, 0);
assert.equal(report.adapter_preflight_startup_ready_count, 1);
assert.equal(report.adapter_preflight_realtime_ready_count, 0);
assert.equal(report.available_file_count >= 8, true);
assert.equal(report.rows[0].platform, 'google_meet');
assert.equal(report.rows[0].selected_surface, 'browser_extension');
assert.equal(report.rows[0].file_coverage_ready, true);
assert.equal(report.rows[0].adapter_blueprint_available, true);
assert.equal(report.rows[0].adapter_blueprint_primary_surface, 'browser_extension');
assert.equal(report.rows[0].adapter_preflight_status, 'needs_live_page_evidence');
assert.equal(report.rows[0].adapter_preflight_selected_surface, 'browser_extension');
assert.equal(report.rows[0].adapter_preflight_startup_ready, true);
assert.equal(report.rows[0].adapter_preflight_realtime_ready, false);
assert.equal(report.rows[0].plan_file, join(importDir, 'google_meet', 'adapter-import-plan.json'));

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 1);

const importPlan = JSON.parse(await readFile(join(importDir, 'google_meet', 'adapter-import-plan.json'), 'utf8'));
assert.equal(importPlan.schema, 'meeting_platform_adapter_import_plan');
assert.equal(importPlan.accepted, true);
assert.equal(importPlan.adapter_blueprint.primary_surface, 'browser_extension');
assert.equal(importPlan.adapter_blueprint.provider_blocks_realtime, false);
assert.equal(importPlan.runtime_contract.adapter_preflight_required_before_realtime_insert, true);
assert.equal(importPlan.adapter_preflight.status, 'needs_live_page_evidence');
assert.equal(importPlan.adapter_preflight.realtime_annotation_ready, false);
assert.equal(importPlan.host_file_coverage.status, 'complete');
assert.equal(importPlan.install_steps.find((step) => step.id === 'run_adapter_preflight').sdk_method, 'platformAdapterPreflight');
assert.equal(importPlan.install_steps.find((step) => step.id === 'insert_realtime_marks').sdk_method, 'insertAnnotation');

const { stdout: missingStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-import-plan.mjs',
  `--dir=${exportDir}`,
  '--platforms=google-meet',
  '--target=static',
  '--available-files=google_meet/adapter-export-package.json',
  `--out-dir=${join(tmpDir, 'missing-imports')}`,
  '--json=true',
], {
  cwd: repoRoot,
});
const missingReport = JSON.parse(missingStdout);
assert.equal(missingReport.package_count, 1);
assert.equal(missingReport.accepted_count, 0);
assert.equal(missingReport.blocked_count, 1);
assert.equal(missingReport.rows[0].first_issue, 'missing_required_host_files');

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-import-plan.mjs',
  `--dir=${exportDir}`,
  '--platforms=google-meet',
  '--target=static',
  `--out-dir=${join(tmpDir, 'text-imports')}`,
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_adapter_import_plan_report/);
assert.match(textStdout, /google_meet: accepted=yes/);
assert.match(textStdout, /preflight=needs_live_page_evidence/);
assert.match(textStdout, /realtime=no/);

console.log('ok meeting platform adapter import plan script');
