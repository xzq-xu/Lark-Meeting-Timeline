import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adapter-launch-plan-script-'));
const exportDir = join(tmpDir, 'exports');
const importDir = join(tmpDir, 'imports');
const manifestFile = join(tmpDir, 'adapter-install-manifest.json');
const launchPlanFile = join(tmpDir, 'adapter-launch-plan.json');
const reportFile = join(tmpDir, 'adapter-launch-plan-report.json');
const baseUrl = 'https://timeline.example.com';

await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-export-package.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,zoom',
  '--target=static',
  `--out-dir=${exportDir}`,
  '--json=true',
], {
  cwd: repoRoot,
});

await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-import-plan.mjs',
  `--dir=${exportDir}`,
  '--platforms=google-meet,zoom',
  '--target=static',
  `--out-dir=${importDir}`,
  '--json=true',
], {
  cwd: repoRoot,
});

await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-install-manifest.mjs',
  `--dir=${importDir}`,
  '--platforms=google-meet,zoom',
  '--target=static',
  `--out-file=${manifestFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-launch-plan.mjs',
  `--manifest-file=${manifestFile}`,
  '--url=https://meet.google.com/abc-defg-hij',
  `--out-file=${launchPlanFile}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_adapter_launch_plan_report');
assert.equal(report.ok, true);
assert.equal(report.platform, 'google_meet');
assert.equal(report.selected_surface, 'browser_extension');
assert.equal(report.adapter_blueprint_primary_surface, 'browser_extension');
assert.equal(report.adapter_blueprint_first_gate, 'local_candidate_preflight_accepts_active_meeting');
assert.equal(report.raw_signal_validation_status, 'ready');
assert.equal(report.first_runtime_action, 'read_adapter_selection');

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.ok, true);

const launchPlan = JSON.parse(await readFile(launchPlanFile, 'utf8'));
assert.equal(launchPlan.schema, 'meeting_platform_adapter_launch_plan');
assert.equal(launchPlan.accepted, true);
assert.equal(launchPlan.adapter_blueprint.primary_surface, 'browser_extension');
assert.equal(launchPlan.adapter_blueprint.first_acceptance_gate, 'local_candidate_preflight_accepts_active_meeting');
assert.equal(launchPlan.raw_signal_validation.status, 'ready');
assert.equal(launchPlan.adapter_selection.axis_surface, 'browser_extension');
assert.equal(launchPlan.axis_contract.raw_signal_validation_required_before_preflight, true);
assert.equal(launchPlan.runtime_actions[0].id, 'read_adapter_selection');
assert.equal(launchPlan.runtime_actions.find((action) => action.id === 'validate_raw_signal').sdk_method, 'platformRawSignalBatch');
assert.equal(launchPlan.surface_entrypoint.content_script.matches.includes('https://meet.google.com/*'), true);

const { stdout: providerStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-launch-plan.mjs',
  `--manifest-file=${manifestFile}`,
  '--platform=google-meet',
  '--surface=provider-reconcile',
  '--json=true',
], {
  cwd: repoRoot,
});
const providerReport = JSON.parse(providerStdout);
assert.equal(providerReport.ok, false);
assert.equal(providerReport.issues.some((issue) => issue.code === 'provider_reconcile_not_realtime_launch_surface'), true);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-launch-plan.mjs',
  `--manifest-file=${manifestFile}`,
  '--platform=zoom',
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_adapter_launch_plan_report/);
assert.match(textStdout, /platform=zoom/);
assert.match(textStdout, /blueprint_surface=native_detector/);

console.log('ok meeting platform adapter launch plan script');
