import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adapter-startup-script-'));
const outDir = join(tmpDir, 'startup-plans');
const reportFile = join(tmpDir, 'adapter-startup-report.json');
const outFile = join(tmpDir, 'google-startup-plan.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-startup.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,teams,zoom',
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
  '--include-plans=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_adapter_startup_plan_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.accepted_count, 3);
assert.equal(report.realtime_startup_ready_count, 3);
assert.equal(report.browser_surface_count, 1);
assert.equal(report.native_surface_count, 2);
assert.equal(report.provider_reconcile_surface_count, 0);
assert.equal(report.written_files.length, 3);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').selected_surface, 'browser_extension');
assert.equal(report.rows.find((row) => row.platform === 'google_meet').observe_action, 'observePlatformCandidates');
assert.equal(report.rows.find((row) => row.platform === 'google_meet').insert_action, 'insertAnnotation');
assert.equal(report.rows.find((row) => row.platform === 'microsoft_teams').selected_surface, 'native_detector');
assert.equal(report.rows.find((row) => row.platform === 'zoom').plan_file, join(outDir, 'zoom.json'));
assert.equal(report.matrix.plans.some((plan) => plan.reports), false);
assert.equal(report.matrix.plans.some((plan) => plan.decision), false);
assert.equal(report.matrix.plans.some((plan) => plan.adapter_blueprint), false);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 3);
assert.equal(writtenReport.rows.find((row) => row.platform === 'zoom').install_target, 'native_or_desktop_observer');

const googlePlan = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googlePlan.schema, 'meeting_platform_adapter_startup_plan');
assert.equal(googlePlan.platform, 'google_meet');
assert.equal(googlePlan.selected_surface, 'browser_extension');
assert.equal(googlePlan.message_contract.insert_annotation, 'meeting_timeline.insert_mark');
assert.equal(googlePlan.runtime_contract.annotation_timestamp_field, 'captured_at_ms');

const zoomPlan = JSON.parse(await readFile(join(outDir, 'zoom.json'), 'utf8'));
assert.equal(zoomPlan.selected_surface, 'native_detector');
assert.equal(zoomPlan.install_target, 'native_or_desktop_observer');

const { stdout: singleStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-startup.mjs',
  `--base-url=${baseUrl}`,
  '--platform=google-meet',
  '--url=https://meet.google.com/abc-defg-hij',
  `--out-file=${outFile}`,
  '--json=true',
  '--write-plans=false',
], {
  cwd: repoRoot,
});
const singleReport = JSON.parse(singleStdout);
assert.equal(singleReport.platform_count, 1);
assert.equal(singleReport.single_plan.platform, 'google_meet');
assert.equal(singleReport.single_plan.selected_surface, 'browser_extension');
assert.equal(singleReport.single_plan.realtime_startup_ready, true);
assert.equal(singleReport.written_files.length, 0);
const singlePlanFile = JSON.parse(await readFile(outFile, 'utf8'));
assert.equal(singlePlanFile.platform, 'google_meet');
assert.equal(singlePlanFile.selected_surface, 'browser_extension');

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-startup.mjs',
  '--platforms=webex',
  '--write-plans=false',
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_adapter_startup_plan_report/);
assert.match(textStdout, /platforms=1/);
assert.match(textStdout, /webex: accepted=yes/);

const { stdout: binStdout } = await execFileAsync(process.execPath, [
  'packages/meeting-timeline-sdk/bin/meeting-platform-adapter-startup.mjs',
  '--platforms=google-meet,zoom',
  '--json=true',
  '--write-plans=false',
], {
  cwd: repoRoot,
});
const binReport = JSON.parse(binStdout);
assert.equal(binReport.type, 'meeting_platform_adapter_startup_plan_report');
assert.equal(binReport.ok, true);
assert.equal(binReport.platform_count, 2);
assert.equal(binReport.rows.find((row) => row.platform === 'google_meet').selected_surface, 'browser_extension');
assert.equal(binReport.rows.find((row) => row.platform === 'zoom').selected_surface, 'native_detector');
assert.equal(binReport.written_files.length, 0);

console.log('ok meeting platform adapter startup script');
