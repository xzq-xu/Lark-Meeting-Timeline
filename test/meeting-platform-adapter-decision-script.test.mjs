import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adapter-decision-script-'));
const outDir = join(tmpDir, 'adapter-decisions');
const reportFile = join(tmpDir, 'adapter-decision-report.json');
const outFile = join(tmpDir, 'google-decision.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-decision.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,teams,zoom',
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
  '--include-decisions=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_adapter_decision_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.accepted_count, 3);
assert.equal(report.realtime_ready_count, 3);
assert.equal(report.host_checklist_ready_count, 3);
assert.equal(report.written_files.length, 3);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').selected_surface, 'browser_extension');
assert.equal(report.rows.find((row) => row.platform === 'google_meet').selected_observer_mode, 'browser_dom_observer');
assert.equal(report.rows.find((row) => row.platform === 'google_meet').host_checklist_ready, true);
assert.equal(report.rows.find((row) => row.platform === 'microsoft_teams').selected_surface, 'native_detector');
assert.equal(report.rows.find((row) => row.platform === 'microsoft_teams').fallback_surfaces.includes('browser_extension'), true);
assert.equal(report.matrix.decisions.some((decision) => decision.reports), false);
assert.equal(report.matrix.decisions.find((decision) => decision.platform === 'google_meet').adaptation_strategy.host_integration_checklist.steps.find((step) => step.id === 'insert_realtime_annotation').timestamp_field, 'captured_at_ms');

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 3);
assert.equal(writtenReport.rows.find((row) => row.platform === 'zoom').first_evidence_to_collect, 'native_window_or_accessibility_snapshot_active_meeting');

const googleDecision = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googleDecision.schema, 'meeting_platform_adapter_decision');
assert.equal(googleDecision.platform, 'google_meet');
assert.equal(googleDecision.adaptation_strategy.host_integration_checklist.ready_for_realtime_host_wiring, true);
assert.equal(googleDecision.adaptation_strategy.host_integration_checklist.steps.find((step) => step.id === 'run_adapter_preflight').sdk_method, 'platformAdapterCurrentWindowPreflight');

const { stdout: singleStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-decision.mjs',
  `--base-url=${baseUrl}`,
  '--platform=google-meet',
  '--url=https://meet.google.com/abc-defg-hij',
  `--out-file=${outFile}`,
  '--json=true',
  '--write-decisions=false',
], {
  cwd: repoRoot,
});
const singleReport = JSON.parse(singleStdout);
assert.equal(singleReport.platform_count, 1);
assert.equal(singleReport.single_decision.platform, 'google_meet');
assert.equal(singleReport.single_decision.selected_surface, 'browser_extension');
assert.equal(singleReport.single_decision.adaptation_strategy.host_integration_checklist.ready_for_realtime_host_wiring, true);
assert.equal(singleReport.written_files.length, 0);
const singleDecisionFile = JSON.parse(await readFile(outFile, 'utf8'));
assert.equal(singleDecisionFile.platform, 'google_meet');
assert.equal(singleDecisionFile.selected_surface, 'browser_extension');

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-decision.mjs',
  '--platforms=webex',
  '--write-decisions=false',
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_adapter_decision_report/);
assert.match(textStdout, /platforms=1/);
assert.match(textStdout, /webex: accepted=yes/);

const { stdout: binStdout } = await execFileAsync(process.execPath, [
  'packages/meeting-timeline-sdk/bin/meeting-platform-adapter-decision.mjs',
  '--platforms=google-meet,zoom',
  '--json=true',
  '--write-decisions=false',
], {
  cwd: repoRoot,
});
const binReport = JSON.parse(binStdout);
assert.equal(binReport.type, 'meeting_platform_adapter_decision_report');
assert.equal(binReport.ok, true);
assert.equal(binReport.platform_count, 2);
assert.equal(binReport.rows.find((row) => row.platform === 'google_meet').selected_surface, 'browser_extension');
assert.equal(binReport.rows.find((row) => row.platform === 'zoom').selected_surface, 'native_detector');
assert.equal(binReport.written_files.length, 0);

console.log('ok meeting platform adapter decision script');
