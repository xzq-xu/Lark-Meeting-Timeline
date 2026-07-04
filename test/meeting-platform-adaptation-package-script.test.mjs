import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adaptation-package-script-'));
const outDir = join(tmpDir, 'packages');
const reportFile = join(tmpDir, 'adaptation-package-report.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adaptation-package.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,teams,zoom',
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_adaptation_package_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.sdk_wiring_ready_count, 3);
assert.equal(report.browser_observer_count, 3);
assert.equal(report.provider_observer_count, 3);
assert.equal(report.written_files.length, 3);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').package_file, join(outDir, 'google_meet.json'));
assert.equal(report.rows.find((row) => row.platform === 'google_meet').browser_match_count, 1);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').runtime_event_action_count, 14);
assert.equal(report.rows.find((row) => row.platform === 'microsoft_teams').browser_match_count, 2);
assert.equal(report.rows.find((row) => row.platform === 'zoom').browser_match_count, 3);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 3);
assert.equal(writtenReport.matrix.packages, undefined);
const googlePackage = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googlePackage.schema, 'meeting_platform_adaptation_package');
assert.equal(googlePackage.extension.matches.includes('https://meet.google.com/*'), true);
assert.equal(googlePackage.annotation_pipeline.insert_endpoint, `${baseUrl}/api/annotations`);
assert.equal(googlePackage.annotation_pipeline.runtime_event_plan.schema, 'meeting_platform_runtime_event_plan');
assert.equal(googlePackage.annotation_pipeline.runtime_event_actions.includes('insert_annotation'), true);
assert.equal(googlePackage.runtime_event_plan.endpoint, `${baseUrl}/api/meeting-platform/runtime-events`);
assert.equal(googlePackage.runtime_event_plan.realtime_contract.provider_events_required_for_realtime, false);
assert.equal(googlePackage.runtime_event_plan.supported_actions.includes('run_manifest'), true);
assert.equal(googlePackage.runtime_event_plan.supported_actions.includes('run_handoff_readiness'), true);
assert.equal(googlePackage.provider_observer.required_for_realtime, false);
assert.equal(googlePackage.transcript.blocks_realtime_annotation, false);
assert.equal(googlePackage.readiness.sdk_wiring_ready, true);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adaptation-package.mjs',
  '--platforms=webex',
  `--out-dir=${join(tmpDir, 'webex-package')}`,
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_adaptation_package_report/);
assert.match(textStdout, /sdk_ready=1/);
assert.match(textStdout, /webex: mode=/);
assert.match(textStdout, /runtime_actions=14/);

console.log('ok meeting platform adaptation package script');
