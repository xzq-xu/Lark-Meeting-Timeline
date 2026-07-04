import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-runtime-event-plan-script-'));
const outDir = join(tmpDir, 'plans');
const reportFile = join(tmpDir, 'runtime-event-plan-report.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-runtime-event-plan.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,teams,zoom',
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_runtime_event_plan_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.realtime_provider_dependency_count, 0);
assert.equal(report.transcript_realtime_dependency_count, 0);
assert.equal(report.written_files.length, 3);
assert.equal(report.rows.some((row) => row.platform === 'google_meet' && row.action === 'insert_annotation'), true);
assert.equal(report.rows.some((row) => row.platform === 'google_meet' && row.action === 'observe_platform_candidates'), true);
assert.equal(report.rows.some((row) => row.platform === 'microsoft_teams' && row.action === 'speaker_track'), true);
assert.equal(report.rows.find((row) => row.platform === 'zoom' && row.action === 'provider_event').plan_file, join(outDir, 'zoom.json'));

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 21);
assert.equal(writtenReport.matrix.plans, undefined);
const googlePlan = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googlePlan.schema, 'meeting_platform_runtime_event_plan');
assert.equal(googlePlan.endpoint, `${baseUrl}/api/meeting-platform/runtime-events`);
assert.equal(googlePlan.examples.insert_annotation.annotation.captured_at_ms, 1_782_614_415_000);
assert.equal(googlePlan.examples.observe_platform_candidates.action, 'observe_platform_candidates');
assert.equal(googlePlan.realtime_contract.provider_events_required_for_realtime, false);
assert.equal(googlePlan.realtime_contract.transcript_required_for_realtime, false);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-runtime-event-plan.mjs',
  '--platforms=webex',
  `--out-dir=${join(tmpDir, 'webex-plan')}`,
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_runtime_event_plan_report/);
assert.match(textStdout, /platforms=1/);
assert.match(textStdout, /webex: action=insert_annotation/);
assert.match(textStdout, /webex: action=observe_platform_candidates/);

console.log('ok meeting platform runtime event plan script');
