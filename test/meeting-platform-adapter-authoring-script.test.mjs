import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adapter-authoring-script-'));
const outDir = join(tmpDir, 'authoring');
const reportFile = join(tmpDir, 'adapter-authoring-report.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-authoring.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,Acme Rooms',
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_adapter_authoring_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 2);
assert.equal(report.built_in_count, 1);
assert.equal(report.external_authoring_count, 1);
assert.equal(report.browser_surface_ready_count, 1);
assert.equal(report.written_files.length, 2);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').plan_file, join(outDir, 'google_meet.json'));
assert.equal(report.rows.find((row) => row.platform === 'google_meet').provider_path, 'google_workspace_events_pubsub');
assert.equal(report.rows.find((row) => row.platform === 'acme_rooms').next_action, 'add_platform_setup_entry');
assert.equal(report.matrix.plans, undefined);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 2);
assert.equal(writtenReport.matrix.plans, undefined);

const googlePlan = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googlePlan.schema, 'meeting_platform_adapter_authoring_plan');
assert.equal(googlePlan.built_in, true);
assert.equal(googlePlan.browser_surface.matches.includes('https://meet.google.com/*'), true);
assert.equal(googlePlan.provider_reconcile.required_for_realtime, false);

const acmePlan = JSON.parse(await readFile(join(outDir, 'acme_rooms.json'), 'utf8'));
assert.equal(acmePlan.schema, 'meeting_platform_adapter_authoring_plan');
assert.equal(acmePlan.built_in, false);
assert.equal(acmePlan.current_sdk_status, 'external_adapter_authoring_required');
assert.match(acmePlan.normalizer_template, /normalizeAcmeRoomsEvent/);
assert.equal(acmePlan.next_actions.includes('add_event_normalizer'), true);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-authoring.mjs',
  '--platforms=webex',
  `--out-dir=${join(tmpDir, 'webex-authoring')}`,
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_adapter_authoring_report/);
assert.match(textStdout, /builtin=1/);
assert.match(textStdout, /webex: builtin=yes/);

console.log('ok meeting platform adapter authoring script');
