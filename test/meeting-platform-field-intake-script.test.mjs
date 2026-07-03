import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-field-intake-script-'));
const reportFile = join(tmpDir, 'report.json');
const outDir = join(tmpDir, 'plans');
const evidenceDir = join(tmpDir, 'evidence');
const baseUrl = 'https://timeline.example.com';

const env = {
  ...process.env,
  GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
  GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL: 'meet-events@example.iam.gserviceaccount.com',
  ZOOM_WEBHOOK_SECRET_TOKEN: 'real-zoom-secret-token',
};

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-field-intake.mjs',
  `--base-url=${baseUrl}`,
  `--evidence-dir=${evidenceDir}`,
  `--out-dir=${outDir}`,
  '--platforms=google-meet,zoom',
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
  env,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_field_intake_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 2);
assert.equal(report.write_plans, true);
assert.equal(report.written_files.length, 2);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').provider_endpoint, `${baseUrl}/api/platform-events/google-meet`);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').missing_env.length, 0);
assert.equal(report.rows.find((row) => row.platform === 'zoom').missing_env.length, 0);
assert.equal(report.rows.find((row) => row.platform === 'zoom').validate_real_intake_command.includes('meeting-platform:real-intake'), true);
assert.equal(report.matrix, undefined);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 2);

const googlePlan = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googlePlan.schema, 'meeting_platform_field_intake_plan');
assert.equal(googlePlan.files.field_evidence_input, `${evidenceDir}/meeting-platform-field-evidence/google_meet.json`);
assert.equal(googlePlan.operator_steps.some((step) => step.id === 'capture_provider_start_end_events'), true);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-field-intake.mjs',
  `--base-url=${baseUrl}`,
  `--evidence-dir=${evidenceDir}`,
  `--out-dir=${outDir}`,
  '--platforms=google-meet',
  '--write-plans=false',
], {
  cwd: repoRoot,
  env,
});
assert.match(textStdout, /meeting_platform_field_intake_report/);
assert.match(textStdout, /platforms=1/);

console.log('ok meeting platform field intake script');
