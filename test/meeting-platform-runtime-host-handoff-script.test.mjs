import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-runtime-host-handoff-script-'));
const reportFile = join(tmpDir, 'runtime-host-handoff-report.json');
const outDir = join(tmpDir, 'handoffs');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-runtime-host-handoff.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,teams,zoom',
  `--report-file=${reportFile}`,
  `--out-dir=${outDir}`,
], {
  cwd: repoRoot,
});

assert.match(stdout, /meeting_platform_runtime_host_handoff_report/);
assert.match(stdout, /accepted=3\/3/);
assert.match(stdout, /google_meet: accepted=yes host=yes/);

const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.type, 'meeting_platform_runtime_host_handoff_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.accepted_count, 3);
assert.equal(report.host_ready_count, 3);
assert.equal(report.provider_blocking_count, 0);
assert.equal(report.transcript_blocking_count, 0);
assert.equal(report.written_files.length, 3);
assert.equal(report.rows.find((row) => row.platform === 'microsoft_teams').host_hook_count, 5);

const googleHandoff = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googleHandoff.schema, 'meeting_platform_runtime_host_handoff');
assert.equal(googleHandoff.acceptance.accepted, true);
assert.equal(googleHandoff.browser.matches[0], 'https://meet.google.com/*');
assert.equal(googleHandoff.host_hooks.some((hook) => hook.call === 'host.candidateMissing()'), true);

const { stdout: jsonStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-runtime-host-handoff.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=webex,lark',
  '--include-handoffs=true',
  '--json=true',
], {
  cwd: repoRoot,
  maxBuffer: 20 * 1024 * 1024,
});
const jsonReport = JSON.parse(jsonStdout);
assert.equal(jsonReport.ok, true);
assert.equal(jsonReport.matrix.handoffs.length, 2);
assert.equal(jsonReport.rows.find((row) => row.platform === 'lark').timestamp_field, 'captured_at_ms');

await execFileAsync(process.execPath, [
  'scripts/meeting-platform-runtime-host-handoff.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet',
  '--fail-on-not-ready=true',
], {
  cwd: repoRoot,
});

console.log('ok meeting platform runtime host handoff script');
