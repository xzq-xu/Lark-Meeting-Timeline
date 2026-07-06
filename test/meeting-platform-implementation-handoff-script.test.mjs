import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-implementation-handoff-script-'));
const outDir = join(tmpDir, 'handoffs');
const reportFile = join(tmpDir, 'implementation-handoff-report.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-implementation-handoff.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,teams,zoom',
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_implementation_handoff_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.implementation_ready_count, 3);
assert.equal(report.pilot_ready_count, 3);
assert.equal(report.production_ready_count, 0);
assert.equal(report.recommended_first_platform, 'google_meet');
assert.equal(report.recommended_first_surface, 'browser_extension');
assert.equal(report.written_files.length, 3);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').handoff_file, join(outDir, 'google_meet.json'));
assert.equal(report.rows.find((row) => row.platform === 'google_meet').runtime_event_endpoint, `${baseUrl}/api/meeting-platform/runtime-events`);
assert.equal(report.rows.find((row) => row.platform === 'microsoft_teams').provider_permission_risk, 'tenant_admin_consent_and_subscription_renewal');
assert.equal(report.rows.find((row) => row.platform === 'zoom').browser_match_count, 3);
assert.equal(report.matrix.handoffs, undefined);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 3);
assert.equal(writtenReport.matrix.handoffs, undefined);

const googleHandoff = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googleHandoff.schema, 'meeting_platform_implementation_handoff');
assert.equal(googleHandoff.platform, 'google_meet');
assert.equal(googleHandoff.install_surface.browser_matches.includes('https://meet.google.com/*'), true);
assert.equal(googleHandoff.runtime_events.message_types.includes('meeting_timeline.sample_tracks'), true);
assert.equal(googleHandoff.provider_reconcile.path, 'google_workspace_events_pubsub');
assert.equal(googleHandoff.contracts.timestamp_field, 'captured_at_ms');
assert.equal(googleHandoff.acceptance.commands.runtime_bundle.includes('meeting-platform:runtime-bundle'), true);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-implementation-handoff.mjs',
  '--platforms=webex',
  `--out-dir=${join(tmpDir, 'webex-handoff')}`,
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_implementation_handoff_report/);
assert.match(textStdout, /implementation_ready=1/);
assert.match(textStdout, /webex: ready=yes/);

console.log('ok meeting platform implementation handoff script');
