import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-runtime-bundle-script-'));
const outDir = join(tmpDir, 'bundles');
const reportFile = join(tmpDir, 'runtime-bundle-report.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-runtime-bundle.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,teams,zoom',
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_runtime_bundle_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.runtime_ready_count, 3);
assert.equal(report.sdk_wiring_ready_count, 3);
assert.equal(report.matrix.lightweight_connector_ready_count, 3);
assert.equal(report.provider_required_for_realtime_count, 0);
assert.equal(report.transcript_blocking_count, 0);
assert.equal(report.written_files.length, 3);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').bundle_file, join(outDir, 'google_meet.json'));
assert.equal(report.rows.find((row) => row.platform === 'google_meet').sample_interval_ms, 10_000);
assert.equal(report.rows.find((row) => row.platform === 'microsoft_teams').mutation_debounce_ms, 180);
assert.equal(report.rows.find((row) => row.platform === 'zoom').browser_match_count, 3);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 3);
assert.equal(writtenReport.matrix.bundles, undefined);
const googleBundle = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googleBundle.schema, 'meeting_platform_runtime_bundle');
assert.equal(googleBundle.browser.matches.includes('https://meet.google.com/*'), true);
assert.equal(googleBundle.host.endpoints.insertMark, `${baseUrl}/api/annotations`);
assert.equal(googleBundle.runtime.lightweight_connector_bridge.install_function, 'installMeetingPlatformConnectorContentScriptBridge');
assert.equal(googleBundle.runtime.lightweight_connector_bridge.options.baseUrl, baseUrl);
assert.equal(googleBundle.messaging.lightweight_connector_message_types.includes('meeting_timeline.sample_tracks'), true);
assert.equal(googleBundle.messaging.lightweight_connector_message_types.includes('meeting_timeline.preflight_current_window'), true);
assert.equal(googleBundle.messaging.accepted_methods.includes('insertMark'), true);
assert.equal(googleBundle.provider_reconcile.required_for_realtime, false);
assert.equal(googleBundle.transcript.blocks_realtime_annotation, false);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-runtime-bundle.mjs',
  '--platforms=webex',
  `--out-dir=${join(tmpDir, 'webex-bundle')}`,
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_runtime_bundle_report/);
assert.match(textStdout, /runtime_ready=1/);
assert.match(textStdout, /webex: runtime=yes/);

console.log('ok meeting platform runtime bundle script');
