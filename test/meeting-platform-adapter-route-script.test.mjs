import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adapter-route-script-'));
const outDir = join(tmpDir, 'routes');
const reportFile = join(tmpDir, 'adapter-route-report.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-route.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,teams,zoom',
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_adapter_route_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.local_observer_first_count, 3);
assert.equal(report.provider_non_blocking_count, 3);
assert.equal(report.transcript_non_blocking_count, 3);
assert.equal(report.blocking_count, 0);
assert.equal(report.written_files.length, 3);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').first_route, 'local_observer_axis');
assert.equal(report.rows.find((row) => row.platform === 'microsoft_teams').route_file, join(outDir, 'microsoft_teams.json'));
assert.equal(report.rows.find((row) => row.platform === 'zoom').provider_blocks_realtime, false);
assert.equal(report.matrix.routes, undefined);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 3);
assert.equal(writtenReport.required_platforms.includes('teams'), true);
const googleRoute = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googleRoute.schema, 'meeting_platform_adapter_route');
assert.equal(googleRoute.routes[0].route, 'local_observer_axis');
assert.equal(googleRoute.entrypoints.browser_extension.matches.includes('https://meet.google.com/*'), true);
assert.equal(googleRoute.routes.find((route) => route.route === 'provider_reconcile').blocks_realtime_if_missing, false);
assert.equal(googleRoute.routes.find((route) => route.route === 'post_meeting_artifact_import').blocks_realtime_if_missing, false);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-route.mjs',
  '--platforms=webex',
  `--out-dir=${join(tmpDir, 'webex-routes')}`,
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_adapter_route_report/);
assert.match(textStdout, /local_first=1/);
assert.match(textStdout, /webex: mode=local_observer_first_provider_reconcile/);

console.log('ok meeting platform adapter route script');
