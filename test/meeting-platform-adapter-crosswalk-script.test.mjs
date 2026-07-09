import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-crosswalk.mjs',
  '--platforms=google-meet,teams,zoom',
  '--json=true',
  '--base-url=https://timeline.example.com',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_adapter_crosswalk_report');
assert.equal(report.platform_count, 3);
assert.equal(report.rows.length, 3);
assert.equal(typeof report.rows.find((row) => row.platform === 'google_meet').compatibility_score, 'number');
assert.equal(report.rows.find((row) => row.platform === 'google_meet').candidate_observer_ready, true);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').provider_path.includes('google_workspace_events_pubsub'), true);
assert.equal(typeof report.rows.find((row) => row.platform === 'microsoft_teams').realtime, 'boolean');
assert.equal(report.totals.ready_for_realtime_count >= 0, true);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-crosswalk.mjs',
  '--platforms=google-meet,teams,zoom',
  '--base-url=https://timeline.example.com',
], {
  cwd: repoRoot,
});

assert.match(textStdout, /meeting_platform_adapter_crosswalk_report \| platforms=3/);
assert.match(textStdout, /google_meet:/);

console.log('ok meeting platform adapter crosswalk script');
