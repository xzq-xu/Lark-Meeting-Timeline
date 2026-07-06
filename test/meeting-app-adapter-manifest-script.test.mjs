import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-app-adapter-manifest-'));
const outDir = join(tmpDir, 'manifests');
const reportFile = join(tmpDir, 'report.json');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-app-adapter-manifest.mjs',
  '--platforms=google-meet,teams,zoom',
  '--base-url=http://localhost:8787',
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
], {
  cwd: new URL('..', import.meta.url),
});

assert.match(stdout, /meeting_app_adapter_manifest_report/);
assert.match(stdout, /ok=yes/);
assert.match(stdout, /platforms=3/);
assert.match(stdout, /accepted=3/);
assert.match(stdout, /written=3/);

const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.type, 'meeting_app_adapter_manifest_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.accepted_count, 3);
assert.equal(report.matrix.manifests, undefined);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').accepted, true);

const google = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(google.schema, 'meeting_app_adapter_manifest');
assert.equal(google.platform, 'google_meet');
assert.equal(google.accepted, true);
assert.equal(google.contracts.timestamp_field, 'captured_at_ms');

const { stdout: jsonStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-app-adapter-manifest.mjs',
  '--platforms=webex',
  '--json=true',
], {
  cwd: new URL('..', import.meta.url),
});
const jsonReport = JSON.parse(jsonStdout);
assert.equal(jsonReport.ok, true);
assert.equal(jsonReport.rows[0].platform, 'webex');
assert.equal(jsonReport.matrix.manifests, undefined);

console.log('ok meeting app adapter manifest script');
