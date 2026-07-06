import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-app-adapter-integration-package-'));
const outDir = join(tmpDir, 'packages');
const reportFile = join(tmpDir, 'report.json');
const inputFile = join(tmpDir, 'snapshots.json');
const evidenceFile = join(tmpDir, 'evidence.json');
await mkdir(tmpDir, { recursive: true });

const completeEvidence = (adapterKey) => ({
  adapter_key: adapterKey,
  live_dom_snapshot_count: 1,
  candidate_observation_count: 1,
  speaker_segments: [{ captured_at_ms: 1_782_614_401_000, speaker_label: 'Ada' }],
  participant_segments: [{ captured_at_ms: 1_782_614_401_000, participant_label: 'Ada' }],
  annotations: [{
    captured_at_ms: 1_782_614_402_000,
    meeting_id: `${adapterKey}-meeting`,
    axis_id: `${adapterKey}-axis`,
  }],
});

await writeFile(inputFile, `${JSON.stringify({
  google_meet: {
    url: 'https://meet.google.com/abc-defg-hij',
    page: {
      controls: [{ label: 'Leave call' }],
      participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace speaking' }],
    },
  },
  zoom: {
    app: { name: 'Zoom Workplace' },
    window: { title: 'Zoom Meeting', controls: [{ label: 'Leave Meeting' }] },
    meeting_id: 'zoom-local',
    tiles: [{ id: 'mira', ariaLabel: 'Mira Patel speaking' }],
  },
}, null, 2)}\n`, 'utf8');
await writeFile(evidenceFile, `${JSON.stringify({
  google_meet: completeEvidence('google_meet'),
  zoom: completeEvidence('zoom'),
}, null, 2)}\n`, 'utf8');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-app-adapter-integration-package.mjs',
  '--platforms=google-meet,zoom',
  `--input-file=${inputFile}`,
  `--evidence-file=${evidenceFile}`,
  '--target=production',
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
], {
  cwd: repoRoot,
});

assert.match(stdout, /meeting_app_adapter_integration_package_report/);
assert.match(stdout, /ok=yes/);
assert.match(stdout, /target=production/);
assert.match(stdout, /platforms=2/);
assert.match(stdout, /target_accepted=2/);
assert.match(stdout, /runtime_ready=2/);

const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.type, 'meeting_app_adapter_integration_package_report');
assert.equal(report.ok, true);
assert.equal(report.target, 'production');
assert.equal(report.platform_count, 2);
assert.equal(report.production_ready_count, 2);
assert.equal(report.runtime_ready_count, 2);
assert.equal(report.local_observer_first_count, 2);
assert.equal(report.matrix.packages, undefined);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').first_runtime_route, 'local_observer_axis');
assert.equal(report.written_files.length, 20);

const googleRuntimeDelivery = JSON.parse(await readFile(join(outDir, 'google_meet', 'runtime-delivery.json'), 'utf8'));
assert.equal(googleRuntimeDelivery.schema, 'meeting_app_adapter_runtime_delivery');
assert.equal(googleRuntimeDelivery.adapter_route.first_route, 'local_observer_axis');
assert.equal(googleRuntimeDelivery.host.annotation_timestamp_field, 'captured_at_ms');
assert.equal(googleRuntimeDelivery.readiness.runtime_ready, true);

const googleDescriptor = JSON.parse(await readFile(join(outDir, 'google_meet', 'integration-package.json'), 'utf8'));
assert.equal(googleDescriptor.schema, 'meeting_app_adapter_integration_package');
assert.equal(googleDescriptor.runtime_delivery.runtime_bundle_schema, 'meeting_platform_runtime_bundle');
assert.equal(googleDescriptor.files.some((file) => 'content' in file), false);
assert.equal(googleDescriptor.handoff_package.files.some((file) => 'content' in file), false);
assert.equal(googleDescriptor.written_files.length, 9);

const { stdout: missingStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-app-adapter-integration-package.mjs',
  '--platforms=google-meet',
  '--target=production',
  '--json=true',
], {
  cwd: repoRoot,
});
const missingReport = JSON.parse(missingStdout);
assert.equal(missingReport.ok, false);
assert.equal(missingReport.target_accepted_count, 0);
assert.equal(missingReport.rows[0].target_accepted, false);
assert.equal(missingReport.rows[0].runtime_ready, true);

console.log('ok meeting app adapter integration package script');
