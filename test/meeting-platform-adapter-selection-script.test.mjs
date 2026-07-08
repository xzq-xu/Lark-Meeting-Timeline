import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adapter-selection-script-'));
const reportFile = join(tmpDir, 'selection-report.json');
const outDir = join(tmpDir, 'selections');
const recordSetFile = join(tmpDir, 'zoom-record-set.json');
const baseUrl = 'https://timeline.example.com';

await writeFile(recordSetFile, JSON.stringify({
  platform: 'zoom',
  records: [
    {
      phase: 'active',
      captured_at_ms: 1_783_356_000_000,
      snapshot: { platform: 'zoom', in_meeting: true },
    },
    {
      phase: 'ended',
      captured_at_ms: 1_783_356_060_000,
      snapshot: { platform: 'zoom', in_meeting: false },
    },
  ],
}, null, 2), 'utf8');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-selection.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,zoom,local-detector',
  `--meeting-app-record-set-file=${recordSetFile}`,
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_adapter_selection_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.selection_ready_count, 3);
assert.equal(report.local_axis_selected_count, 3);
assert.equal(report.provider_reconcile_count, 2);
assert.equal(report.post_meeting_artifact_count, 3);
assert.equal(report.pilot_evidence_ready_count, 1);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').axis_source, 'local_observer_axis');
assert.equal(report.rows.find((row) => row.platform === 'google_meet').pilot_evidence_ready, false);
assert.equal(report.rows.find((row) => row.platform === 'zoom').axis_surface, 'native_detector');
assert.equal(report.rows.find((row) => row.platform === 'zoom').pilot_evidence_ready, true);
assert.equal(report.rows.find((row) => row.platform === 'local_detector').axis_source, 'host_detector_axis');
assert.equal(report.written_files.length, 3);
assert.equal(report.matrix.selections, undefined);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 3);
assert.equal(writtenReport.rows.find((row) => row.platform === 'zoom').selection_file.endsWith('zoom.json'), true);
const zoomSelection = JSON.parse(await readFile(join(outDir, 'zoom.json'), 'utf8'));
assert.equal(zoomSelection.selection.axis_source, 'local_observer_axis');
assert.equal(zoomSelection.selection.axis_surface, 'native_detector');
assert.equal(zoomSelection.readiness.pilot_evidence_ready, true);

const { stdout: compactStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-selection.mjs',
  '--platforms=webex',
  '--write-selections=false',
  '--json=true',
], {
  cwd: repoRoot,
});
const compactReport = JSON.parse(compactStdout);
assert.equal(compactReport.platform_count, 1);
assert.equal(compactReport.written_files.length, 0);
assert.equal(compactReport.rows[0].platform, 'webex');
assert.equal(compactReport.rows[0].provider_events_block_realtime, false);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-selection.mjs',
  '--platforms=lark',
  '--write-selections=false',
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_adapter_selection_report/);
assert.match(textStdout, /lark: axis=local_observer_axis/);
assert.match(textStdout, /provider=provider_reconcile/);

console.log('ok meeting platform adapter selection script');
