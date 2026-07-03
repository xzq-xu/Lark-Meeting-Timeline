import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-clock-sync-script-'));
const inputDir = join(tmpDir, 'input');
const outDir = join(tmpDir, 'clock-sync');
const reportFile = join(tmpDir, 'clock-sync-report.json');
const inputFile = join(inputDir, 'clock-input.json');
const startMs = 1_782_614_400_000;

await mkdir(inputDir, { recursive: true });
await writeFile(inputFile, JSON.stringify({
  google_meet: {
    samples: [
      {
        id: 'google-best',
        client_send_at_ms: startMs,
        server_time_ms: startMs + 60,
        client_receive_at_ms: startMs + 100,
      },
    ],
    annotation: {
      id: 'google-note',
      label: 'why?',
      captured_at_ms: startMs + 1_000,
    },
  },
  zoom: {
    clock_sync: {
      offset_ms: 250,
      rtt_ms: 40,
    },
    annotation: {
      id: 'zoom-note',
      label: 'follow up',
      captured_at_ms: startMs + 2_000,
    },
  },
  webex: {
    samples: [
      {
        client_send_at_ms: startMs,
        server_time_ms: startMs + 2_000,
        client_receive_at_ms: startMs + 4_000,
      },
    ],
  },
}, null, 2), 'utf8');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-clock-sync.mjs',
  '--platforms=google-meet,zoom,webex,lark',
  `--input=${inputFile}`,
  `--report-file=${reportFile}`,
  `--out-dir=${outDir}`,
], {
  cwd: repoRoot,
});

assert.match(stdout, /meeting_platform_clock_sync_report/);
assert.match(stdout, /platforms=4/);
assert.match(stdout, /accepted=2/);
assert.match(stdout, /ready=2/);
assert.match(stdout, /unstable=1/);
assert.match(stdout, /missing=1/);
const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.type, 'meeting_platform_clock_sync_report_bundle');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 4);
assert.equal(report.accepted_count, 2);
assert.equal(report.ready_count, 2);
assert.equal(report.unstable_count, 1);
assert.equal(report.missing_count, 1);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').recommended_offset_ms, 10);
assert.equal(report.rows.find((row) => row.platform === 'zoom').recommended_offset_ms, 250);
assert.equal(report.rows.find((row) => row.platform === 'webex').status, 'clock_sync_unstable');
assert.equal(report.rows.find((row) => row.platform === 'lark').status, 'clock_sync_missing');
assert.equal(report.written_files.length, 4);

const googleReport = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googleReport.schema, 'meeting_platform_clock_sync_report');
assert.equal(googleReport.calibrated_annotation.captured_at_ms, startMs + 1_010);

const zoomReport = JSON.parse(await readFile(join(outDir, 'zoom.json'), 'utf8'));
assert.equal(zoomReport.samples[0].source, 'direct_offset');
assert.equal(zoomReport.calibrated_annotation.captured_at_ms, startMs + 2_250);

const { stdout: jsonStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-clock-sync.mjs',
  '--platforms=lark',
  '--json=true',
], {
  cwd: repoRoot,
});
const emptyReport = JSON.parse(jsonStdout);
assert.equal(emptyReport.ok, true);
assert.equal(emptyReport.platform_count, 1);
assert.equal(emptyReport.missing_count, 1);
assert.equal(emptyReport.rows[0].status, 'clock_sync_missing');

console.log('ok meeting platform clock sync script');
