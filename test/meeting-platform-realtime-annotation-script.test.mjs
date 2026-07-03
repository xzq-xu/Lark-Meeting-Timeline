import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-realtime-annotation-script-'));
const inputDir = join(tmpDir, 'input');
const outDir = join(tmpDir, 'pipelines');
const reportFile = join(tmpDir, 'realtime-annotation-report.json');
const inputFile = join(inputDir, 'pipeline-input.json');
const startMs = 1_782_614_400_000;

await mkdir(inputDir, { recursive: true });
await writeFile(inputFile, JSON.stringify({
  google_meet: {
    clock_sync: {
      offset_ms: 10,
      rtt_ms: 40,
    },
    current_meeting: {
      platform: 'google_meet',
      meeting_id: 'abc-defg-hij',
      meeting_url: 'https://meet.google.com/abc-defg-hij',
      start_time_ms: startMs,
    },
    local_observer: {
      url: 'https://meet.google.com/abc-defg-hij',
      observed_at_ms: startMs + 1_000,
    },
    annotation: {
      id: 'google-note',
      label: 'why?',
      captured_at_ms: startMs + 90_000,
    },
  },
  zoom: {
    clock_sync: {
      offset_ms: 0,
      rtt_ms: 40,
    },
    local_observer: {
      url: 'https://zoom.us/j/987654321',
      observed_at_ms: startMs + 1_000,
    },
    annotation: {
      id: 'zoom-note',
      label: 'follow up',
      captured_at_ms: startMs + 120_000,
    },
  },
  lark: {
    clock_sync: {
      offset_ms: 0,
      rtt_ms: 40,
    },
    annotation: {
      id: 'lark-note',
      label: 'why?',
      captured_at_ms: startMs + 3_000,
    },
  },
  webex: {
    clock_sync: {
      offset_ms: 0,
      rtt_ms: 40,
    },
    current_meeting: {
      platform: 'webex',
      meeting_id: 'webex-a',
      start_time_ms: startMs,
    },
    signals: [
      {
        type: 'meeting_started',
        meeting: {
          platform: 'webex',
          meeting_id: 'webex-b',
        },
        occurred_at_ms: startMs + 1_000,
      },
    ],
    annotation: {
      id: 'webex-note',
      label: 'conflict',
      captured_at_ms: startMs + 3_000,
    },
  },
}, null, 2), 'utf8');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-realtime-annotation.mjs',
  '--platforms=google-meet,zoom,lark,webex',
  `--input=${inputFile}`,
  `--report-file=${reportFile}`,
  `--out-dir=${outDir}`,
], {
  cwd: repoRoot,
});

assert.match(stdout, /meeting_platform_realtime_annotation_report/);
assert.match(stdout, /platforms=4/);
assert.match(stdout, /accepted=3/);
assert.match(stdout, /ready=1/);
assert.match(stdout, /start_axis=1/);
assert.match(stdout, /pending=1/);
assert.match(stdout, /conflicts=1/);
const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.type, 'meeting_platform_realtime_annotation_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 4);
assert.equal(report.accepted_count, 3);
assert.equal(report.ready_count, 1);
assert.equal(report.start_axis_count, 1);
assert.equal(report.pending_count, 1);
assert.equal(report.conflict_count, 1);
assert.equal(report.needs_clock_count, 0);
assert.equal(report.provider_blocking_count, 0);
assert.equal(report.transcript_blocking_count, 0);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').status, 'ready_to_insert');
assert.equal(report.rows.find((row) => row.platform === 'zoom').status, 'start_axis_then_insert');
assert.equal(report.rows.find((row) => row.platform === 'lark').status, 'pending_real_meeting');
assert.equal(report.rows.find((row) => row.platform === 'webex').status, 'binding_conflict');
assert.equal(report.written_files.length, 4);

const googlePipeline = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googlePipeline.schema, 'meeting_platform_realtime_annotation');
assert.equal(googlePipeline.insert_payload.time_ms, 90_010);
assert.equal(googlePipeline.actions.includes('insert_mark'), true);

const zoomPipeline = JSON.parse(await readFile(join(outDir, 'zoom.json'), 'utf8'));
assert.equal(zoomPipeline.start_payload.meeting_id, '987654321');
assert.deepEqual(zoomPipeline.actions, ['start_meeting_session', 'insert_mark']);

const { stdout: jsonStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-realtime-annotation.mjs',
  '--platforms=google-meet',
  '--json=true',
], {
  cwd: repoRoot,
});
const emptyReport = JSON.parse(jsonStdout);
assert.equal(emptyReport.ok, true);
assert.equal(emptyReport.platform_count, 1);
assert.equal(emptyReport.needs_clock_count, 1);
assert.equal(emptyReport.rows[0].status, 'needs_clock_sync');

console.log('ok meeting platform realtime annotation script');
