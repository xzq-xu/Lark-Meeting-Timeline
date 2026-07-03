import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-artifact-handoff-script-'));
const inputDir = join(tmpDir, 'input');
const outDir = join(tmpDir, 'handoffs');
const reportFile = join(tmpDir, 'artifact-handoff-report.json');
const inputFile = join(inputDir, 'artifact-signals.json');
const startMs = 1_782_614_400_000;

await mkdir(inputDir, { recursive: true });
await writeFile(inputFile, JSON.stringify({
  signals: [
    {
      type: 'artifact_ready',
      meeting: { platform: 'google_meet', meeting_id: 'google-record-001' },
      occurred_at_ms: startMs,
      artifact_kind: 'transcript',
      artifact_id: 'transcript-1',
    },
    {
      type: 'artifact_ready',
      meeting: { platform: 'zoom', meeting_id: '987654321' },
      occurred_at_ms: startMs + 60_000,
      artifact_kind: 'transcript',
      artifact_url: 'https://zoom.us/recording/transcript.vtt',
    },
    {
      type: 'artifact_ready',
      meeting: { platform: 'webex', meeting_id: 'webex-meeting-001' },
      occurred_at_ms: startMs + 90_000,
      artifact_kind: 'recording',
      artifact_url: 'https://webex.example/recording.mp4',
    },
  ],
}, null, 2), 'utf8');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-artifact-handoff.mjs',
  '--base-url=https://timeline.example.com',
  '--platforms=google-meet,zoom,webex,lark',
  `--signals-file=${inputFile}`,
  `--report-file=${reportFile}`,
  `--out-dir=${outDir}`,
], {
  cwd: repoRoot,
});

assert.match(stdout, /meeting_platform_artifact_handoff_report/);
assert.match(stdout, /platforms=4/);
assert.match(stdout, /fetch_requests=3/);
assert.match(stdout, /realtime_blocking=0/);
const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.type, 'meeting_platform_artifact_handoff_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 4);
assert.equal(report.fetch_request_count, 3);
assert.equal(report.transcript_import_count, 2);
assert.equal(report.recording_count, 1);
assert.equal(report.transcript_supported_count, 4);
assert.equal(report.realtime_blocking_count, 0);
assert.equal(report.rows.find((row) => row.platform === 'lark').status, 'no_artifact_signals');
assert.equal(report.written_files.length, 4);

const googleHandoff = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googleHandoff.schema, 'meeting_platform_artifact_handoff');
assert.equal(googleHandoff.fetch_request_count, 1);
assert.match(googleHandoff.fetch_requests[0].url, /meet.googleapis.com/);
assert.equal(googleHandoff.transcript_blocks_realtime, false);

const { stdout: jsonStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-artifact-handoff.mjs',
  '--platforms=lark',
  '--json=true',
], {
  cwd: repoRoot,
});
const emptyReport = JSON.parse(jsonStdout);
assert.equal(emptyReport.ok, true);
assert.equal(emptyReport.platform_count, 1);
assert.equal(emptyReport.fetch_request_count, 0);
assert.equal(emptyReport.rows[0].status, 'no_artifact_signals');

console.log('ok meeting platform artifact handoff script');
