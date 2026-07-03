import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-speaker-track-script-'));
const inputDir = join(tmpDir, 'input');
const outDir = join(tmpDir, 'tracks');
const reportFile = join(tmpDir, 'speaker-track-report.json');
const inputFile = join(inputDir, 'speaker-samples.json');
const startMs = 1_782_614_400_000;

await mkdir(inputDir, { recursive: true });
await writeFile(inputFile, JSON.stringify({
  samples: {
    google_meet: [
      {
        meeting: {
          platform: 'google_meet',
          meeting_id: 'abc-defg-hij',
          meeting_url: 'https://meet.google.com/abc-defg-hij',
        },
        activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
        observedAtMs: startMs,
      },
      {
        meeting: {
          platform: 'google_meet',
          meeting_id: 'abc-defg-hij',
          meeting_url: 'https://meet.google.com/abc-defg-hij',
        },
        activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
        observedAtMs: startMs + 300,
      },
      {
        meeting: {
          platform: 'google_meet',
          meeting_id: 'abc-defg-hij',
          meeting_url: 'https://meet.google.com/abc-defg-hij',
        },
        observedAtMs: startMs + 1_400,
      },
      {
        meeting: {
          platform: 'google_meet',
          meeting_id: 'abc-defg-hij',
          meeting_url: 'https://meet.google.com/abc-defg-hij',
        },
        observedAtMs: startMs + 1_950,
      },
    ],
  },
  signals: {
    zoom: [
      {
        type: 'speaker_started',
        meeting: { platform: 'zoom', meeting_id: '987654321' },
        occurred_at_ms: startMs + 5_000,
        speaker_name: 'Grace',
      },
      {
        type: 'speaker_ended',
        meeting: { platform: 'zoom', meeting_id: '987654321' },
        occurred_at_ms: startMs + 7_000,
        speaker_name: 'Grace',
      },
    ],
  },
}, null, 2), 'utf8');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-speaker-track.mjs',
  '--platforms=google-meet,zoom,lark',
  `--samples-file=${inputFile}`,
  `--report-file=${reportFile}`,
  `--out-dir=${outDir}`,
  '--min-stable-ms=250',
  '--end-idle-ms=500',
  '--min-segment-ms=800',
], {
  cwd: repoRoot,
});

assert.match(stdout, /meeting_platform_speaker_track_report/);
assert.match(stdout, /platforms=3/);
assert.match(stdout, /marks=2/);
const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.type, 'meeting_platform_speaker_track_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.mark_count, 2);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').mark_count, 1);
assert.equal(report.rows.find((row) => row.platform === 'zoom').mark_count, 1);
assert.equal(report.rows.find((row) => row.platform === 'lark').mark_count, 0);
assert.equal(report.written_files.length, 3);

const googleTrack = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googleTrack.schema, 'meeting_platform_speaker_track');
assert.equal(googleTrack.marks[0].intent, 'speaker_track');
assert.equal(googleTrack.marks[0].payload.segment.duration_ms, 1_400);
assert.equal(JSON.stringify(googleTrack.marks).includes('transcript'), false);

const { stdout: jsonStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-speaker-track.mjs',
  '--platforms=lark',
  '--json=true',
], {
  cwd: repoRoot,
});
const emptyReport = JSON.parse(jsonStdout);
assert.equal(emptyReport.ok, true);
assert.equal(emptyReport.mark_count, 0);
assert.equal(emptyReport.rows[0].status, 'no_speaker_markers');

console.log('ok meeting platform speaker track script');
