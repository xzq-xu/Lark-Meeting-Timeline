import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const tmp = mkdtempSync(join(tmpdir(), 'meeting-platform-participant-track-'));
const inputFile = join(tmp, 'participant-input.json');
const reportFile = join(tmp, 'participant-report.json');
const outDir = join(tmp, 'tracks');
const startMs = 1_782_614_400_000;

writeFileSync(inputFile, JSON.stringify({
  snapshots: {
    google_meet: [
      {
        meeting: { platform: 'google_meet', meeting_id: 'abc-defg-hij' },
        observed_at_ms: startMs,
        participants: [{ id: 'ada', name: 'Ada' }],
      },
      {
        meeting: { platform: 'google_meet', meeting_id: 'abc-defg-hij' },
        observed_at_ms: startMs + 1_000,
        participants: [{ id: 'ada', name: 'Ada' }, { id: 'bob', name: 'Bob' }],
      },
      {
        meeting: { platform: 'google_meet', meeting_id: 'abc-defg-hij' },
        observed_at_ms: startMs + 3_000,
        participants: [{ id: 'ada', name: 'Ada' }],
      },
      {
        meeting: { platform: 'google_meet', meeting_id: 'abc-defg-hij' },
        observed_at_ms: startMs + 4_700,
        participants: [{ id: 'ada', name: 'Ada' }],
      },
    ],
  },
  signals: {
    zoom: [
      {
        type: 'participant_joined',
        meeting: { platform: 'zoom', meeting_id: '987654321' },
        occurred_at_ms: startMs,
        participant_id: 'grace',
        participant_name: 'Grace',
      },
      {
        type: 'participant_left',
        meeting: { platform: 'zoom', meeting_id: '987654321' },
        occurred_at_ms: startMs + 8_000,
        participant_id: 'grace',
        participant_name: 'Grace',
      },
    ],
  },
}, null, 2));

const stdout = execFileSync('node', [
  'scripts/meeting-platform-participant-track.mjs',
  '--platforms=google-meet,zoom,lark',
  `--input=${inputFile}`,
  `--report-file=${reportFile}`,
  `--out-dir=${outDir}`,
  '--leave-stable-ms=1000',
], {
  cwd: process.cwd(),
  encoding: 'utf8',
});
assert.match(stdout, /meeting_platform_participant_track_report/);
assert.match(stdout, /platforms=3/);
assert.match(stdout, /marks=4/);

const report = JSON.parse(readFileSync(reportFile, 'utf8'));
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.mark_count, 4);
assert.equal(report.written_files.length, 3);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').mark_count, 2);
assert.equal(report.rows.find((row) => row.platform === 'zoom').mark_count, 2);
assert.equal(report.rows.find((row) => row.platform === 'lark').status, 'no_participant_markers');

const googleTrack = JSON.parse(readFileSync(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googleTrack.schema, 'meeting_platform_participant_track');
assert.equal(googleTrack.mark_count, 2);
assert.deepEqual(googleTrack.marks.map((mark) => mark.label), ['Bob joined', 'Bob left']);
assert.equal(googleTrack.diagnostics.emitted_pending_leave_count, 1);

const jsonStdout = execFileSync('node', [
  'scripts/meeting-platform-participant-track.mjs',
  '--platforms=lark',
  '--json=true',
], {
  cwd: process.cwd(),
  encoding: 'utf8',
});
const jsonReport = JSON.parse(jsonStdout);
assert.equal(jsonReport.ok, true);
assert.equal(jsonReport.platform_count, 1);
assert.equal(jsonReport.mark_count, 0);
assert.equal(jsonReport.rows[0].status, 'no_participant_markers');

console.log('ok meeting platform participant track script');
