import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const tmp = mkdtempSync(join(tmpdir(), 'meeting-platform-timeline-view-'));
const inputFile = join(tmp, 'timeline-input.json');
const reportFile = join(tmp, 'timeline-report.json');
const outDir = join(tmp, 'views');
const startMs = 1_782_614_400_000;

writeFileSync(inputFile, JSON.stringify({
  google_meet: {
    meeting: {
      platform: 'google_meet',
      meeting_id: 'abc-defg-hij',
      start_time_ms: startMs,
      end_time_ms: startMs + 600_000,
    },
    annotations: [
      {
        id: 'note-1',
        label: 'why?',
        captured_at_ms: startMs + 90_000,
      },
    ],
    speakerTrack: {
      marks: [
        {
          id: 'speaker-1',
          intent: 'speaker_track',
          label: 'Ada speaking',
          captured_at_ms: startMs + 120_000,
        },
      ],
    },
  },
  zoom: {
    meeting: {
      platform: 'zoom',
      meeting_id: '987654321',
      start_time_ms: startMs,
      duration_ms: 600_000,
    },
    participantTrack: {
      marks: [
        {
          id: 'participant-1',
          intent: 'participant_track',
          label: 'Grace joined',
          captured_at_ms: startMs + 180_000,
        },
      ],
    },
  },
}, null, 2));

const stdout = execFileSync('node', [
  'scripts/meeting-platform-timeline-view.mjs',
  '--platforms=google-meet,zoom,lark',
  `--input=${inputFile}`,
  `--report-file=${reportFile}`,
  `--out-dir=${outDir}`,
  '--viewport-start-ms=60000',
  '--viewport-duration-ms=240000',
], {
  cwd: process.cwd(),
  encoding: 'utf8',
});
assert.match(stdout, /meeting_platform_timeline_view_report/);
assert.match(stdout, /platforms=3/);
assert.match(stdout, /markers=3/);
assert.match(stdout, /visible=3/);

const report = JSON.parse(readFileSync(reportFile, 'utf8'));
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.marker_count, 3);
assert.equal(report.visible_marker_count, 3);
assert.equal(report.written_files.length, 3);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').marker_count, 2);
assert.equal(report.rows.find((row) => row.platform === 'zoom').marker_count, 1);
assert.equal(report.rows.find((row) => row.platform === 'lark').status, 'empty_timeline_view');

const googleView = JSON.parse(readFileSync(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googleView.schema, 'meeting_platform_timeline_view');
assert.equal(googleView.viewport.start_ms, 60_000);
assert.equal(googleView.viewport.end_ms, 300_000);
assert.equal(googleView.diagnostics.rail_counts.annotations, 1);
assert.equal(googleView.diagnostics.rail_counts.speaker, 1);
assert.deepEqual(googleView.visible_markers.map((marker) => marker.id), ['note-1', 'speaker-1']);

const jsonStdout = execFileSync('node', [
  'scripts/meeting-platform-timeline-view.mjs',
  '--platforms=lark',
  '--json=true',
], {
  cwd: process.cwd(),
  encoding: 'utf8',
});
const jsonReport = JSON.parse(jsonStdout);
assert.equal(jsonReport.ok, true);
assert.equal(jsonReport.platform_count, 1);
assert.equal(jsonReport.marker_count, 0);
assert.equal(jsonReport.rows[0].status, 'empty_timeline_view');

console.log('ok meeting platform timeline view script');
