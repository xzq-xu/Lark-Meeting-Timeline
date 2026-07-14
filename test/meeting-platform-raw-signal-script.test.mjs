import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-raw-signal-script-'));
const inputFile = join(tmpDir, 'raw-signals.json');
const outFile = join(tmpDir, 'raw-signal-batch.json');
const reportFile = join(tmpDir, 'raw-signal-report.json');
const baseMs = 1_782_614_400_000;

await writeFile(inputFile, `${JSON.stringify([
  {
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Weekly sync',
    observed_at_ms: baseMs,
    active_speaker: { id: 'alex', display_name: 'Alex' },
  },
  {
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Weekly sync',
    observed_at_ms: baseMs + 250,
    active_speaker: { id: 'alex', display_name: 'Alex' },
  },
  {
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Weekly sync',
    observed_at_ms: baseMs + 700,
    active_speaker: { id: 'alex', display_name: 'Alex' },
  },
  {
    kind: 'annotation',
    platform: 'google-meet',
    label: 'why?',
    captured_at_ms: baseMs + 1_200,
    current_meeting: {
      meeting_id: 'google_meet-sample-meeting',
      meeting_url: 'https://meet.google.com/abc-defg-hij',
    },
  },
], null, 2)}\n`, 'utf8');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-raw-signal.mjs',
  `--input-file=${inputFile}`,
  `--out-file=${outFile}`,
  `--report-file=${reportFile}`,
  '--filter-active-speaker-samples=true',
  '--min-stable-ms=500',
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_raw_signal_report');
assert.equal(report.ok, true);
assert.equal(report.mode, 'input');
assert.equal(report.signal_count, 4);
assert.equal(report.runtime_event_count, 5);
assert.equal(report.filtered_speaker_event_count, 1);
assert.equal(report.platforms.includes('google_meet'), true);
assert.equal(report.action_counts.observe_meeting_app, 3);
assert.equal(report.action_counts.insert_annotation, 1);
assert.equal(report.action_counts.speaker_track, 1);
assert.equal(report.runtime_events.find((event) => event.action === 'speaker_track').signals[0].speaker_name, 'Alex');

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.runtime_event_count, 5);
const writtenBatch = JSON.parse(await readFile(outFile, 'utf8'));
assert.equal(writtenBatch.schema, 'meeting_platform_raw_signal_batch');
assert.equal(writtenBatch.runtime_events.length, 5);

const { stdout: examplesStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-raw-signal.mjs',
  '--examples=true',
  '--platforms=google-meet,teams,zoom',
  '--json=true',
], {
  cwd: repoRoot,
});
const examplesReport = JSON.parse(examplesStdout);
assert.equal(examplesReport.mode, 'examples');
assert.equal(examplesReport.platform_count, 3);
assert.equal(examplesReport.signal_count, 9);
assert.equal(examplesReport.action_counts.observe_meeting_app, 3);
assert.equal(examplesReport.action_counts.insert_annotation, 3);
assert.equal(examplesReport.action_counts.speaker_track, 6);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-raw-signal.mjs',
  '--examples=true',
  '--platforms=webex',
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_raw_signal_report/);
assert.match(textStdout, /webex:meeting_app_snapshot/);
assert.match(textStdout, /actions=/);

const { stdout: binStdout } = await execFileAsync(process.execPath, [
  'packages/meeting-timeline-sdk/bin/meeting-platform-raw-signal.mjs',
  '--examples=true',
  '--platforms=google-meet,zoom',
  '--json=true',
], {
  cwd: repoRoot,
});
const binReport = JSON.parse(binStdout);
assert.equal(binReport.type, 'meeting_platform_raw_signal_report');
assert.equal(binReport.platform_count, 2);
assert.equal(binReport.signal_count, 6);
assert.equal(binReport.runtime_actions.includes('observe_meeting_app'), true);
assert.equal(binReport.runtime_actions.includes('insert_annotation'), true);

console.log('ok meeting platform raw signal script');
