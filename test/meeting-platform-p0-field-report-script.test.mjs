import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-p0-field-report-'));
const inputDir = join(tmpDir, 'evidence');
const reportFile = join(tmpDir, 'report.json');
await mkdir(inputDir, { recursive: true });

function runEvidence(index, overrides = {}) {
  const startMs = 1_783_800_000_000 + (index * 300_000);
  const annotations = Array.from({ length: 5 }, (_, markIndex) => ({
    id: `run-${index}-mark-${markIndex}`,
    source: 'hanwang_epaper',
    kind: 'handwriting_trigger',
    intent: 'question',
    captured_at_ms: startMs + 10_000 + (markIndex * 5_000),
    visible_at_ms: startMs + 10_120 + (markIndex * 5_000),
    visible_latency_ms: 120,
    expected_position_ms: 10_000 + (markIndex * 5_000),
    actual_position_ms: 10_000 + (markIndex * 5_000),
    timeline_error_ms: 0,
    duplicate_delivery: false,
  }));
  return {
    schema: 'meeting_platform_field_evidence_input',
    schema_version: 1,
    platform: 'google_meet',
    meeting_id: `google-real-${index}`,
    run_id: `google-meet-run-${index}`,
    meeting: {
      platform: 'google_meet',
      meeting_id: `google-real-${index}`,
      start_time: new Date(startMs).toISOString(),
      end_time: new Date(startMs + 120_000).toISOString(),
      source: 'open_meeting_session',
    },
    meetingAppRecords: [
      {
        id: `active-${index}`,
        platform: 'google_meet',
        phase: 'active',
        captured_at_ms: startMs + 1_000,
        source: 'meeting_app_extension_auto_evidence',
        snapshot: { platform: 'google_meet', title: 'Real Google Meet', source: 'meeting_app_extension_auto_evidence' },
      },
      {
        id: `ended-${index}`,
        platform: 'google_meet',
        phase: 'ended',
        captured_at_ms: startMs + 120_000,
        source: 'meeting_app_extension_auto_evidence',
        snapshot: { platform: 'google_meet', title: 'Google Meet ended', source: 'meeting_app_extension_auto_evidence' },
      },
    ],
    annotations,
    speaker_markers: [
      {
        id: `speaker-a-${index}`,
        kind: 'speaker_started',
        speaker_id: 'speaker-a',
        captured_at_ms: startMs + 30_000,
        visible_at_ms: startMs + 30_700,
        visible_latency_ms: 700,
      },
      {
        id: `speaker-b-${index}`,
        kind: 'speaker_started',
        speaker_id: 'speaker-b',
        captured_at_ms: startMs + 60_000,
        visible_at_ms: startMs + 60_700,
        visible_latency_ms: 700,
      },
    ],
    participant_markers: [],
    measurements: {
      active_observer_captured_at_ms: startMs + 1_000,
      active_axis_visible_at_ms: startMs + 1_100,
      operator_join_at_ms: startMs,
      start_detection_latency_ms: 1_100,
      operator_leave_at_ms: startMs + 115_000,
      ended_axis_visible_at_ms: startMs + 120_000,
      end_detection_latency_ms: 5_000,
      previous_meeting_annotation_count_on_new_axis: 0,
    },
    ...overrides,
  };
}

for (let index = 1; index <= 3; index += 1) {
  await writeFile(join(inputDir, `run-${index}.json`), `${JSON.stringify(runEvidence(index), null, 2)}\n`, 'utf8');
}

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-p0-field-report.mjs',
  '--platforms=google-meet',
  `--dir=${inputDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
  '--fail-on-incomplete=true',
], { cwd: repoRoot });

const report = JSON.parse(stdout);
assert.equal(report.schema, 'meeting_platform_p0_field_report');
assert.equal(report.accepted, true);
assert.equal(report.accepted_platform_count, 1);
assert.equal(report.platforms[0].status, 'p0_field_accepted');
assert.equal(report.platforms[0].accepted_run_count, 3);
assert.equal(report.platforms[0].latest_runs[0].annotation_visible_p95_ms, 120);
assert.equal(report.platforms[0].latest_runs[0].distinct_speaker_count, 2);
assert.equal(report.platforms[0].latest_runs[0].speaker_marker_visible_max_ms, 700);
assert.equal(report.platforms[0].latest_runs[0].observer_surface, 'browser_extension');

const written = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(written.accepted, true);

const retriedEvidence = runEvidence(4);
retriedEvidence.annotations[0].duplicate_delivery = true;
retriedEvidence.annotations[0].delivery_attempt_count = 2;
retriedEvidence.annotations[0].visible_instance_count = 1;
retriedEvidence.annotations[0].duplicate_visible = false;
await writeFile(join(inputDir, 'run-4.json'), `${JSON.stringify(retriedEvidence, null, 2)}\n`, 'utf8');
const { stdout: retriedStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-p0-field-report.mjs',
  '--platforms=google-meet',
  `--dir=${inputDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
], { cwd: repoRoot });
const retried = JSON.parse(retriedStdout);
assert.equal(retried.accepted, true);
assert.equal(retried.platforms[0].latest_runs.at(-1).duplicate_annotation_count, 0);

await writeFile(join(inputDir, 'run-4.json'), `${JSON.stringify(runEvidence(4, {
  speaker_markers: [],
}), null, 2)}\n`, 'utf8');

const { stdout: failedStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-p0-field-report.mjs',
  '--platforms=google-meet',
  `--dir=${inputDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
], { cwd: repoRoot });
const failed = JSON.parse(failedStdout);
assert.equal(failed.accepted, false);
assert.equal(failed.platforms[0].status, 'latest_consecutive_runs_failed');
assert.equal(failed.platforms[0].failed_check_ids.includes('speaker_switches'), true);
assert.equal(failed.platforms[0].failed_check_ids.includes('speaker_marker_visible_latency'), true);

const slowSpeakerEvidence = runEvidence(4);
slowSpeakerEvidence.speaker_markers[1].visible_at_ms += 1_000;
slowSpeakerEvidence.speaker_markers[1].visible_latency_ms = 1_700;
await writeFile(join(inputDir, 'run-4.json'), `${JSON.stringify(slowSpeakerEvidence, null, 2)}\n`, 'utf8');

const { stdout: slowSpeakerStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-p0-field-report.mjs',
  '--platforms=google-meet',
  `--dir=${inputDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
], { cwd: repoRoot });
const slowSpeaker = JSON.parse(slowSpeakerStdout);
assert.equal(slowSpeaker.accepted, false);
assert.equal(slowSpeaker.platforms[0].failed_check_ids.includes('speaker_switches'), false);
assert.equal(slowSpeaker.platforms[0].failed_check_ids.includes('speaker_marker_visible_latency'), true);

console.log('ok meeting platform p0 field report script');
