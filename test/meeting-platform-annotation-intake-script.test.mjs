import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-annotation-intake-script-'));
const inputDir = join(tmpDir, 'input');
const outDir = join(tmpDir, 'decisions');
const reportFile = join(tmpDir, 'annotation-intake-report.json');
const inputFile = join(inputDir, 'annotation-input.json');
const startMs = 1_782_614_400_000;

await mkdir(inputDir, { recursive: true });
await writeFile(inputFile, JSON.stringify({
  google_meet: {
    current_meeting: {
      platform: 'google_meet',
      meeting_id: 'abc-defg-hij',
      start_time_ms: startMs,
    },
    annotation: {
      id: 'google-note',
      label: 'why?',
      captured_at_ms: startMs + 30_000,
    },
  },
  zoom: {
    current_meeting: {
      platform: 'zoom',
      meeting_id: '987654321',
      start_time_ms: startMs,
      end_time_ms: startMs + 60_000,
    },
    annotation: {
      id: 'zoom-late-note',
      label: 'late',
      captured_at_ms: startMs + 120_000,
    },
  },
  lark: {
    annotation: {
      id: 'lark-note',
      label: 'follow up',
      captured_at_ms: startMs + 90_000,
    },
  },
}, null, 2), 'utf8');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-annotation-intake.mjs',
  '--platforms=google-meet,zoom,lark',
  `--input=${inputFile}`,
  `--report-file=${reportFile}`,
  `--out-dir=${outDir}`,
], {
  cwd: repoRoot,
});

assert.match(stdout, /meeting_platform_annotation_intake_report/);
assert.match(stdout, /platforms=3/);
assert.match(stdout, /accepted=2/);
assert.match(stdout, /ready=1/);
assert.match(stdout, /open_session=1/);
assert.match(stdout, /after_end=1/);
const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.type, 'meeting_platform_annotation_intake_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.accepted_count, 2);
assert.equal(report.ready_count, 1);
assert.equal(report.open_session_count, 1);
assert.equal(report.after_end_count, 1);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').status, 'ready_to_insert_current_axis');
assert.equal(report.rows.find((row) => row.platform === 'zoom').status, 'after_meeting_end');
assert.equal(report.rows.find((row) => row.platform === 'lark').status, 'start_open_session_then_insert');
assert.equal(report.written_files.length, 3);

const googleDecision = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googleDecision.schema, 'meeting_platform_annotation_intake');
assert.equal(googleDecision.insert_payload.time_ms, 30_000);
assert.equal(googleDecision.accepted_for_realtime, true);

const larkDecision = JSON.parse(await readFile(join(outDir, 'lark.json'), 'utf8'));
assert.equal(larkDecision.open_session_payload.meeting_id, `annotation-open-session-${startMs + 90_000}`);
assert.equal(larkDecision.open_session_payload.suppress_auto_annotations, true);

const { stdout: jsonStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-annotation-intake.mjs',
  '--platforms=lark',
  '--json=true',
], {
  cwd: repoRoot,
});
const emptyReport = JSON.parse(jsonStdout);
assert.equal(emptyReport.ok, true);
assert.equal(emptyReport.platform_count, 1);
assert.equal(emptyReport.missing_captured_count, 1);
assert.equal(emptyReport.rows[0].status, 'needs_device_captured_at');

console.log('ok meeting platform annotation intake script');
