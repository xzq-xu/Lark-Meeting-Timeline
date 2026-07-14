import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-session-binding-script-'));
const inputDir = join(tmpDir, 'input');
const outDir = join(tmpDir, 'bindings');
const reportFile = join(tmpDir, 'session-binding-report.json');
const inputFile = join(inputDir, 'binding-input.json');
const startMs = 1_782_614_400_000;

await mkdir(inputDir, { recursive: true });
await writeFile(inputFile, JSON.stringify({
  google_meet: {
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
  },
  zoom: {
    local_observer: {
      url: 'https://zoom.us/j/987654321',
      observed_at_ms: startMs + 2_000,
    },
  },
  webex: {
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
  },
  lark: {
    annotation: {
      id: 'lark-note',
      label: 'why?',
      captured_at_ms: startMs + 3_000,
    },
  },
}, null, 2), 'utf8');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-session-binding.mjs',
  '--platforms=google-meet,zoom,webex,lark',
  `--input=${inputFile}`,
  `--report-file=${reportFile}`,
  `--out-dir=${outDir}`,
], {
  cwd: repoRoot,
});

assert.match(stdout, /meeting_platform_session_binding_report/);
assert.match(stdout, /platforms=4/);
assert.match(stdout, /bound=1/);
assert.match(stdout, /start_axis=1/);
assert.match(stdout, /pending=1/);
assert.match(stdout, /conflicts=1/);
const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.type, 'meeting_platform_session_binding_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 4);
assert.equal(report.bound_count, 1);
assert.equal(report.start_axis_count, 1);
assert.equal(report.pending_count, 1);
assert.equal(report.conflict_count, 1);
assert.equal(report.provider_blocking_count, 0);
assert.equal(report.transcript_blocking_count, 0);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').status, 'bound_to_current_axis');
assert.equal(report.rows.find((row) => row.platform === 'zoom').status, 'open_axis_from_local_observer');
assert.equal(report.rows.find((row) => row.platform === 'webex').status, 'binding_conflict');
assert.equal(report.rows.find((row) => row.platform === 'lark').status, 'pending_binding');
assert.equal(report.written_files.length, 4);

const googleBinding = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googleBinding.schema, 'meeting_platform_session_binding');
assert.equal(googleBinding.selected_meeting.meeting_id, 'abc-defg-hij');
assert.equal(googleBinding.bind_to_current_axis, true);

const zoomBinding = JSON.parse(await readFile(join(outDir, 'zoom.json'), 'utf8'));
assert.equal(zoomBinding.start_payload.meeting_id, '987654321');
assert.equal(zoomBinding.start_payload.detector_source, 'local_observer_session_binding');

const { stdout: jsonStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-session-binding.mjs',
  '--platforms=lark',
  '--json=true',
], {
  cwd: repoRoot,
});
const emptyReport = JSON.parse(jsonStdout);
assert.equal(emptyReport.ok, true);
assert.equal(emptyReport.platform_count, 1);
assert.equal(emptyReport.insufficient_count, 1);
assert.equal(emptyReport.rows[0].status, 'insufficient_identity');

console.log('ok meeting platform session binding script');
