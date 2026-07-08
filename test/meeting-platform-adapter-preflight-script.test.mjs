import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

import {
  buildMeetingAppFixtureSnapshot,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adapter-preflight-script-'));
const matrixInputFile = join(tmpDir, 'matrix-input.json');
const candidateInputFile = join(tmpDir, 'candidate-input.json');
const reportFile = join(tmpDir, 'adapter-preflight-report.json');
const outFile = join(tmpDir, 'adapter-preflight-result.json');
const baseUrl = 'https://timeline.example.com';

const googleActive = buildMeetingAppFixtureSnapshot('google-meet', {
  state: 'active',
  observedAtMs: 1_783_356_000_000,
});
await writeFile(matrixInputFile, JSON.stringify({
  snapshots: {
    'google-meet': [googleActive],
  },
  inputByPlatform: {
    teams: {
      platform: 'teams',
      window: {
        id: 'teams-native-main',
        title: 'Microsoft Teams Meeting',
        active: true,
        visible: true,
        inMeeting: true,
        meeting_id: 'teams-123',
      },
      process: { name: 'Microsoft Teams' },
      audio: { call_active: true },
    },
  },
}, null, 2), 'utf8');

await writeFile(candidateInputFile, JSON.stringify({
  windows: [{
    id: 'browser-main',
    tabs: [{
      id: 7,
      url: 'https://meet.google.com/abc-defg-hij',
      title: 'Design review - Google Meet',
      active: true,
      snapshots: [googleActive],
    }],
  }],
}, null, 2), 'utf8');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-preflight.mjs',
  `--base-url=${baseUrl}`,
  '--mode=matrix',
  '--platforms=google-meet,teams',
  `--input-file=${matrixInputFile}`,
  `--report-file=${reportFile}`,
  `--out-file=${outFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_adapter_preflight_report');
assert.equal(report.ok, true);
assert.equal(report.mode, 'matrix');
assert.equal(report.platform_count, 2);
assert.equal(report.accepted_count, 2);
assert.equal(report.realtime_ready_count, 2);
assert.equal(report.live_evidence_ready_count, 2);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').selected_surface, 'browser_extension');
assert.equal(report.result.preflights, undefined);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.accepted_count, 2);
const writtenResult = JSON.parse(await readFile(outFile, 'utf8'));
assert.equal(writtenResult.schema, 'meeting_platform_adapter_preflight_matrix');
assert.equal(writtenResult.preflights, undefined);

const { stdout: candidateStdout } = await execFileAsync(process.execPath, [
  'packages/meeting-timeline-sdk/bin/meeting-platform-adapter-preflight.mjs',
  `--base-url=${baseUrl}`,
  '--mode=candidates',
  `--input-file=${candidateInputFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});
const candidateReport = JSON.parse(candidateStdout);
assert.equal(candidateReport.ok, true);
assert.equal(candidateReport.mode, 'candidates');
assert.equal(candidateReport.candidate_count, 1);
assert.equal(candidateReport.accepted_count, 1);
assert.equal(candidateReport.selected_platform, 'google_meet');
assert.equal(candidateReport.selection_strategy, 'score_accepted_live_current_window_then_active_candidate');
assert.equal(candidateReport.selected_candidate_index, 0);
assert.equal(candidateReport.selected_candidate_score > 0, true);
assert.equal(candidateReport.selected_candidate_reason, 'accepted_live_candidate');
assert.equal(candidateReport.rows[0].realtime_annotation_ready, true);
assert.equal(candidateReport.rows[0].selection_rank, 1);
assert.equal(candidateReport.rows[0].selection_score, candidateReport.selected_candidate_score);

const { stdout: candidateTextStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-preflight.mjs',
  `--base-url=${baseUrl}`,
  '--mode=candidates',
  `--input-file=${candidateInputFile}`,
], {
  cwd: repoRoot,
});
assert.match(candidateTextStdout, /rank=1/);
assert.match(candidateTextStdout, /score=\d+/);
assert.match(candidateTextStdout, /reason=accepted_live_candidate/);

let strictStdout = '';
let strictCode = 0;
try {
  const result = await execFileAsync(process.execPath, [
    'scripts/meeting-platform-adapter-preflight.mjs',
    '--mode=single',
    '--platform=google-meet',
    '--url=https://meet.google.com/abc-defg-hij',
    '--fail-on-rejected=true',
  ], {
    cwd: repoRoot,
  });
  strictStdout = result.stdout;
} catch (error) {
  strictStdout = error.stdout;
  strictCode = error.code;
}
assert.equal(strictCode, 2);
assert.match(strictStdout, /ok=no/);
assert.match(strictStdout, /needs_live_page_evidence/);

console.log('ok meeting platform adapter preflight script');
