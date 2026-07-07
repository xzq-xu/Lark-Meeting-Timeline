import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-provider-replay-script-'));
const larkRecordsFile = join(tmpDir, 'lark-records.json');
const replayMatrixFile = join(tmpDir, 'provider-replay-matrix.json');
const cliReportFile = join(tmpDir, 'provider-replay-report.json');
const baseMs = 1_782_614_400_000;

const { stdout: generatedStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-provider-replay.mjs',
  '--platforms=google-meet,teams,zoom,webex,lark',
  `--base-received-at-ms=${baseMs}`,
  '--json=true',
], {
  cwd: repoRoot,
});
const generatedReport = JSON.parse(generatedStdout);
assert.equal(generatedReport.type, 'meeting_platform_provider_replay_cli_report');
assert.equal(generatedReport.ok, true);
assert.equal(generatedReport.generated_samples, true);
assert.equal(generatedReport.platform_count, 5);
assert.equal(generatedReport.accepted_count, 5);
assert.equal(generatedReport.rows.every((row) => row.coverage.meeting_start === true), true);
assert.equal(generatedReport.rows.every((row) => row.coverage.meeting_end === true), true);
assert.equal(generatedReport.rows.every((row) => row.provider_events_block_realtime === false), true);

await writeFile(larkRecordsFile, `${JSON.stringify({
  platform: 'lark',
  records: [{
    header: {
      event_id: 'lark-provider-replay-start',
      event_type: 'vc.meeting.all_meeting_started_v1',
      create_time: String(baseMs),
    },
    event: {
      meeting: {
        id: 'lark-provider-replay-meeting',
        meeting_no: '123456789',
        topic: 'Provider replay fixture',
        url: 'https://vc.feishu.cn/j/lark-provider-replay-meeting',
        start_time: String(Math.round(baseMs / 1000)),
      },
      minute_token: 'minute-token-provider-replay',
    },
  }],
}, null, 2)}\n`, 'utf8');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-provider-replay.mjs',
  '--platforms=lark',
  `--records-file=${larkRecordsFile}`,
  '--required-coverage=meeting_start',
  `--out-file=${replayMatrixFile}`,
  `--report-file=${cliReportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});
const report = JSON.parse(stdout);
assert.equal(report.ok, true);
assert.equal(report.generated_samples, false);
assert.equal(report.platform_count, 1);
assert.equal(report.accepted_count, 1);
assert.equal(report.rows[0].platform, 'lark');
assert.equal(report.rows[0].coverage.meeting_start, true);
assert.equal(report.rows[0].coverage.meeting_end, false);
assert.equal(report.rows[0].required_coverage.includes('meeting_start'), true);
assert.equal(report.replay_matrix.reports[0].rows[0].meetings[0].minute_token, 'minute-token-provider-replay');

const writtenCliReport = JSON.parse(await readFile(cliReportFile, 'utf8'));
assert.equal(writtenCliReport.ok, true);

const replayMatrix = JSON.parse(await readFile(replayMatrixFile, 'utf8'));
assert.equal(replayMatrix.schema, 'meeting_platform_provider_replay_matrix');
assert.equal(replayMatrix.accepted, true);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-provider-replay.mjs',
  '--platforms=google-meet,zoom',
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_provider_replay_cli_report/);
assert.match(textStdout, /google_meet: accepted=yes/);
assert.match(textStdout, /provider_nonblocking=yes/);
assert.match(textStdout, /zoom: accepted=yes/);

const { stdout: missingStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-provider-replay.mjs',
  `--records-file=${join(tmpDir, 'missing-records.json')}`,
  '--json=true',
], {
  cwd: repoRoot,
});
const missingReport = JSON.parse(missingStdout);
assert.equal(missingReport.ok, false);
assert.match(missingReport.records_read_error, /ENOENT/);

console.log('ok meeting platform provider replay script');
