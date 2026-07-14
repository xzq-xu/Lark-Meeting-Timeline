import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tempDir = await mkdtemp(join(tmpdir(), 'sdk-release-acceptance-test-'));
const reportFile = join(tempDir, 'report.json');

try {
  const { stdout } = await execFileAsync(process.execPath, [
    'scripts/meeting-timeline-sdk-release-acceptance.mjs',
    '--platforms=google-meet,lark',
    '--runs=1',
    '--annotations=5',
    `--report-file=${reportFile}`,
    '--json=true',
  ], { cwd: repoRoot });
  const report = JSON.parse(stdout);
  assert.equal(report.schema, 'meeting_timeline_sdk_release_acceptance_report');
  assert.equal(report.scope, 'deterministic_sdk_contract_and_transport');
  assert.equal(report.field_pilot_claimed, false);
  assert.equal(report.device_required, false);
  assert.equal(report.accepted, true);
  assert.equal(report.accepted_platform_count, 2);
  assert.equal(report.accepted_run_count, 2);
  assert.equal(report.platforms.every((platform) => platform.runs[0].annotation_count === 5), true);
  assert.equal(report.platforms.every((platform) => platform.runs[0].speaker_count === 2), true);
  assert.equal(report.platforms.every((platform) => platform.runs[0].duplicate_visible_annotation_count === 0), true);
  assert.equal(report.platforms[0].runs[0].checks.find((check) => check.id === 'retry_keeps_stable_id').passed, true);
  assert.deepEqual(JSON.parse(await readFile(reportFile, 'utf8')), report);
} finally {
  await rm(tempDir, { recursive: true, force: true });
}

console.log('ok meeting timeline SDK release acceptance script');
