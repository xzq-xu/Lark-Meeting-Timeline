import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-app-adapter-verify-'));
const specDir = join(tmpDir, 'specs');
const evidenceDir = join(tmpDir, 'evidence');
const reportFile = join(tmpDir, 'report.json');
await mkdir(specDir, { recursive: true });
await mkdir(evidenceDir, { recursive: true });

const wherebySpec = {
  adapter_key: 'whereby',
  display_name: 'Whereby',
  matches: ['https://whereby.com/*'],
  host_permissions: ['https://whereby.com/*'],
  control_selectors: ['[aria-label*="Leave" i]', '[data-testid*="toolbar" i]'],
  participant_selectors: ['[data-participant-id]', '[aria-label*="speaking" i]'],
  text_selectors: ['[role="status"]', '[aria-live]'],
  mutation_track_selectors: ['[data-participant-id]', '[role="status"]'],
};
const completeEvidence = (adapterKey) => ({
  adapter_key: adapterKey,
  live_dom_snapshot_count: 1,
  candidate_observation_count: 1,
  speaker_segments: [{ captured_at_ms: 1_782_614_401_000, speaker_label: 'Ada' }],
  participant_segments: [{ captured_at_ms: 1_782_614_401_000, participant_label: 'Ada' }],
  annotations: [{
    captured_at_ms: 1_782_614_402_000,
    meeting_id: `${adapterKey}-meeting`,
    axis_id: `${adapterKey}-axis`,
  }],
});
await writeFile(join(specDir, 'whereby.json'), `${JSON.stringify(wherebySpec, null, 2)}\n`, 'utf8');
await writeFile(join(evidenceDir, 'whereby.json'), `${JSON.stringify(completeEvidence('whereby'), null, 2)}\n`, 'utf8');
await writeFile(join(evidenceDir, 'google_meet.json'), `${JSON.stringify(completeEvidence('google_meet'), null, 2)}\n`, 'utf8');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-app-adapter-verify.mjs',
  '--platforms=google-meet',
  `--spec-dir=${specDir}`,
  `--evidence-dir=${evidenceDir}`,
  `--report-file=${reportFile}`,
], {
  cwd: repoRoot,
});

assert.match(stdout, /meeting_app_adapter_verification_report/);
assert.match(stdout, /ok=yes/);
assert.match(stdout, /accepted=2\/2/);
assert.match(stdout, /pilot_ready=2/);
assert.match(stdout, /production_ready=2/);
assert.match(stdout, /missing_evidence=0/);

const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.type, 'meeting_app_adapter_verification_cli_report');
assert.equal(report.ok, true);
assert.equal(report.report_count, 2);
assert.equal(report.accepted_count, 2);
assert.equal(report.matrix.reports, undefined);
assert.equal(report.rows.find((row) => row.adapter_key === 'whereby').production_ready, true);
assert.equal(report.rows.find((row) => row.adapter_key === 'google_meet').production_ready, true);

const { stdout: missingStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-app-adapter-verify.mjs',
  '--platforms=google-meet',
  '--json=true',
], {
  cwd: repoRoot,
});
const missingReport = JSON.parse(missingStdout);
assert.equal(missingReport.ok, false);
assert.equal(missingReport.rows[0].missing_evidence_count, 5);

console.log('ok meeting app adapter verify script');
