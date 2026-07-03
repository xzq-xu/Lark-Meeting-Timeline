import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import { buildMeetingAppLiveEvidencePackage } from '../packages/meeting-timeline-sdk/adapters/meeting-app-profile.mjs';
import { buildMeetingAppSnapshotRecordSet } from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-app-evidence-gate-'));
const evidenceFile = join(tmpDir, 'google-meet-evidence.json');
const reportFile = join(tmpDir, 'gate-report.json');
const observedAtMs = 1_783_356_000_000;

const activeSnapshot = buildMeetingAppFixtureSnapshot('google-meet', {
  observedAtMs,
  state: 'active',
});
const endedSnapshot = buildMeetingAppFixtureSnapshot('google-meet', {
  observedAtMs: observedAtMs + 1_000,
  state: 'prejoin',
});
const recordSet = buildMeetingAppSnapshotRecordSet([
  { snapshot: activeSnapshot, platform: 'google_meet', phase: 'active', capturedAtMs: observedAtMs },
  { snapshot: endedSnapshot, platform: 'google_meet', phase: 'ended', capturedAtMs: observedAtMs + 1_000 },
], {
  id: 'google-meet-live-record-set',
  createdAtMs: observedAtMs + 2_000,
});
const evidencePackage = buildMeetingAppLiveEvidencePackage({
  platforms: ['google-meet'],
  recordSet,
  recordSetId: 'google-meet-live-record-set',
  packageId: 'google-meet-live-evidence',
});
await writeFile(evidenceFile, `${JSON.stringify(evidencePackage, null, 2)}\n`, 'utf8');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-app-evidence-gate.mjs',
  `--input=${evidenceFile}`,
  `--report-file=${reportFile}`,
  '--json=true',
], { cwd: repoRoot });

const summary = JSON.parse(stdout);
assert.equal(summary.type, 'meeting_app_evidence_gate');
assert.equal(summary.ok, true);
assert.equal(summary.production_ready, true);
assert.equal(summary.record_count, 2);
assert.deepEqual(summary.platforms, ['google_meet']);
assert.equal(summary.rows[0].platform, 'google_meet');
assert.equal(summary.rows[0].production_ready, true);
assert.equal(summary.rows[0].evidence_level, 'captured_dom');
assert.deepEqual(summary.rows[0].missing_required_coverage, []);

const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.ok, true);
assert.equal(report.record_set_id, 'google-meet-live-record-set');

console.log('ok meeting app evidence gate script');
