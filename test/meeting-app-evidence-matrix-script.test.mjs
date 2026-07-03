import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import { buildMeetingAppLiveEvidencePackage } from '../packages/meeting-timeline-sdk/adapters/meeting-app-profile.mjs';
import { buildMeetingAppSnapshotRecordSet } from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-app-evidence-matrix-'));
const evidenceDir = join(tmpDir, 'evidence');
const reportFile = join(tmpDir, 'matrix-report.json');
const observedAtMs = 1_783_356_000_000;
await mkdir(evidenceDir, { recursive: true });

function liveEvidence(platform) {
  const activeSnapshot = buildMeetingAppFixtureSnapshot(platform, {
    observedAtMs,
    state: 'active',
  });
  const endedSnapshot = buildMeetingAppFixtureSnapshot(platform, {
    observedAtMs: observedAtMs + 1_000,
    state: 'prejoin',
  });
  const platformKey = platform.replaceAll('-', '_');
  const recordSet = buildMeetingAppSnapshotRecordSet([
    { snapshot: activeSnapshot, platform: platformKey, phase: 'active', capturedAtMs: observedAtMs },
    { snapshot: endedSnapshot, platform: platformKey, phase: 'ended', capturedAtMs: observedAtMs + 1_000 },
  ], {
    id: `${platformKey}-record-set`,
    createdAtMs: observedAtMs + 2_000,
  });
  return buildMeetingAppLiveEvidencePackage({
    platforms: [platform],
    recordSet,
    recordSetId: `${platformKey}-record-set`,
    packageId: `${platformKey}-evidence`,
  });
}

await writeFile(join(evidenceDir, 'google-meet.json'), `${JSON.stringify(liveEvidence('google-meet'), null, 2)}\n`, 'utf8');
await writeFile(join(evidenceDir, 'zoom-empty.json'), JSON.stringify({
  platforms: ['zoom'],
  records: [],
}, null, 2), 'utf8');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-app-evidence-matrix.mjs',
  `--dir=${evidenceDir}`,
  '--required-platforms=google-meet,zoom,teams',
  `--report-file=${reportFile}`,
  '--json=true',
], { cwd: repoRoot });

const summary = JSON.parse(stdout);
assert.equal(summary.type, 'meeting_app_evidence_matrix');
assert.equal(summary.ok, false);
assert.equal(summary.evidence_file_count, 2);
assert.equal(summary.evaluated_file_count, 2);
assert.equal(summary.required_platforms.length, 3);
assert.equal(summary.production_ready_count, 1);
assert.deepEqual(summary.missing_platforms, ['microsoft_teams']);

const googleRow = summary.rows.find((row) => row.platform === 'google_meet');
assert.equal(googleRow.production_ready, true);
assert.equal(googleRow.evidence_level, 'captured_dom');
assert.equal(googleRow.source_file.endsWith('google-meet.json'), true);

const zoomRow = summary.rows.find((row) => row.platform === 'zoom');
assert.equal(zoomRow.status, 'failed');
assert.equal(zoomRow.production_ready, false);
assert.equal(zoomRow.blocking_issue_codes.includes('missing_required_coverage'), true);

const teamsRow = summary.rows.find((row) => row.platform === 'microsoft_teams');
assert.equal(teamsRow.status, 'missing');
assert.equal(teamsRow.next_actions.includes('capture_live_dom_snapshots_for_this_platform'), true);

const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.type, 'meeting_app_evidence_matrix');
assert.equal(report.ok, false);

console.log('ok meeting app evidence matrix script');
