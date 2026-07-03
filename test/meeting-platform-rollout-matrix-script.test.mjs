import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import { buildMeetingAppLiveEvidencePackage } from '../packages/meeting-timeline-sdk/adapters/meeting-app-profile.mjs';
import { buildMeetingAppSnapshotRecordSet } from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';
import { capturePlatformWebhookEvent } from '../packages/meeting-timeline-sdk/adapters/platform-capture.mjs';
import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-rollout-matrix-'));
const providerDir = join(tmpDir, 'provider');
const domDir = join(tmpDir, 'dom');
const reportFile = join(tmpDir, 'rollout-report.json');
const baseUrl = 'https://timeline.example.com';
const observedAtMs = 1_783_356_000_000;
await mkdir(providerDir, { recursive: true });
await mkdir(domDir, { recursive: true });

function providerRecords(platform) {
  return [
    capturePlatformWebhookEvent(platform, buildPlatformFixtureEvent(platform, 'meeting_start'), {
      capturedAtMs: observedAtMs,
    }),
    capturePlatformWebhookEvent(platform, buildPlatformFixtureEvent(platform, 'meeting_end'), {
      capturedAtMs: observedAtMs + 60_000,
    }),
  ];
}

function domEvidence(platform) {
  const platformKey = platform.replaceAll('-', '_');
  const recordSet = buildMeetingAppSnapshotRecordSet([
    {
      platform: platformKey,
      phase: 'active',
      capturedAtMs: observedAtMs,
      snapshot: buildMeetingAppFixtureSnapshot(platform, {
        observedAtMs,
        state: 'active',
      }),
    },
    {
      platform: platformKey,
      phase: 'ended',
      capturedAtMs: observedAtMs + 60_000,
      snapshot: buildMeetingAppFixtureSnapshot(platform, {
        observedAtMs: observedAtMs + 60_000,
        state: 'prejoin',
      }),
    },
  ], {
    id: `${platformKey}-record-set`,
    createdAtMs: observedAtMs + 61_000,
  });
  return buildMeetingAppLiveEvidencePackage({
    platforms: [platform],
    recordSet,
    recordSetId: `${platformKey}-record-set`,
    packageId: `${platformKey}-evidence`,
  });
}

await writeFile(join(providerDir, 'google-meet.json'), `${JSON.stringify({
  platform: 'google-meet',
  records: providerRecords('google-meet'),
}, null, 2)}\n`, 'utf8');
await writeFile(join(domDir, 'google-meet.json'), `${JSON.stringify(domEvidence('google-meet'), null, 2)}\n`, 'utf8');
await writeFile(join(domDir, 'zoom.json'), `${JSON.stringify(domEvidence('zoom'), null, 2)}\n`, 'utf8');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-rollout-matrix.mjs',
  `--provider-dir=${providerDir}`,
  `--dom-dir=${domDir}`,
  '--required-platforms=google-meet,zoom,teams',
  `--base-url=${baseUrl}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
  env: {
    ...process.env,
    GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
  },
});

const summary = JSON.parse(stdout);
assert.equal(summary.type, 'meeting_platform_rollout_matrix');
assert.equal(summary.ok, false);
assert.equal(summary.provider_file_count, 1);
assert.equal(summary.dom_file_count, 2);
assert.equal(summary.provider_evaluated_file_count, 1);
assert.equal(summary.dom_evaluated_file_count, 2);
assert.equal(summary.production_ready_count, 1);
assert.equal(summary.realtime_ready_count, 2);
assert.deepEqual(summary.pending_platforms, ['zoom', 'microsoft_teams']);

const googleRow = summary.rows.find((row) => row.platform === 'google_meet');
assert.equal(googleRow.status, 'production_ready');
assert.equal(googleRow.production_ready, true);
assert.equal(googleRow.provider.evidence_level, 'captured_events');
assert.equal(googleRow.local_dom.evidence_level, 'captured_dom');
assert.equal(googleRow.provider.source_files.some((file) => file.endsWith('google-meet.json')), true);

const zoomRow = summary.rows.find((row) => row.platform === 'zoom');
assert.equal(zoomRow.status, 'realtime_ready_provider_pending');
assert.equal(zoomRow.production_ready, false);
assert.equal(zoomRow.ready_for_realtime_annotations, true);
assert.equal(zoomRow.local_dom.production_ready, true);
assert.equal(zoomRow.provider.production_ready, false);

const teamsRow = summary.rows.find((row) => row.platform === 'microsoft_teams');
assert.equal(teamsRow.production_ready, false);
assert.equal(teamsRow.ready_for_realtime_annotations, false);
assert.equal(teamsRow.next_actions.includes('capture_live_dom_snapshots_for_local_observer'), true);

const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.type, 'meeting_platform_rollout_matrix');
assert.equal(report.rows.length, 3);

console.log('ok meeting platform rollout matrix script');
