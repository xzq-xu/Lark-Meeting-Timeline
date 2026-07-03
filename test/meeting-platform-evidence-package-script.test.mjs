import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import { buildMeetingAppSnapshotRecordSet } from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';
import { buildMeetingPlatformEvidencePackage } from '../packages/meeting-timeline-sdk/adapters/platform-evidence-package.mjs';
import { capturePlatformWebhookEvent } from '../packages/meeting-timeline-sdk/adapters/platform-capture.mjs';
import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-evidence-package-script-'));
const inputDir = join(tmpDir, 'packages');
const reportFile = join(tmpDir, 'report.json');
const realtimeReportFile = join(tmpDir, 'realtime-report.json');
const baseUrl = 'https://timeline.example.com';
const observedAtMs = 1_783_961_000_000;
await mkdir(inputDir, { recursive: true });

function providerRecords(platform) {
  return [
    capturePlatformWebhookEvent(platform, buildPlatformFixtureEvent(platform, 'meeting_start', {
      startMs: observedAtMs,
      durationMs: 90_000,
    }), {
      capturedAtMs: observedAtMs,
    }),
    capturePlatformWebhookEvent(platform, buildPlatformFixtureEvent(platform, 'meeting_end', {
      startMs: observedAtMs,
      durationMs: 90_000,
    }), {
      capturedAtMs: observedAtMs + 90_000,
    }),
  ];
}

function meetingAppRecordSet(platform) {
  const key = platform.replaceAll('-', '_');
  return buildMeetingAppSnapshotRecordSet([
    {
      platform: key,
      phase: 'active',
      capturedAtMs: observedAtMs,
      snapshot: buildMeetingAppFixtureSnapshot(platform, {
        observedAtMs,
        state: 'active',
      }),
    },
    {
      platform: key,
      phase: 'ended',
      capturedAtMs: observedAtMs + 90_000,
      snapshot: buildMeetingAppFixtureSnapshot(platform, {
        observedAtMs: observedAtMs + 90_000,
        state: 'prejoin',
      }),
    },
  ], {
    id: `${key}-record-set`,
    createdAtMs: observedAtMs + 91_000,
  });
}

const googlePackage = buildMeetingPlatformEvidencePackage('google-meet', {
  providerRecords: providerRecords('google-meet'),
  meetingAppRecordSet: meetingAppRecordSet('google-meet'),
}, {
  baseUrl,
  env: {
    GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
  },
  createdAtMs: observedAtMs + 92_000,
});
const zoomPackage = buildMeetingPlatformEvidencePackage('zoom', {
  meetingAppRecordSet: meetingAppRecordSet('zoom'),
}, {
  baseUrl,
  createdAtMs: observedAtMs + 93_000,
  includeRunbook: false,
});

await writeFile(join(inputDir, 'google-meet.json'), `${JSON.stringify(googlePackage, null, 2)}\n`, 'utf8');
await writeFile(join(inputDir, 'zoom.json'), `${JSON.stringify(zoomPackage, null, 2)}\n`, 'utf8');

const env = {
  ...process.env,
  GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
};

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-evidence-package.mjs',
  `--input-dir=${inputDir}`,
  `--base-url=${baseUrl}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
  env,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_evidence_package_report');
assert.equal(report.ok, false);
assert.equal(report.file_count, 2);
assert.equal(report.evaluated_file_count, 2);
assert.equal(report.passed_count, 1);
assert.equal(report.production_ready_count, 1);
assert.equal(report.realtime_ready_count, 2);
assert.equal(report.correlation_passed_count, 2);
assert.equal(report.stale_embedded_plan_count, 0);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').passed, true);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').correlation_status, 'matched');
assert.equal(report.rows.find((row) => row.platform === 'zoom').passed, false);
assert.equal(report.rows.find((row) => row.platform === 'zoom').correlation_status, 'single_source');
assert.equal(report.rows.find((row) => row.platform === 'zoom').ready_for_realtime_annotations, true);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 2);

const { stdout: realtimeStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-evidence-package.mjs',
  `--input-dir=${inputDir}`,
  `--base-url=${baseUrl}`,
  `--report-file=${realtimeReportFile}`,
  '--require-production-ready=false',
  '--json=true',
], {
  cwd: repoRoot,
  env,
});
const realtimeReport = JSON.parse(realtimeStdout);
assert.equal(realtimeReport.ok, true);
assert.equal(realtimeReport.requirement, 'ready_for_realtime_annotations');
assert.equal(realtimeReport.passed_count, 2);

console.log('ok meeting platform evidence package script');
