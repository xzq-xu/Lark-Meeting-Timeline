import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import { buildMeetingAppSnapshotRecordSet } from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';
import { capturePlatformWebhookEvent } from '../packages/meeting-timeline-sdk/adapters/platform-capture.mjs';
import { buildMeetingPlatformEvidencePackage } from '../packages/meeting-timeline-sdk/adapters/platform-evidence-package.mjs';
import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-live-readiness-'));
const inputDir = join(tmpDir, 'packages');
const reportFile = join(tmpDir, 'live-readiness-report.json');
const pilotReportFile = join(tmpDir, 'live-readiness-pilot-report.json');
const baseUrl = 'https://timeline.example.com';
const observedAtMs = 1_784_225_000_000;
await mkdir(inputDir, { recursive: true });

function providerRecords(platform) {
  return [
    capturePlatformWebhookEvent(platform, buildPlatformFixtureEvent(platform, 'meeting_start', {
      startMs: observedAtMs,
      durationMs: 120_000,
    }), {
      capturedAtMs: observedAtMs,
    }),
    capturePlatformWebhookEvent(platform, buildPlatformFixtureEvent(platform, 'meeting_end', {
      startMs: observedAtMs,
      durationMs: 120_000,
    }), {
      capturedAtMs: observedAtMs + 120_000,
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
      capturedAtMs: observedAtMs + 120_000,
      snapshot: buildMeetingAppFixtureSnapshot(platform, {
        observedAtMs: observedAtMs + 120_000,
        state: 'prejoin',
      }),
    },
  ], {
    id: `${key}-record-set`,
    createdAtMs: observedAtMs + 121_000,
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
  createdAtMs: observedAtMs + 122_000,
});
const zoomPackage = buildMeetingPlatformEvidencePackage('zoom', {
  meetingAppRecordSet: meetingAppRecordSet('zoom'),
}, {
  baseUrl,
  createdAtMs: observedAtMs + 123_000,
  includeRunbook: false,
});

await writeFile(join(inputDir, 'google-meet.json'), `${JSON.stringify(googlePackage, null, 2)}\n`, 'utf8');
await writeFile(join(inputDir, 'zoom.json'), `${JSON.stringify(zoomPackage, null, 2)}\n`, 'utf8');

const env = {
  ...process.env,
  GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
};

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-live-readiness.mjs',
  `--input-dir=${inputDir}`,
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,zoom,teams',
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
  env,
  maxBuffer: 10 * 1024 * 1024,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_live_readiness_report');
assert.equal(report.ok, false);
assert.equal(report.target, 'production');
assert.equal(report.file_count, 2);
assert.equal(report.evaluated_file_count, 2);
assert.equal(report.platform_count, 3);
assert.equal(report.passed_count, 1);
assert.equal(report.production_ready_count, 1);
assert.equal(report.realtime_ready_count, 2);
assert.equal(report.blocked_count, 2);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').status, 'ready');
assert.equal(report.rows.find((row) => row.platform === 'google_meet').verification.correlation_status, 'matched');
assert.equal(report.rows.find((row) => row.platform === 'zoom').passed, false);
assert.equal(report.rows.find((row) => row.platform === 'zoom').blocking_codes.includes('production_evidence_ready'), true);
assert.equal(report.rows.find((row) => row.platform === 'microsoft_teams').blocking_codes.includes('pilot_realtime_axis_ready'), true);
assert.equal(report.package_files_by_platform.google_meet.endsWith('google-meet.json'), true);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 3);

const { stdout: pilotStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-live-readiness.mjs',
  `--input-dir=${inputDir}`,
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,zoom',
  '--require-production-ready=false',
  `--report-file=${pilotReportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
  env,
  maxBuffer: 10 * 1024 * 1024,
});

const pilotReport = JSON.parse(pilotStdout);
assert.equal(pilotReport.ok, true);
assert.equal(pilotReport.target, 'pilot');
assert.equal(pilotReport.passed_count, 2);
assert.equal(pilotReport.ready_count, 1);
assert.equal(pilotReport.warning_count, 1);
assert.equal(pilotReport.rows.find((row) => row.platform === 'zoom').status, 'warning');
assert.equal(pilotReport.rows.find((row) => row.platform === 'zoom').warning_codes.includes('production_evidence_ready'), true);

console.log('ok meeting platform live readiness script');
