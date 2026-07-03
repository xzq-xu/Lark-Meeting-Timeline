import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import { buildMeetingAppSnapshotRecordSet } from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';
import { capturePlatformWebhookEvent } from '../packages/meeting-timeline-sdk/adapters/platform-capture.mjs';
import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-field-evidence-script-'));
const inputDir = join(tmpDir, 'field-evidence');
const packageDir = join(tmpDir, 'packages');
const bundleDir = join(tmpDir, 'bundles');
const reportFile = join(tmpDir, 'field-evidence-report.json');
const baseUrl = 'https://timeline.example.com';
const startMs = 1_783_702_000_000;
const endMs = startMs + 120_000;

await mkdir(inputDir, { recursive: true });

function meetingAppRecordSet(platform) {
  return buildMeetingAppSnapshotRecordSet([
    {
      platform,
      phase: 'active',
      capturedAtMs: startMs,
      snapshot: buildMeetingAppFixtureSnapshot(platform, {
        observedAtMs: startMs,
        state: 'active',
      }),
    },
    {
      platform,
      phase: 'ended',
      capturedAtMs: endMs,
      snapshot: buildMeetingAppFixtureSnapshot(platform, {
        observedAtMs: endMs,
        state: 'prejoin',
      }),
    },
  ], {
    id: `${platform}-field-evidence-script-dom`,
    createdAtMs: endMs + 1_000,
  });
}

function providerRecords(platform) {
  return [
    capturePlatformWebhookEvent(platform, buildPlatformFixtureEvent(platform, 'meeting_start', {
      startMs,
      durationMs: 120_000,
    }), {
      capturedAtMs: startMs,
      label: `${platform} start`,
    }),
    capturePlatformWebhookEvent(platform, buildPlatformFixtureEvent(platform, 'meeting_end', {
      startMs,
      durationMs: 120_000,
    }), {
      capturedAtMs: endMs,
      label: `${platform} end`,
    }),
  ];
}

await writeFile(join(inputDir, 'zoom-dom.json'), `${JSON.stringify({
  platform: 'zoom',
  meetingAppRecordSet: meetingAppRecordSet('zoom'),
}, null, 2)}\n`, 'utf8');

await writeFile(join(inputDir, 'webex-field.json'), `${JSON.stringify({
  platform: 'webex',
  providerRecords: providerRecords('webex'),
  meetingAppRecordSet: meetingAppRecordSet('webex'),
}, null, 2)}\n`, 'utf8');

const env = {
  ...process.env,
  WEBEX_WEBHOOK_SECRET: 'secret',
};

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-field-evidence.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=zoom,webex',
  `--dir=${inputDir}`,
  `--package-dir=${packageDir}`,
  `--bundle-dir=${bundleDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
  env,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_field_evidence_report');
assert.equal(report.ok, false);
assert.equal(report.platform_count, 2);
assert.equal(report.passed_count, 1);
assert.equal(report.production_ready_count, 1);
assert.equal(report.realtime_ready_count, 2);
assert.equal(report.rows.find((row) => row.platform === 'webex').verified, true);
assert.equal(report.rows.find((row) => row.platform === 'webex').source_files.length, 1);
assert.equal(report.rows.find((row) => row.platform === 'zoom').status, 'pilot_ready_provider_pending');
assert.equal(report.rows.find((row) => row.platform === 'zoom').missing_items.includes('capture_real_provider_start_end_events'), true);
assert.equal(report.written_files.some((file) => file.endsWith('/webex.json')), true);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 2);
const webexPackage = JSON.parse(await readFile(join(packageDir, 'webex.json'), 'utf8'));
assert.equal(webexPackage.schema, 'meeting_platform_evidence_package');
assert.equal(webexPackage.platform, 'webex');
const webexBundle = JSON.parse(await readFile(join(bundleDir, 'webex.json'), 'utf8'));
assert.equal(webexBundle.schema, 'meeting_platform_field_evidence_bundle');
assert.equal(webexBundle.verification.passed, true);

const { stdout: realtimeStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-field-evidence.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=zoom',
  `--dir=${inputDir}`,
  `--package-dir=${join(tmpDir, 'realtime-packages')}`,
  `--bundle-dir=${join(tmpDir, 'realtime-bundles')}`,
  '--require-production-ready=false',
  '--json=true',
], {
  cwd: repoRoot,
});
const realtimeReport = JSON.parse(realtimeStdout);
assert.equal(realtimeReport.ok, true);
assert.equal(realtimeReport.requirement, 'ready_for_realtime_annotations');
assert.equal(realtimeReport.passed_count, 1);

await assert.rejects(
  execFileAsync(process.execPath, [
    'scripts/meeting-platform-field-evidence.mjs',
    `--base-url=${baseUrl}`,
    '--platforms=zoom',
    `--dir=${inputDir}`,
    '--fail-on-incomplete=true',
  ], {
    cwd: repoRoot,
  }),
  (error) => error.code === 2,
);

console.log('ok meeting platform field evidence script');
