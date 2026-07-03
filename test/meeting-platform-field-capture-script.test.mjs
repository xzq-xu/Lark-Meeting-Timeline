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
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-field-capture-script-'));
const packageDir = join(tmpDir, 'packages');
const reportFile = join(tmpDir, 'field-capture-report.json');
const baseUrl = 'https://timeline.example.com';
const startMs = 1_783_702_000_000;
const endMs = startMs + 120_000;

await mkdir(packageDir, { recursive: true });

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
    id: `${platform}-field-capture-script-dom`,
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

const zoomDomOnly = buildMeetingPlatformEvidencePackage('zoom', {
  meetingAppRecordSet: meetingAppRecordSet('zoom'),
}, {
  baseUrl,
  includeRunbook: false,
});

const webexReady = buildMeetingPlatformEvidencePackage('webex', {
  providerRecords: providerRecords('webex'),
  meetingAppRecordSet: meetingAppRecordSet('webex'),
}, {
  baseUrl,
  env: {
    WEBEX_WEBHOOK_SECRET: 'secret',
  },
  includeRunbook: false,
});

await writeFile(join(packageDir, 'zoom.json'), `${JSON.stringify(zoomDomOnly, null, 2)}\n`, 'utf8');
await writeFile(join(packageDir, 'webex.json'), `${JSON.stringify(webexReady, null, 2)}\n`, 'utf8');

const env = {
  ...process.env,
  WEBEX_WEBHOOK_SECRET: 'secret',
};

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-field-capture.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,zoom,webex',
  `--dir=${packageDir}`,
  `--report-file=${reportFile}`,
], {
  cwd: repoRoot,
  env,
});

assert.match(stdout, /meeting_platform_field_capture_report/);
assert.match(stdout, /production_ready=1\/3/);
assert.match(stdout, /missing_items=/);

const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.type, 'meeting_platform_field_capture_report');
assert.equal(report.ok, false);
assert.equal(report.platform_count, 3);
assert.equal(report.production_ready_count, 1);
assert.equal(report.realtime_ready_count, 2);
assert.equal(report.rows.find((row) => row.platform === 'webex').status, 'production_ready');
assert.equal(report.rows.find((row) => row.platform === 'zoom').status, 'pilot_ready_provider_pending');
assert.equal(report.rows.find((row) => row.platform === 'google_meet').missing_items.includes('capture_required_snapshot:active_speaker'), true);
assert.equal(report.rows.find((row) => row.platform === 'zoom').missing_items.includes('capture_real_provider_start_end_events'), true);
assert.equal(report.package_files_by_platform.zoom.endsWith('zoom.json'), true);

const { stdout: jsonStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-field-capture.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet',
  '--json=true',
  '--plan-only=true',
], {
  cwd: repoRoot,
});
const planOnlyReport = JSON.parse(jsonStdout);
assert.equal(planOnlyReport.ok, true);
assert.equal(planOnlyReport.plan_only, true);
assert.equal(planOnlyReport.rows[0].status, 'needs_local_and_provider_evidence');

await assert.rejects(
  execFileAsync(process.execPath, [
    'scripts/meeting-platform-field-capture.mjs',
    `--base-url=${baseUrl}`,
    '--platforms=zoom',
    `--dir=${packageDir}`,
    '--fail-on-incomplete=true',
  ], {
    cwd: repoRoot,
  }),
  (error) => error.code === 2,
);

console.log('ok meeting platform field capture script');
