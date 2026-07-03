import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

import {
  buildMeetingAppSnapshotRecordSet,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';
import {
  capturePlatformWebhookEvent,
} from '../packages/meeting-timeline-sdk/adapters/platform-capture.mjs';
import {
  buildPlatformFixtureEvent,
} from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';
import {
  normalizeMeetingPlatform,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-real-intake-script-'));
const providerDir = join(tmpDir, 'provider');
const domDir = join(tmpDir, 'dom');
const packageDir = join(tmpDir, 'packages');
const reportFile = join(tmpDir, 'report.json');
const baseUrl = 'https://timeline.example.com';
const startMs = 1_784_250_000_000;
const durationMs = 180_000;

await mkdir(providerDir, { recursive: true });
await mkdir(domDir, { recursive: true });
await mkdir(packageDir, { recursive: true });

function platformUrl(platform) {
  const key = normalizeMeetingPlatform(platform);
  return {
    google_meet: 'https://meet.google.com/real-abcd-efg',
    zoom: 'https://zoom.us/j/1234567890',
  }[key];
}

function realSnapshot(platform, phase, capturedAtMs) {
  const key = normalizeMeetingPlatform(platform);
  const active = phase === 'active';
  const title = active
    ? 'Real intake meeting'
    : key === 'zoom' ? 'Join Meeting - Zoom' : 'Ready to join';
  const participants = active ? [{
    id: `real-${key}-speaker-001`,
    name: 'Ada Real',
    display_name: 'Ada Real',
    ariaLabel: 'Ada Real is speaking',
    label: 'Ada Real is speaking',
    audioLevel: 0.83,
  }] : [];
  const controls = active
    ? key === 'zoom'
      ? [{ label: 'Leave Meeting', ariaLabel: 'Leave Meeting' }, { label: 'Participants', ariaLabel: 'Participants' }]
      : [{ label: 'Leave call', ariaLabel: 'Leave call' }, { label: 'Participants', ariaLabel: 'Participants' }]
    : key === 'zoom'
      ? [{ label: 'Join Meeting', ariaLabel: 'Join Meeting' }]
      : [{ label: 'Join now', ariaLabel: 'Join now' }];
  return {
    id: `real-dom-${key}-${phase}-${capturedAtMs}`,
    source: 'chrome_extension_capture',
    observedAtMs: capturedAtMs,
    platform: key,
    provider: key,
    url: platformUrl(key),
    meeting_url: platformUrl(key),
    meeting_id: `real-${key}-meeting-001`,
    title,
    page: {
      url: platformUrl(key),
      title,
      documentVisible: true,
      buttons: controls,
      controls,
      tiles: participants,
      participants,
    },
    dom: {
      url: platformUrl(key),
      title,
      controls,
      participants,
    },
  };
}

function meetingAppRecordSet(platform) {
  const key = normalizeMeetingPlatform(platform);
  return buildMeetingAppSnapshotRecordSet([
    {
      platform: key,
      phase: 'active',
      capturedAtMs: startMs,
      source: 'chrome_extension_capture',
      snapshot: realSnapshot(key, 'active', startMs),
    },
    {
      platform: key,
      phase: 'ended',
      capturedAtMs: startMs + durationMs,
      source: 'chrome_extension_capture',
      snapshot: realSnapshot(key, 'ended', startMs + durationMs),
    },
  ], {
    id: `real-${key}-record-set`,
    source: 'chrome_extension_capture',
    createdAtMs: startMs + durationMs + 1_000,
  });
}

function providerRecord(platform, signalType, capturedAtMs) {
  const key = normalizeMeetingPlatform(platform);
  const body = buildPlatformFixtureEvent(key, signalType, {
    startMs,
    durationMs,
    title: 'Real intake meeting',
    participantId: `real-${key}-user-001`,
    participantName: 'Ada Real',
    googleRecordId: 'real-google-record-001',
    zoomUuid: 'real-zoom-uuid-001',
    zoomNumericId: 1234567890,
    googleUrl: platformUrl('google-meet'),
    zoomUrl: platformUrl('zoom'),
  });
  return capturePlatformWebhookEvent({
    platform: key,
    method: 'POST',
    url: `${baseUrl}/api/platform-events/${key}`,
    body,
    id: `real-${key}-${signalType}`,
  }, undefined, {
    capturedAtMs,
  });
}

function providerRecords(platform) {
  return [
    providerRecord(platform, 'meeting_start', startMs),
    providerRecord(platform, 'participant_join', startMs + 30_000),
    providerRecord(platform, 'participant_left', startMs + durationMs - 30_000),
    providerRecord(platform, 'meeting_end', startMs + durationMs),
  ];
}

await writeFile(join(providerDir, 'records.json'), `${JSON.stringify({
  records: [
    ...providerRecords('google-meet'),
    ...providerRecords('zoom'),
  ],
}, null, 2)}\n`, 'utf8');

await writeFile(join(domDir, 'records.json'), `${JSON.stringify({
  records: [
    ...meetingAppRecordSet('google-meet').records,
    ...meetingAppRecordSet('zoom').records,
  ],
}, null, 2)}\n`, 'utf8');

const env = {
  ...process.env,
  GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
  GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL: 'meet-events@example.iam.gserviceaccount.com',
  ZOOM_WEBHOOK_SECRET_TOKEN: 'real-zoom-secret-token',
};

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-real-intake.mjs',
  `--provider-dir=${providerDir}`,
  `--dom-dir=${domDir}`,
  `--package-dir=${packageDir}`,
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,zoom',
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
  env,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_real_intake_report');
assert.equal(report.matrix, undefined);
assert.equal(report.ok, true);
assert.equal(report.platform_count, 2);
assert.equal(report.accepted_count, 2);
assert.equal(report.rejected_count, 0);
assert.equal(report.production_ready_count, 2);
assert.equal(report.realtime_ready_count, 2);
assert.equal(report.provider_evaluated_file_count, 1);
assert.equal(report.dom_evaluated_file_count, 1);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').accepted, true);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').fixture_evidence_count, 0);
assert.equal(report.rows.find((row) => row.platform === 'zoom').provider_record_count, 4);
assert.equal(report.rows.find((row) => row.platform === 'zoom').meeting_app_record_count, 2);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 2);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-real-intake.mjs',
  `--provider-dir=${providerDir}`,
  `--dom-dir=${domDir}`,
  `--package-dir=${packageDir}`,
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,zoom',
], {
  cwd: repoRoot,
  env,
});
assert.match(textStdout, /meeting_platform_real_intake_report/);
assert.match(textStdout, /accepted=2\/2/);

console.log('ok meeting platform real intake script');
