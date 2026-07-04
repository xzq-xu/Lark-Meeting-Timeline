import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import { capturePlatformWebhookEvent } from '../packages/meeting-timeline-sdk/adapters/platform-capture.mjs';
import { buildMeetingPlatformEvidencePackage } from '../packages/meeting-timeline-sdk/adapters/platform-evidence-package.mjs';
import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';

const execFileAsync = promisify(execFile);
const baseUrl = 'https://timeline.example.com';
const startMs = 1_784_010_000_000;
const durationMs = 120_000;
const tmp = await mkdtemp(join(tmpdir(), 'meeting-platform-handoff-readiness-'));
const packageDir = join(tmp, 'packages');
const reportFile = join(tmp, 'report.json');
await mkdir(packageDir, { recursive: true });

function realSnapshot(state, observedAtMs) {
  return {
    ...buildMeetingAppFixtureSnapshot('google-meet', {
      state,
      observedAtMs,
      title: state === 'active' ? 'Handoff CLI meeting' : 'Ready to join',
      speakerId: 'real-speaker-001',
      speakerName: 'Ada Real',
      mutedId: 'real-muted-001',
      mutedName: 'Grace Real',
    }),
    id: `handoff-cli-google-${state}-${observedAtMs}`,
    source: 'chrome_extension_capture',
    fixture_state: undefined,
  };
}

function realProviderRecord(signalType, capturedAtMs) {
  const body = buildPlatformFixtureEvent('google-meet', signalType, {
    startMs,
    durationMs,
    title: 'Handoff CLI meeting',
    participantId: 'real-user-001',
    participantName: 'Ada Real',
    speakerId: 'real-speaker-001',
    speakerName: 'Ada Real',
    googleRecordId: 'real-google-record-001',
    transcriptId: 'real-transcript-001',
  });
  return capturePlatformWebhookEvent({
    platform: 'google-meet',
    method: 'POST',
    url: `${baseUrl}/api/platform-events/google-meet`,
    body,
    id: `handoff-cli-google-${signalType}`,
  }, undefined, {
    capturedAtMs,
  });
}

const evidencePackage = buildMeetingPlatformEvidencePackage('google-meet', {
  providerRecords: [
    realProviderRecord('meeting_start', startMs),
    realProviderRecord('participant_join', startMs + 30_000),
    realProviderRecord('participant_left', startMs + 90_000),
    realProviderRecord('meeting_end', startMs + durationMs),
    realProviderRecord('transcript_ready', startMs + durationMs + 120_000),
  ],
  meetingAppRecords: [
    realSnapshot('active', startMs),
    realSnapshot('prejoin', startMs + durationMs),
  ],
}, {
  baseUrl,
  env: {
    GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
    GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL: 'meet-events@example.iam.gserviceaccount.com',
  },
  includeRunbook: false,
});
await writeFile(join(packageDir, 'google_meet.json'), `${JSON.stringify(evidencePackage, null, 2)}\n`, 'utf8');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-handoff-readiness.mjs',
  '--platforms=google-meet',
  `--base-url=${baseUrl}`,
  `--package-dir=${packageDir}`,
  `--report-file=${reportFile}`,
], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
    GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL: 'meet-events@example.iam.gserviceaccount.com',
  },
});
assert.match(stdout, /meeting_platform_handoff_readiness_report/);
assert.match(stdout, /handoff_ready=1\/1/);
assert.match(stdout, /production_ready=1/);
assert.match(stdout, /runtime_replay=1\/1/);

const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.ok, true);
assert.equal(report.platform_count, 1);
assert.equal(report.handoff_ready_count, 1);
assert.equal(report.production_ready_count, 1);
assert.equal(report.runtime_host_replay_ready_count, 1);
assert.equal(report.rows[0].status, 'production_ready');
assert.equal(report.rows[0].runtime_host_replay_accepted, true);
assert.equal(report.rows[0].evidence_package_file.endsWith('google_meet.json'), true);

console.log('ok meeting platform handoff readiness script');
