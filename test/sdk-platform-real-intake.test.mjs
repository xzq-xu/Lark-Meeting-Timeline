import assert from 'node:assert/strict';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';
import { capturePlatformWebhookEvent } from '../packages/meeting-timeline-sdk/adapters/platform-capture.mjs';
import {
  assertMeetingPlatformRealEvidenceIntake,
  assertMeetingPlatformRealEvidenceIntakeMatrix,
  buildMeetingPlatformRealEvidenceIntakeMatrix,
  buildMeetingPlatformRealEvidenceIntakePlan,
  buildMeetingPlatformRealEvidenceIntakeReport,
} from '../packages/meeting-timeline-sdk/adapters/platform-real-intake.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';
const startMs = 1_784_010_000_000;
const durationMs = 120_000;
const productionEnv = {
  GOOGLE_PUBSUB_OIDC_AUDIENCE: 'https://timeline.example.com/api/platform-events/google-meet',
  GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL: 'meet-events@example.iam.gserviceaccount.com',
  MICROSOFT_GRAPH_CLIENT_STATE: 'real-teams-client-state',
  ZOOM_WEBHOOK_SECRET_TOKEN: 'real-zoom-secret-token',
  WEBEX_WEBHOOK_SECRET: 'real-webex-secret',
};

function realSnapshot(platform, state, observedAtMs) {
  return {
    ...buildMeetingAppFixtureSnapshot(platform, {
      state,
      observedAtMs,
      title: state === 'active' ? 'Real intake meeting' : 'Ready to join',
      speakerId: 'real-speaker-001',
      speakerName: 'Ada Real',
      mutedId: 'real-muted-001',
      mutedName: 'Grace Real',
    }),
    id: `real-${platform}-${state}-${observedAtMs}`,
    source: 'chrome_extension_capture',
    fixture_state: undefined,
  };
}

function realProviderRecord(platform, signalType, capturedAtMs) {
  const body = buildPlatformFixtureEvent(platform, signalType, {
    startMs,
    durationMs,
    title: 'Real intake meeting',
    participantId: 'real-user-001',
    participantName: 'Ada Real',
    speakerId: 'real-speaker-001',
    speakerName: 'Ada Real',
    googleRecordId: 'real-google-record-001',
    teamsMeetingId: 'real-teams-meeting-001',
    zoomUuid: 'real-zoom-uuid-001',
    webexMeetingId: 'real-webex-meeting-001',
    larkMeetingId: 'real-lark-meeting-001',
    transcriptId: 'real-transcript-001',
    recordingId: 'real-recording-001',
  });
  return capturePlatformWebhookEvent({
    platform,
    method: 'POST',
    url: `${baseUrl}/api/platform-events/${platform}`,
    body,
    id: `real-${platform}-${signalType}`,
  }, undefined, {
    capturedAtMs,
  });
}

function realInput(platform) {
  return {
    providerRecords: [
      realProviderRecord(platform, 'meeting_start', startMs),
      realProviderRecord(platform, 'participant_join', startMs + 30_000),
      realProviderRecord(platform, 'participant_left', startMs + 90_000),
      realProviderRecord(platform, 'meeting_end', startMs + durationMs),
      realProviderRecord(platform, 'transcript_ready', startMs + durationMs + 120_000),
    ],
    meetingAppRecords: [
      realSnapshot(platform, 'active', startMs),
      realSnapshot(platform, 'prejoin', startMs + durationMs),
    ],
  };
}

const plan = buildMeetingPlatformRealEvidenceIntakePlan('google-meet', { baseUrl });
assert.equal(plan.schema, 'meeting_platform_real_evidence_intake_plan');
assert.equal(plan.platform, 'google_meet');
assert.equal(plan.provider_endpoint, `${baseUrl}/api/platform-events/google-meet`);
assert.equal(plan.forbidden_inputs.includes('meeting_app_fixture_source'), true);

const googleReport = buildMeetingPlatformRealEvidenceIntakeReport('google-meet', realInput('google-meet'), {
  baseUrl,
  env: productionEnv,
});
assert.equal(googleReport.schema, 'meeting_platform_real_evidence_intake');
assert.equal(googleReport.platform, 'google_meet');
assert.equal(googleReport.accepted, true);
assert.equal(googleReport.provider_record_count, 5);
assert.equal(googleReport.meeting_app_record_count, 2);
assert.equal(googleReport.fixture_evidence_count, 0);
assert.equal(googleReport.verification.passed, true);
assert.equal(googleReport.readiness.status, 'ready');
assert.equal(googleReport.next_actions.includes('handoff_evidence_package_to_host_project'), true);

const fixtureRejected = buildMeetingPlatformRealEvidenceIntakeReport('google-meet', {
  providerRecords: [realProviderRecord('google-meet', 'meeting_start', startMs)],
  meetingAppRecords: [buildMeetingAppFixtureSnapshot('google-meet', { state: 'active', observedAtMs: startMs })],
}, {
  baseUrl,
  env: productionEnv,
  requireProductionReady: false,
});
assert.equal(fixtureRejected.accepted, false);
assert.equal(fixtureRejected.blocking_checks.some((check) => check.code === 'no_fixture_evidence'), true);

assert.equal(assertMeetingPlatformRealEvidenceIntake('zoom', realInput('zoom'), { baseUrl, env: productionEnv }).accepted, true);

const matrix = buildMeetingPlatformRealEvidenceIntakeMatrix({
  platforms: ['google-meet', 'zoom'],
  google_meet: realInput('google-meet'),
  zoom: realInput('zoom'),
}, {
  baseUrl,
  env: productionEnv,
});
assert.equal(matrix.schema, 'meeting_platform_real_evidence_intake_matrix');
assert.equal(matrix.platform_count, 2);
assert.equal(matrix.accepted_count, 2);
assert.equal(matrix.rejected_count, 0);
assert.equal(assertMeetingPlatformRealEvidenceIntakeMatrix({
  platforms: ['webex'],
  webex: realInput('webex'),
}, {
  baseUrl,
  env: productionEnv,
}).accepted_count, 1);

const client = {
  async startMeeting(input) { return { ok: true, input }; },
  async endMeeting(input) { return { ok: true, input }; },
  async insertMark(input) { return { ok: true, input }; },
  async insertMarks(input) { return { ok: true, input }; },
};
const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl,
  env: productionEnv,
  verify: false,
});

assert.equal(kit.platformRealEvidenceIntakePlan('google-meet').platform, 'google_meet');
assert.equal(kit.platformRealEvidenceIntake('google-meet', realInput('google-meet')).accepted, true);
assert.equal(kit.platformRealEvidenceIntakeMatrix({
  platforms: ['lark'],
  lark: realInput('lark'),
}).accepted_count, 1);
assert.equal(kit.assertPlatformRealEvidenceIntake('teams', realInput('teams')).accepted, true);
assert.equal(kit.assertPlatformRealEvidenceIntakeMatrix({
  platforms: ['zoom'],
  zoom: realInput('zoom'),
}).accepted_count, 1);

console.log('ok meeting platform real evidence intake');
