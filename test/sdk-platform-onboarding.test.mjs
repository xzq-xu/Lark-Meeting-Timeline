import assert from 'node:assert/strict';

import {
  buildAllMeetingPlatformOnboardingReports,
  buildMeetingPlatformOnboardingReport,
  buildMeetingPlatformOnboardingSummary,
} from '../packages/meeting-timeline-sdk/adapters/platform-onboarding.mjs';
import { platformCapabilityContract } from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';

const baseUrl = 'https://timeline.example.com';
const startMs = 1_782_442_800_000;
const startIso = new Date(startMs).toISOString();

const googleEnv = {
  GOOGLE_PUBSUB_OIDC_AUDIENCE: 'https://timeline.example.com/api/platform-events/google-meet',
  GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL: 'pubsub@demo.iam.gserviceaccount.com',
};

const googleStart = {
  id: 'google-onboarding-start-1',
  type: 'google.workspace.meet.conference.v2.started',
  time: startIso,
  data: {
    conferenceRecord: { name: 'conferenceRecords/google-onboarding-001' },
    meetingUri: 'https://meet.google.com/abc-defg-hij',
    title: 'Google onboarding review',
  },
};

const googleTranscript = {
  id: 'google-onboarding-transcript-1',
  type: 'google.workspace.meet.transcript.v2.fileGenerated',
  time: new Date(startMs + 35 * 60 * 1000).toISOString(),
  data: {
    transcript: {
      name: 'conferenceRecords/google-onboarding-001/transcripts/transcript-1',
      docsDestination: { document: 'https://docs.google.com/document/d/transcript-1' },
    },
  },
};

assert.equal(
  platformCapabilityContract('google-meet', { baseUrl }).sdk_modules.onboarding,
  '@ai-annotation/meeting-timeline-sdk/adapters/platform-onboarding',
);

const googleReport = buildMeetingPlatformOnboardingReport('google-meet', {
  baseUrl,
  env: googleEnv,
  features: ['axis', 'participants', 'transcript', 'recording', 'security'],
  samples: {
    google_meet: [
      { label: 'provider start', body: googleStart },
      { label: 'transcript ready', body: googleTranscript },
    ],
  },
});

assert.equal(googleReport.platform, 'google_meet');
assert.equal(googleReport.status, 'ready');
assert.equal(googleReport.ready, true);
assert.equal(googleReport.runtime_contract.annotation_time_field, 'captured_at_ms');
assert.equal(googleReport.runtime_contract.primary_axis_source, 'local_observer');
assert.equal(googleReport.permission_plan.required_scopes.includes('https://www.googleapis.com/auth/drive.meet.readonly'), true);
assert.equal(googleReport.acceptance.accepted, true);
assert.equal(googleReport.acceptance.coverage.meeting_start, true);
assert.equal(googleReport.acceptance.coverage.artifact_ready, true);
assert.equal(googleReport.artifact_import_plans.length, 1);
assert.equal(googleReport.artifact_import_plans[0].transcript_import.normalizer, 'normalizeGoogleMeetTranscriptEntries');
assert.equal(googleReport.next_actions.includes('artifact:google_meet:fetch_transcript'), true);

const zoomBlocked = buildMeetingPlatformOnboardingReport('zoom', {
  baseUrl,
  env: {},
  samples: {
    zoom: [{
      body: {
        event: 'meeting.started',
        event_ts: startMs,
        payload: {
          object: {
            uuid: 'zoom-onboarding-001',
            id: 123456789,
            join_url: 'https://zoom.us/j/123456789',
          },
        },
      },
    }],
  },
});
assert.equal(zoomBlocked.status, 'blocked_by_setup');
assert.equal(zoomBlocked.ready, false);
assert.equal(zoomBlocked.next_actions.includes('configure_env:ZOOM_WEBHOOK_SECRET_TOKEN'), true);
assert.equal(zoomBlocked.next_actions.includes('fix_setup:required_security_env'), true);

const allReports = buildAllMeetingPlatformOnboardingReports({
  baseUrl,
  env: googleEnv,
  samples: {
    google_meet: [{ body: googleStart }],
  },
});
assert.equal(allReports.length, 6);
assert.deepEqual(allReports.map((item) => item.platform), ['local_detector', 'lark', 'google_meet', 'microsoft_teams', 'zoom', 'webex']);
assert.equal(allReports.find((item) => item.platform === 'google_meet').ready, true);
assert.equal(allReports.find((item) => item.platform === 'zoom').status, 'blocked_by_setup');

const summary = buildMeetingPlatformOnboardingSummary({
  baseUrl,
  env: googleEnv,
  samples: {
    google_meet: [{ body: googleStart }],
  },
});
assert.equal(summary.reports.length, 6);
assert.equal(summary.ok, false);
assert.equal(summary.ready_count >= 1, true);
assert.equal(summary.blocked_count >= 1, true);
assert.equal(summary.acceptance_summary.reports.length, 6);

console.log('ok meeting platform onboarding reports');
