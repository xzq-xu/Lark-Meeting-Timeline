import assert from 'node:assert/strict';

import {
  PLATFORM_ACCEPTANCE_COVERAGE_KEYS,
  buildAllPlatformAcceptanceReports,
  buildMeetingPlatformAcceptanceSummary,
  buildPlatformAcceptanceReport,
} from '../packages/meeting-timeline-sdk/adapters/platform-acceptance.mjs';
import { platformCapabilityContract } from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';

const baseUrl = 'https://timeline.example.com';
const startMs = 1_782_442_800_000;
const startIso = new Date(startMs).toISOString();
const endIso = new Date(startMs + 30 * 60 * 1000).toISOString();
const googleEnv = {
  GOOGLE_PUBSUB_OIDC_AUDIENCE: 'https://timeline.example.com/api/platform-events/google-meet',
  GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL: 'pubsub@demo.iam.gserviceaccount.com',
};

const googleStart = {
  id: 'google-accept-start-1',
  type: 'google.workspace.meet.conference.v2.started',
  time: startIso,
  data: {
    conferenceRecord: { name: 'conferenceRecords/google-accept-001' },
    meetingUri: 'https://meet.google.com/abc-defg-hij',
    title: 'Google acceptance review',
  },
};

const googleEnd = {
  id: 'google-accept-end-1',
  type: 'google.workspace.meet.conference.v2.ended',
  time: endIso,
  data: {
    conferenceRecord: { name: 'conferenceRecords/google-accept-001' },
    meetingUri: 'https://meet.google.com/abc-defg-hij',
    title: 'Google acceptance review',
  },
};

assert.equal(PLATFORM_ACCEPTANCE_COVERAGE_KEYS.includes('meeting_start'), true);
assert.equal(
  platformCapabilityContract('google-meet', { baseUrl }).sdk_modules.acceptance,
  '@ai-annotation/meeting-timeline-sdk/adapters/platform-acceptance',
);

const googleAccepted = buildPlatformAcceptanceReport('google-meet', {
  baseUrl,
  env: googleEnv,
  requireEndEvent: true,
  samples: {
    google_meet: [
      { label: 'provider start', body: googleStart },
      { label: 'provider end', body: googleEnd },
    ],
  },
});
assert.equal(googleAccepted.platform, 'google_meet');
assert.equal(googleAccepted.status, 'accepted');
assert.equal(googleAccepted.accepted, true);
assert.equal(googleAccepted.readiness.ready, true);
assert.deepEqual(googleAccepted.required_coverage, ['meeting_start', 'meeting_end']);
assert.deepEqual(googleAccepted.missing_required_coverage, []);
assert.equal(googleAccepted.coverage.meeting_start, true);
assert.equal(googleAccepted.coverage.meeting_end, true);
assert.equal(googleAccepted.sample_count, 2);
assert.equal(googleAccepted.actionable_sample_count, 2);
assert.equal(googleAccepted.signal_types.includes('meeting_started'), true);
assert.equal(googleAccepted.signal_types.includes('meeting_ended'), true);
assert.equal(googleAccepted.samples[0].diagnostic.meetings[0].meeting_url, 'https://meet.google.com/abc-defg-hij');

const googleMissingEnd = buildPlatformAcceptanceReport('google-meet', {
  baseUrl,
  env: googleEnv,
  requireEndEvent: true,
  samples: {
    meet: [{ label: 'provider start only', body: googleStart }],
  },
});
assert.equal(googleMissingEnd.status, 'missing_required_coverage');
assert.equal(googleMissingEnd.accepted, false);
assert.deepEqual(googleMissingEnd.missing_required_coverage, ['meeting_end']);
assert.equal(googleMissingEnd.issues.some((item) => item.code === 'missing_required_coverage'), true);

const zoomBlocked = buildPlatformAcceptanceReport('zoom', {
  baseUrl: 'http://timeline.example.com',
  env: {},
  samples: {
    zoom: [{
      body: {
        event: 'meeting.started',
        event_ts: startMs,
        payload: {
          object: {
            uuid: 'zoom-accept-001',
            id: 123456789,
            join_url: 'https://zoom.us/j/123456789',
          },
        },
      },
    }],
  },
});
assert.equal(zoomBlocked.status, 'blocked');
assert.equal(zoomBlocked.accepted, false);
assert.equal(zoomBlocked.readiness.ready, false);
assert.equal(zoomBlocked.issues.some((item) => item.code === 'platform_setup_not_ready'), true);

const allReports = buildAllPlatformAcceptanceReports({
  baseUrl,
  env: googleEnv,
  samples: {
    google_meet: [{ body: googleStart }],
    local_detector: [{
      body: {
        type: 'meeting_started',
        detected_platform: 'google_meet',
        meeting_url: 'https://meet.google.com/abc-defg-hij',
        start_time_ms: startMs,
      },
    }],
  },
});
assert.equal(allReports.length, 6);
assert.equal(allReports.find((item) => item.platform === 'local_detector').status, 'accepted');
assert.equal(allReports.find((item) => item.platform === 'microsoft_teams').status, 'blocked');

const summary = buildMeetingPlatformAcceptanceSummary({
  baseUrl,
  env: googleEnv,
  samples: {
    google_meet: [{ body: googleStart }],
    local_detector: [{
      body: {
        type: 'meeting_started',
        detected_platform: 'google_meet',
        meeting_id: 'local-accept-001',
        start_time_ms: startMs,
      },
    }],
  },
});
assert.equal(summary.reports.length, 6);
assert.equal(summary.ok, false);
assert.equal(summary.accepted_count >= 2, true);
assert.equal(summary.blocked_count >= 1, true);

console.log('ok meeting platform acceptance reports');
