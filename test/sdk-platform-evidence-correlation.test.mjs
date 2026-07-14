import assert from 'node:assert/strict';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import { buildMeetingAppSnapshotRecordSet } from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';
import { capturePlatformWebhookEvent } from '../packages/meeting-timeline-sdk/adapters/platform-capture.mjs';
import {
  MEETING_PLATFORM_EVIDENCE_CORRELATION_SCHEMA,
  buildMeetingPlatformEvidenceCorrelation,
} from '../packages/meeting-timeline-sdk/adapters/platform-evidence-correlation.mjs';
import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';
const observedAtMs = 1_783_961_000_000;

function providerRecords(platform, startMs = observedAtMs, durationMs = 90_000) {
  return [
    capturePlatformWebhookEvent(platform, buildPlatformFixtureEvent(platform, 'meeting_start', {
      startMs,
      durationMs,
    }), {
      capturedAtMs: startMs,
    }),
    capturePlatformWebhookEvent(platform, buildPlatformFixtureEvent(platform, 'meeting_end', {
      startMs,
      durationMs,
    }), {
      capturedAtMs: startMs + durationMs,
    }),
  ];
}

function meetingAppRecordSet(platform, startMs = observedAtMs, durationMs = 90_000) {
  const key = platform.replaceAll('-', '_');
  return buildMeetingAppSnapshotRecordSet([
    {
      platform: key,
      phase: 'active',
      capturedAtMs: startMs,
      snapshot: buildMeetingAppFixtureSnapshot(platform, {
        observedAtMs: startMs,
        state: 'active',
      }),
    },
    {
      platform: key,
      phase: 'ended',
      capturedAtMs: startMs + durationMs,
      snapshot: buildMeetingAppFixtureSnapshot(platform, {
        observedAtMs: startMs + durationMs,
        state: 'prejoin',
      }),
    },
  ], {
    id: `${key}-correlation-record-set`,
    createdAtMs: startMs + durationMs + 1_000,
  });
}

const google = buildMeetingPlatformEvidenceCorrelation('google-meet', {
  providerRecords: providerRecords('google-meet'),
  meetingAppRecordSet: meetingAppRecordSet('google-meet'),
});
assert.equal(google.schema, MEETING_PLATFORM_EVIDENCE_CORRELATION_SCHEMA);
assert.equal(google.platform, 'google_meet');
assert.equal(google.passed, true);
assert.equal(google.status, 'matched');
assert.equal(google.confidence, 'high');
assert.equal(google.identity_match.shared_tokens.some((token) => token.startsWith('url:')), true);
assert.equal(google.time_alignment.start_delta_ms, 0);
assert.equal(google.coverage.provider_has_start, true);
assert.equal(google.coverage.local_has_active, true);

const zoomSingleSource = buildMeetingPlatformEvidenceCorrelation('zoom', {
  meetingAppRecordSet: meetingAppRecordSet('zoom'),
});
assert.equal(zoomSingleSource.passed, true);
assert.equal(zoomSingleSource.status, 'single_source');
assert.equal(zoomSingleSource.confidence, 'none');
assert.equal(zoomSingleSource.issues.some((item) => item.code === 'single_source_correlation'), true);

const shiftedDom = buildMeetingPlatformEvidenceCorrelation('webex', {
  providerRecords: providerRecords('webex'),
  meetingAppRecordSet: meetingAppRecordSet('webex', observedAtMs + 3 * 60 * 60 * 1000),
}, {
  maxClockSkewMs: 1_000,
});
assert.equal(shiftedDom.passed, false);
assert.equal(shiftedDom.status, 'failed');
assert.equal(shiftedDom.issues.some((item) => item.code === 'time_window_mismatch'), true);

const noSharedIdButTimed = buildMeetingPlatformEvidenceCorrelation('zoom', {
  providerRecords: providerRecords('zoom'),
  meetingAppRecordSet: buildMeetingAppSnapshotRecordSet([
    {
      platform: 'zoom',
      phase: 'active',
      capturedAtMs: observedAtMs,
      meeting_id: 'local-only-meeting',
      url: 'https://zoom.us/j/local-only',
      snapshot: buildMeetingAppFixtureSnapshot('zoom', {
        observedAtMs,
        state: 'active',
        url: 'https://zoom.us/j/local-only',
      }),
    },
    {
      platform: 'zoom',
      phase: 'ended',
      capturedAtMs: observedAtMs + 90_000,
      meeting_id: 'local-only-meeting',
      url: 'https://zoom.us/j/local-only',
      snapshot: buildMeetingAppFixtureSnapshot('zoom', {
        observedAtMs: observedAtMs + 90_000,
        state: 'prejoin',
        url: 'https://zoom.us/j/local-only',
      }),
    },
  ], {
    id: 'zoom-local-only-record-set',
    createdAtMs: observedAtMs + 91_000,
  }),
});
assert.equal(noSharedIdButTimed.passed, true);
assert.equal(noSharedIdButTimed.status, 'time_matched_identity_unconfirmed');
assert.equal(noSharedIdButTimed.confidence, 'medium');
assert.equal(noSharedIdButTimed.issues.some((item) => item.code === 'identity_not_confirmed'), true);

const kit = createMeetingPlatformTimelineKit({ baseUrl, verify: false });
assert.equal(kit.platformEvidenceCorrelation('google-meet', {
  providerRecords: providerRecords('google-meet'),
  meetingAppRecordSet: meetingAppRecordSet('google-meet'),
}).passed, true);

console.log('ok meeting platform evidence correlation');
