import assert from 'node:assert/strict';

import {
  buildMeetingAppFixtureSnapshot,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import {
  assertMeetingPlatformAdapterPreflight,
  assertMeetingPlatformAdapterPreflightMatrix,
  buildMeetingPlatformAdapterPreflight,
  buildMeetingPlatformAdapterPreflightMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-preflight.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import {
  createMeetingAppTimelineSdk,
} from '../packages/meeting-timeline-sdk/index.mjs';

const baseUrl = 'https://timeline.example.com';

const googleActive = buildMeetingAppFixtureSnapshot('google-meet', {
  state: 'active',
  observedAtMs: 1_783_356_000_000,
});
const googleEnded = buildMeetingAppFixtureSnapshot('google-meet', {
  state: 'prejoin',
  observedAtMs: 1_783_356_600_000,
});

const urlOnly = buildMeetingPlatformAdapterPreflight({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Google Meet',
}, {
  baseUrl,
});
assert.equal(urlOnly.schema, 'meeting_platform_adapter_preflight');
assert.equal(urlOnly.platform, 'google_meet');
assert.equal(urlOnly.startup.realtime_startup_ready, true);
assert.equal(urlOnly.readiness.static_startup_ready, true);
assert.equal(urlOnly.readiness.live_evidence_ready, false);
assert.equal(urlOnly.accepted, false);
assert.equal(urlOnly.status, 'needs_live_page_evidence');
assert.equal(urlOnly.issues.some((issue) => issue.code === 'missing_live_page_evidence'), true);
assert.equal(urlOnly.next_actions.includes('collect_live_dom_snapshots_from_current_meeting_window'), true);

const googleLive = buildMeetingPlatformAdapterPreflight({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Google Meet',
  snapshots: [googleActive],
}, {
  baseUrl,
  requireSpeakerTrack: true,
});
assert.equal(googleLive.accepted, true);
assert.equal(googleLive.status, 'ready_for_realtime_annotations');
assert.equal(googleLive.readiness.realtime_annotation_ready, true);
assert.equal(googleLive.readiness.speaker_track_ready, true);
assert.equal(googleLive.readiness.meeting_start_ready, true);
assert.equal(googleLive.readiness.meeting_end_ready, false);
assert.equal(googleLive.readiness.production_lifecycle_ready, false);
assert.equal(googleLive.summary.active_speaker_matched, true);
assert.equal(googleLive.startup.runtime_contract.transcript_blocks_realtime, false);
assert.equal(googleLive.next_actions.includes('open_adapter_session_and_insert_marks_with_captured_at_ms'), true);
assert.equal(assertMeetingPlatformAdapterPreflight({
  platform: 'google-meet',
  snapshots: [googleActive],
}, {
  baseUrl,
  requireSpeakerTrack: true,
}).accepted, true);

const googleLifecycle = buildMeetingPlatformAdapterPreflight({
  platform: 'google-meet',
  snapshots: [googleActive, googleEnded],
}, {
  baseUrl,
  requireSpeakerTrack: true,
  requireCompleteLifecycle: true,
});
assert.equal(googleLifecycle.accepted, true);
assert.equal(googleLifecycle.status, 'ready_for_realtime_annotations_and_lifecycle');
assert.equal(googleLifecycle.readiness.meeting_end_ready, true);
assert.equal(googleLifecycle.readiness.production_lifecycle_ready, true);

const teamsActive = buildMeetingAppFixtureSnapshot('teams', {
  state: 'active',
  observedAtMs: 1_783_356_010_000,
});
const matrix = buildMeetingPlatformAdapterPreflightMatrix({}, {
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom'],
  snapshots: {
    'google-meet': [googleActive],
    teams: [teamsActive],
  },
});
assert.equal(matrix.schema, 'meeting_platform_adapter_preflight_matrix');
assert.equal(matrix.platform_count, 3);
assert.equal(matrix.accepted_count, 2);
assert.equal(matrix.realtime_ready_count, 2);
assert.equal(matrix.live_evidence_ready_count, 2);
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').status, 'needs_live_page_evidence');
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').realtime_annotation_ready, true);
const inputSnapshotMatrix = buildMeetingPlatformAdapterPreflightMatrix({
  snapshots: {
    'google-meet': [googleActive],
  },
}, {
  baseUrl,
  platforms: ['google-meet'],
});
assert.equal(inputSnapshotMatrix.accepted_count, 1);
assert.equal(inputSnapshotMatrix.rows[0].live_evidence_ready, true);
assert.throws(
  () => assertMeetingPlatformAdapterPreflightMatrix({}, {
    baseUrl,
    platforms: ['google-meet', 'zoom'],
    snapshots: {
      'google-meet': [googleActive],
    },
  }),
  /Meeting platform adapter preflight matrix is not accepted/,
);
assert.equal(assertMeetingPlatformAdapterPreflightMatrix({}, {
  baseUrl,
  platforms: ['google-meet'],
  snapshots: {
    'google-meet': [googleActive, googleEnded],
  },
  requireCompleteLifecycle: true,
}).accepted_count, 1);

const kit = createMeetingPlatformTimelineKit({}, {
  baseUrl,
});
assert.equal(kit.platformAdapterPreflight({
  platform: 'zoom',
}).status, 'needs_live_page_evidence');
assert.equal(kit.platformAdapterPreflightMatrix({}, {
  platforms: ['google-meet'],
  snapshots: {
    'google-meet': [googleActive],
  },
}).realtime_ready_count, 1);

const sdk = createMeetingAppTimelineSdk({
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(sdk.platformAdapterPreflight({
  url: 'https://meet.google.com/abc-defg-hij',
  snapshots: [googleActive],
}).accepted, true);
assert.equal(sdk.adapterPreflight({
  platform: 'google-meet',
  snapshots: [googleActive],
}).readiness.realtime_annotation_ready, true);
assert.equal(sdk.platformAdapterPreflightMatrix({}, {
  platforms: ['google-meet', 'zoom'],
  snapshots: {
    'google-meet': [googleActive],
  },
}).accepted_count, 1);
assert.equal(sdk.assertAdapterPreflight({
  platform: 'google-meet',
  snapshots: [googleActive],
}).accepted, true);

console.log('ok meeting platform adapter preflight');
