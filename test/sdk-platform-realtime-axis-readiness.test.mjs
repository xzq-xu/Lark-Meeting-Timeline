import assert from 'node:assert/strict';

import {
  buildMeetingAppFixtureSnapshot,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import {
  assertMeetingPlatformRealtimeAxisReadiness,
  assertMeetingPlatformRealtimeAxisReadinessMatrix,
  buildMeetingPlatformRealtimeAxisReadiness,
  buildMeetingPlatformRealtimeAxisReadinessMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-realtime-axis-readiness.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import {
  createMeetingAppTimelineSdk,
} from '../packages/meeting-timeline-sdk/index.mjs';

const baseUrl = 'https://timeline.example.com';

const snapshots = {
  'google-meet': [buildMeetingAppFixtureSnapshot('google-meet', { state: 'active' })],
  webex: [buildMeetingAppFixtureSnapshot('webex', { state: 'active' })],
  lark: [buildMeetingAppFixtureSnapshot('lark', { state: 'active' })],
};

const inputByPlatform = {
  teams: {
    window: {
      id: 'teams-native-main',
      title: 'Microsoft Teams Meeting',
      active: true,
      visible: true,
      inMeeting: true,
      meeting_id: 'teams-123',
    },
    process: { name: 'Microsoft Teams' },
    audio: { call_active: true },
    activeSpeaker: { name: 'Ada Lovelace' },
  },
  zoom: {
    window: {
      id: 'zoom-native-main',
      title: 'Zoom Meeting',
      active: true,
      visible: true,
      inMeeting: true,
      meeting_id: '987654321',
    },
    process: { name: 'Zoom Workplace' },
    audio: { call_active: true },
    activeSpeaker: { name: 'Lin Chen' },
  },
};

const google = buildMeetingPlatformRealtimeAxisReadiness({
  platform: 'google-meet',
  snapshots: snapshots['google-meet'],
}, {
  baseUrl,
  requireSpeakerTrack: true,
});
assert.equal(google.schema, 'meeting_platform_realtime_axis_readiness');
assert.equal(google.accepted, true);
assert.equal(google.status, 'ready_for_realtime_axis');
assert.equal(google.realtime_axis_ready, true);
assert.equal(google.production_handoff_ready, true);
assert.equal(google.axis_source, 'local_observer_axis');
assert.equal(google.axis_surface, 'browser_extension');
assert.equal(google.selected_surface, 'browser_extension');
assert.equal(google.local_observer.realtime_annotation_ready, true);
assert.equal(google.local_observer.meeting_start_ready, true);
assert.equal(google.local_observer.speaker_track_ready, true);
assert.equal(google.provider_replay.accepted, true);
assert.equal(google.provider_replay.coverage.artifact_ready, true);
assert.equal(google.runtime_policy.provider_events_block_realtime, false);
assert.equal(google.runtime_policy.transcript_blocks_realtime, false);
assert.equal(assertMeetingPlatformRealtimeAxisReadiness(google).accepted, true);

const matrix = buildMeetingPlatformRealtimeAxisReadinessMatrix({
  snapshots,
  inputByPlatform,
}, {
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  requireSpeakerTrack: true,
});
assert.equal(matrix.schema, 'meeting_platform_realtime_axis_readiness_matrix');
assert.equal(matrix.platform_count, 5);
assert.equal(matrix.accepted_count, 5);
assert.equal(matrix.realtime_axis_ready_count, 5);
assert.equal(matrix.production_handoff_ready_count, 5);
assert.equal(matrix.local_observer_ready_count, 5);
assert.equal(matrix.provider_replay_accepted_count, 5);
assert.equal(matrix.provider_blocking_count, 0);
assert.equal(matrix.transcript_blocking_count, 0);
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').selected_surface, 'native_detector');
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').axis_source, 'local_observer_axis');
assert.equal(matrix.rows.find((row) => row.platform === 'webex').selected_surface, 'browser_extension');
assert.equal(assertMeetingPlatformRealtimeAxisReadinessMatrix(matrix).accepted_count, 5);

const waiting = buildMeetingPlatformRealtimeAxisReadinessMatrix({}, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(waiting.accepted_count, 0);
assert.equal(waiting.realtime_axis_ready_count, 0);
assert.equal(waiting.provider_replay_accepted_count, 2);
assert.equal(waiting.waiting_for_live_evidence_count, 2);
assert.equal(waiting.rows.find((row) => row.platform === 'google_meet').status, 'waiting_for_live_meeting_evidence');
assert.equal(waiting.rows.find((row) => row.platform === 'google_meet').issue_codes.includes('missing_live_meeting_evidence'), true);
assert.throws(
  () => assertMeetingPlatformRealtimeAxisReadinessMatrix(waiting),
  /realtime axis readiness matrix is not accepted/,
);

const kit = createMeetingPlatformTimelineKit({}, {
  baseUrl,
});
assert.equal(kit.platformRealtimeAxisReadiness({
  platform: 'google-meet',
  snapshots: snapshots['google-meet'],
}, {
  requireSpeakerTrack: true,
}).accepted, true);
assert.equal(kit.platformRealtimeAxisReadinessMatrix({
  snapshots,
  inputByPlatform,
}, {
  platforms: ['google-meet', 'teams', 'zoom'],
  requireSpeakerTrack: true,
}).accepted_count, 3);
assert.equal(kit.assertPlatformRealtimeAxisReadinessMatrix({
  snapshots,
  inputByPlatform,
}, {
  platforms: ['google-meet', 'teams'],
  requireSpeakerTrack: true,
}).accepted_count, 2);
assert.equal(kit.report({
  platforms: ['google-meet'],
  snapshots: {
    'google-meet': snapshots['google-meet'],
  },
  requireSpeakerTrack: true,
}).platform_realtime_axis_readiness_matrix.accepted_count, 1);

const sdk = createMeetingAppTimelineSdk({
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(sdk.platformRealtimeAxisReadiness({
  platform: 'google-meet',
  snapshots: snapshots['google-meet'],
}, {
  requireSpeakerTrack: true,
}).accepted, true);
assert.equal(sdk.realtimeAxisReadinessMatrix({
  snapshots,
  inputByPlatform,
}, {
  platforms: ['google-meet', 'zoom'],
  requireSpeakerTrack: true,
}).accepted_count, 2);
assert.equal(sdk.assertRealtimeAxisReadiness({
  platform: 'google-meet',
  snapshots: snapshots['google-meet'],
}, {
  requireSpeakerTrack: true,
}).accepted, true);

console.log('ok meeting platform realtime axis readiness');
