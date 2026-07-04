import assert from 'node:assert/strict';

import {
  MEETING_APP_TRACK_PIPELINE_SCHEMA,
  buildMeetingAppTrackPipeline,
  createMeetingAppTrackPipeline,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-track-pipeline.mjs';

const startMs = 1_783_356_000_000;

function snapshot(observedAtMs, {
  activeSpeaker = null,
  participants = [],
} = {}) {
  return {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    title: 'Design review - Google Meet',
    observedAtMs,
    inMeeting: true,
    activeSpeaker,
    participants,
  };
}

const ada = { id: 'ada', name: 'Ada Lovelace', speaking: true };
const bob = { id: 'bob', name: 'Bob Stone', speaking: false };
const grace = { id: 'grace', name: 'Grace Hopper', speaking: true };

const snapshots = [
  snapshot(startMs, {
    activeSpeaker: ada,
    participants: [ada, bob],
  }),
  snapshot(startMs + 300, {
    activeSpeaker: ada,
    participants: [ada, bob],
  }),
  snapshot(startMs + 1_400, {
    participants: [{ ...ada, speaking: false }, bob],
  }),
  snapshot(startMs + 2_000, {
    activeSpeaker: grace,
    participants: [{ ...ada, speaking: false }, bob, grace],
  }),
  snapshot(startMs + 2_300, {
    activeSpeaker: grace,
    participants: [{ ...ada, speaking: false }, bob, grace],
  }),
  snapshot(startMs + 3_300, {
    participants: [{ ...ada, speaking: false }, { ...grace, speaking: false }],
  }),
  snapshot(startMs + 4_000, {
    participants: [{ ...ada, speaking: false }, { ...grace, speaking: false }],
  }),
];

const report = buildMeetingAppTrackPipeline(snapshots, {
  speakerTrackOptions: {
    minStableMs: 250,
    switchStableMs: 250,
    endIdleMs: 500,
    minSegmentMs: 800,
  },
  participantTrackOptions: {
    leaveStableMs: 500,
    closePendingLeavesAtMs: startMs + 4_500,
  },
});

assert.equal(MEETING_APP_TRACK_PIPELINE_SCHEMA, 'meeting_app_track_pipeline');
assert.equal(report.schema, 'meeting_app_track_pipeline');
assert.equal(report.status, 'track_marks_ready');
assert.equal(report.platform, 'google_meet');
assert.equal(report.meeting.meeting_id, 'abc-defg-hij');
assert.equal(report.coverage.normalized_snapshot_count, 7);
assert.equal(report.coverage.scoped_snapshot_count, 7);
assert.equal(report.coverage.transcript_required, false);
assert.equal(report.coverage.provider_event_required, false);
assert.equal(report.speaker_track.mark_count, 2);
assert.equal(report.participant_track.mark_count, 2);
assert.deepEqual(report.speaker_track.marks.map((mark) => mark.payload.speaker_name), [
  'Ada Lovelace',
  'Grace Hopper',
]);
assert.deepEqual(report.participant_track.marks.map((mark) => `${mark.payload.participant_name}:${mark.kind}`), [
  'Grace Hopper:participant_joined',
  'Bob Stone:participant_left',
]);
assert.deepEqual(report.marks.map((mark) => mark.intent), [
  'speaker_track',
  'participant_track',
  'speaker_track',
  'participant_track',
]);
assert.equal(report.marks[0].captured_at_ms, startMs);
assert.equal(JSON.stringify(report.marks).includes('transcript'), false);

const accumulator = createMeetingAppTrackPipeline({
  speakerTrackOptions: {
    minStableMs: 250,
    switchStableMs: 250,
    endIdleMs: 500,
    minSegmentMs: 800,
  },
  participantTrackOptions: {
    leaveStableMs: 500,
    closePendingLeavesAtMs: startMs + 4_500,
  },
});
let streamed = accumulator.push(snapshots.slice(0, 2));
assert.equal(streamed.speaker_track.mark_count, 0);
streamed = accumulator.push(snapshots.slice(2));
assert.equal(streamed.mark_count, report.mark_count);
assert.equal(accumulator.getSnapshots().length, snapshots.length);
assert.equal(accumulator.reset([]).mark_count, 0);

const teamsReport = buildMeetingAppTrackPipeline([
  {
    platform: 'microsoft_teams',
    meeting_id: 'teams-call-1',
    meeting_url: 'https://teams.microsoft.com/l/meetup-join/19%3Ameeting_abc%40thread.v2/0',
    observedAtMs: startMs,
    activeSpeaker: { id: 'mira', name: 'Mira Patel', speaking: true },
    participants: [{ id: 'mira', name: 'Mira Patel', speaking: true }],
  },
  {
    platform: 'microsoft_teams',
    meeting_id: 'teams-call-1',
    meeting_url: 'https://teams.microsoft.com/l/meetup-join/19%3Ameeting_abc%40thread.v2/0',
    observedAtMs: startMs + 400,
    activeSpeaker: { id: 'mira', name: 'Mira Patel', speaking: true },
    participants: [{ id: 'mira', name: 'Mira Patel', speaking: true }],
  },
], {
  speakerTrackOptions: {
    minStableMs: 250,
    minSegmentMs: 0,
    closeOpenSegmentsAtMs: startMs + 1_000,
  },
});
assert.equal(teamsReport.platform, 'microsoft_teams');
assert.equal(teamsReport.speaker_track.mark_count, 1);
assert.equal(teamsReport.participant_track.mark_count, 0);

console.log('ok meeting app track pipeline');
