import assert from 'node:assert/strict';

import {
  buildMeetingPlatformSpeakerTrack,
  buildMeetingPlatformSpeakerTrackMatrix,
  buildMeetingPlatformSpeakerTrackPlan,
} from '../packages/meeting-timeline-sdk/adapters/platform-speaker-track.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const startMs = 1_782_614_400_000;
const meeting = {
  platform: 'google_meet',
  meeting_id: 'abc-defg-hij',
  meeting_url: 'https://meet.google.com/abc-defg-hij',
  title: 'Design review',
};

const ada = (observedAtMs) => ({
  meeting,
  activeSpeaker: {
    id: 'ada',
    name: 'Ada',
    speaking: true,
  },
  observedAtMs,
});
const grace = (observedAtMs) => ({
  meeting,
  activeSpeaker: {
    id: 'grace',
    name: 'Grace',
    speaking: true,
  },
  observedAtMs,
});
const silence = (observedAtMs) => ({
  meeting,
  observedAtMs,
});

const plan = buildMeetingPlatformSpeakerTrackPlan('google-meet');
assert.equal(plan.schema, 'meeting_platform_speaker_track_plan');
assert.equal(plan.platform, 'google_meet');
assert.equal(plan.realtime_ready_when_samples_available, true);
assert.equal(plan.provider_events_block_realtime, false);
assert.equal(plan.transcript_blocks_realtime, false);
assert.equal(plan.track_source_order[0].id, 'local_active_speaker_observer');
assert.equal(plan.output_contract.transcript_required, false);
assert.equal(plan.output_contract.mark_intent, 'speaker_track');

const matrix = buildMeetingPlatformSpeakerTrackMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(matrix.schema, 'meeting_platform_speaker_track_matrix');
assert.equal(matrix.platform_count, 5);
assert.equal(matrix.provider_blocking_count, 0);
assert.equal(matrix.transcript_blocking_count, 0);
assert.equal(matrix.realtime_ready_when_samples_available_count, 5);

const track = buildMeetingPlatformSpeakerTrack('google-meet', {
  samples: [
    ada(startMs),
    ada(startMs + 120),
    ada(startMs + 260),
    silence(startMs + 1_400),
    silence(startMs + 1_950),
    grace(startMs + 2_100),
    grace(startMs + 2_360),
    silence(startMs + 2_450),
    silence(startMs + 3_050),
  ],
}, {
  minStableMs: 250,
  switchStableMs: 250,
  endIdleMs: 500,
  minSegmentMs: 800,
});
assert.equal(track.schema, 'meeting_platform_speaker_track');
assert.equal(track.status, 'speaker_track_ready');
assert.equal(track.generated_signal_count, 4);
assert.equal(track.segment_count, 1);
assert.equal(track.mark_count, 1);
assert.equal(track.segments[0].speaker_id, 'ada');
assert.equal(track.segments[0].duration_ms, 1_400);
assert.equal(track.diagnostics.dropped_short_segment_count, 1);
assert.equal(track.marks[0].intent, 'speaker_track');
assert.equal(track.marks[0].kind, 'speaker_started');
assert.equal(track.marks[0].captured_at_ms, startMs);
assert.equal(track.marks[0].payload.segment.duration_ms, 1_400);
assert.equal(track.marks[0].payload.meeting.meeting_id, 'abc-defg-hij');

const signalTrack = buildMeetingPlatformSpeakerTrack('zoom', {
  signals: [
    {
      type: 'speaker_started',
      meeting: { platform: 'zoom', meeting_id: '987654321' },
      occurred_at_ms: startMs + 10_000,
      speaker_id: 'ada',
      speaker_name: 'Ada',
      source_event_id: 'zoom-speaker-1',
    },
    {
      type: 'speaker_started',
      meeting: { platform: 'zoom', meeting_id: '987654321' },
      occurred_at_ms: startMs + 10_500,
      speaker_id: 'ada',
      speaker_name: 'Ada',
      source_event_id: 'zoom-speaker-dup',
    },
    {
      type: 'speaker_ended',
      meeting: { platform: 'zoom', meeting_id: '987654321' },
      occurred_at_ms: startMs + 12_500,
      speaker_id: 'ada',
      speaker_name: 'Ada',
      source_event_id: 'zoom-speaker-end',
    },
  ],
}, {
  duplicateWindowMs: 1_200,
  minSegmentMs: 500,
});
assert.equal(signalTrack.normalized_signal_count, 3);
assert.equal(signalTrack.diagnostics.dropped_duplicate_count, 1);
assert.equal(signalTrack.segment_count, 1);
assert.equal(signalTrack.segments[0].duration_ms, 2_500);
assert.equal(signalTrack.marks[0].source, 'zoom_speaker_track');
assert.equal(JSON.stringify(signalTrack.marks).includes('transcript'), false);

const client = {
  async startMeeting(input) { return { ok: true, input }; },
  async endMeeting(input) { return { ok: true, input }; },
  async insertMark(input) { return { ok: true, input }; },
  async insertMarks(input) { return { ok: true, input }; },
};
const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl: 'https://timeline.example.com',
  verify: false,
});
assert.equal(kit.platformSpeakerTrackPlan('webex').schema, 'meeting_platform_speaker_track_plan');
assert.equal(kit.platformSpeakerTrackMatrix({ platforms: ['zoom'] }).platform_count, 1);
assert.equal(kit.platformSpeakerTrack('google-meet', { samples: [ada(startMs), ada(startMs + 300)] }, {
  minStableMs: 250,
}).mark_count, 0);
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_speaker_track_matrix.platform_count, 1);

console.log('ok meeting platform speaker track');
