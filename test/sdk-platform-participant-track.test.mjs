import assert from 'node:assert/strict';

import {
  buildMeetingPlatformParticipantTrack,
  buildMeetingPlatformParticipantTrackMatrix,
  buildMeetingPlatformParticipantTrackPlan,
} from '../packages/meeting-timeline-sdk/adapters/platform-participant-track.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const startMs = 1_782_614_400_000;
const meeting = {
  platform: 'google_meet',
  meeting_id: 'abc-defg-hij',
  meeting_url: 'https://meet.google.com/abc-defg-hij',
  title: 'Design review',
};

const plan = buildMeetingPlatformParticipantTrackPlan('google-meet');
assert.equal(plan.schema, 'meeting_platform_participant_track_plan');
assert.equal(plan.platform, 'google_meet');
assert.equal(plan.realtime_ready_when_snapshots_available, true);
assert.equal(plan.provider_events_block_realtime, false);
assert.equal(plan.transcript_blocks_realtime, false);
assert.equal(plan.track_source_order[0].id, 'local_roster_observer');
assert.equal(plan.output_contract.transcript_required, false);
assert.equal(plan.output_contract.mark_intent, 'participant_track');

const matrix = buildMeetingPlatformParticipantTrackMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(matrix.schema, 'meeting_platform_participant_track_matrix');
assert.equal(matrix.platform_count, 5);
assert.equal(matrix.provider_blocking_count, 0);
assert.equal(matrix.transcript_blocking_count, 0);
assert.equal(matrix.realtime_ready_when_snapshots_available_count, 5);

const signalTrack = buildMeetingPlatformParticipantTrack('zoom', {
  signals: [
    {
      type: 'participant_joined',
      meeting: { platform: 'zoom', meeting_id: '987654321' },
      occurred_at_ms: startMs,
      participant_id: 'ada',
      participant_name: 'Ada',
      source_event_id: 'zoom-join-1',
    },
    {
      type: 'participant_joined',
      meeting: { platform: 'zoom', meeting_id: '987654321' },
      occurred_at_ms: startMs + 300,
      participant_id: 'ada',
      participant_name: 'Ada',
      source_event_id: 'zoom-join-dup',
    },
    {
      type: 'participant_left',
      meeting: { platform: 'zoom', meeting_id: '987654321' },
      occurred_at_ms: startMs + 3_000,
      participant_id: 'ada',
      participant_name: 'Ada',
      source_event_id: 'zoom-left-reconnect',
    },
    {
      type: 'participant_joined',
      meeting: { platform: 'zoom', meeting_id: '987654321' },
      occurred_at_ms: startMs + 4_500,
      participant_id: 'ada',
      participant_name: 'Ada',
      source_event_id: 'zoom-rejoin',
    },
    {
      type: 'participant_left',
      meeting: { platform: 'zoom', meeting_id: '987654321' },
      occurred_at_ms: startMs + 20_000,
      participant_id: 'ada',
      participant_name: 'Ada',
      source_event_id: 'zoom-left-final',
    },
  ],
}, {
  duplicateWindowMs: 1_000,
  suppressReconnectGapMs: 5_000,
});
assert.equal(signalTrack.schema, 'meeting_platform_participant_track');
assert.equal(signalTrack.status, 'participant_track_ready');
assert.equal(signalTrack.normalized_signal_count, 5);
assert.equal(signalTrack.diagnostics.dropped_duplicate_count, 1);
assert.equal(signalTrack.diagnostics.suppressed_reconnect_pair_count, 1);
assert.equal(signalTrack.mark_count, 2);
assert.deepEqual(signalTrack.marks.map((mark) => mark.kind), ['participant_joined', 'participant_left']);
assert.equal(signalTrack.marks[0].captured_at_ms, startMs);
assert.equal(signalTrack.marks[0].intent, 'participant_track');
assert.equal(signalTrack.marks[0].payload.participant_name, 'Ada');
assert.equal(JSON.stringify(signalTrack.marks).includes('transcript'), false);

const snapshotTrack = buildMeetingPlatformParticipantTrack('google-meet', {
  snapshots: [
    {
      meeting,
      observed_at_ms: startMs,
      participants: [{ id: 'ada', name: 'Ada' }],
    },
    {
      meeting,
      observed_at_ms: startMs + 1_000,
      participants: [{ id: 'ada', name: 'Ada' }, { id: 'bob', name: 'Bob' }],
    },
    {
      meeting,
      observed_at_ms: startMs + 2_000,
      participants: [{ id: 'ada', name: 'Ada' }],
    },
    {
      meeting,
      observed_at_ms: startMs + 2_500,
      participants: [{ id: 'ada', name: 'Ada' }, { id: 'bob', name: 'Bob' }],
    },
    {
      meeting,
      observed_at_ms: startMs + 4_000,
      participants: [{ id: 'ada', name: 'Ada' }],
    },
    {
      meeting,
      observed_at_ms: startMs + 5_700,
      participants: [{ id: 'ada', name: 'Ada' }],
    },
  ],
}, {
  leaveStableMs: 1_000,
  suppressReconnectGapMs: 200,
});
assert.equal(snapshotTrack.input_snapshot_count, 6);
assert.equal(snapshotTrack.generated_signal_count, 2);
assert.equal(snapshotTrack.diagnostics.suppressed_pending_leave_count, 1);
assert.equal(snapshotTrack.diagnostics.emitted_pending_leave_count, 1);
assert.equal(snapshotTrack.mark_count, 2);
assert.deepEqual(snapshotTrack.marks.map((mark) => mark.label), ['Bob joined', 'Bob left']);
assert.equal(snapshotTrack.marks[0].captured_at_ms, startMs + 1_000);
assert.equal(snapshotTrack.marks[1].captured_at_ms, startMs + 4_000);

const initialRosterTrack = buildMeetingPlatformParticipantTrack('google-meet', {
  snapshots: [
    {
      meeting,
      observed_at_ms: startMs,
      participants: [{ id: 'ada', name: 'Ada' }, { id: 'bob', name: 'Bob' }],
    },
  ],
}, {
  emitInitialRoster: false,
});
assert.equal(initialRosterTrack.mark_count, 0);

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
assert.equal(kit.platformParticipantTrackPlan('webex').schema, 'meeting_platform_participant_track_plan');
assert.equal(kit.platformParticipantTrackMatrix({ platforms: ['zoom'] }).platform_count, 1);
assert.equal(kit.platformParticipantTrack('google-meet', {
  signals: [
    {
      type: 'participant_joined',
      meeting,
      occurred_at_ms: startMs,
      participant_id: 'ada',
      participant_name: 'Ada',
    },
  ],
}).mark_count, 1);
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_participant_track_matrix.platform_count, 1);

console.log('ok meeting platform participant track');
