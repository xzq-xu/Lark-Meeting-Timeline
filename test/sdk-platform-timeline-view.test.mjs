import assert from 'node:assert/strict';

import {
  buildMeetingPlatformTimelineView,
  buildMeetingPlatformTimelineViewMatrix,
  buildMeetingPlatformTimelineViewPlan,
  zoomMeetingPlatformTimelineViewport,
} from '../packages/meeting-timeline-sdk/adapters/platform-timeline-view.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const startMs = 1_782_614_400_000;
const meeting = {
  platform: 'google_meet',
  meeting_id: 'abc-defg-hij',
  meeting_url: 'https://meet.google.com/abc-defg-hij',
  title: 'Design review',
  start_time_ms: startMs,
  end_time_ms: startMs + 10 * 60_000,
};

const plan = buildMeetingPlatformTimelineViewPlan('google-meet');
assert.equal(plan.schema, 'meeting_platform_timeline_view_plan');
assert.equal(plan.platform, 'google_meet');
assert.equal(plan.output_contract.renderer_agnostic, true);
assert.equal(plan.realtime_policy.provider_events_block_realtime, false);
assert.equal(plan.realtime_policy.transcript_blocks_realtime, false);
assert.equal(plan.rails[0].id, 'speaker');

const matrix = buildMeetingPlatformTimelineViewMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(matrix.schema, 'meeting_platform_timeline_view_matrix');
assert.equal(matrix.platform_count, 5);
assert.equal(matrix.renderer_agnostic_count, 5);
assert.equal(matrix.provider_blocking_count, 0);
assert.equal(matrix.transcript_blocking_count, 0);

const view = buildMeetingPlatformTimelineView('google-meet', {
  meeting,
  annotations: [
    {
      id: 'mark-why',
      label: 'why?',
      captured_at_ms: startMs + 90_000,
      kind: 'handwriting',
      intent: 'attention',
    },
    {
      id: 'mark-after-end',
      label: 'late note',
      captured_at_ms: startMs + 12 * 60_000,
      kind: 'handwriting',
    },
    {
      id: 'mark-uncalibrated',
      label: 'pending note',
      kind: 'handwriting',
    },
  ],
  events: [
    {
      type: 'meeting_started',
      meeting,
      occurred_at_ms: startMs,
    },
  ],
  speakerTrack: {
    marks: [
      {
        id: 'speaker-ada',
        intent: 'speaker_track',
        kind: 'speaker_started',
        label: 'Ada speaking',
        captured_at_ms: startMs + 120_000,
      },
    ],
  },
  participantTrack: {
    marks: [
      {
        id: 'participant-bob',
        intent: 'participant_track',
        kind: 'participant_joined',
        label: 'Bob joined',
        captured_at_ms: startMs + 180_000,
      },
    ],
  },
  artifactHandoff: {
    rows: [
      {
        artifact_kind: 'transcript',
        status: 'requires_provider_fetch',
        occurred_at_ms: startMs + 11 * 60_000,
      },
    ],
  },
  transcript: {
    segments: [
      {
        id: 'transcript-1',
        start_ms: 240_000,
        text: 'Hello world',
      },
    ],
  },
}, {
  viewportStartMs: 60_000,
  viewportDurationMs: 240_000,
});

assert.equal(view.schema, 'meeting_platform_timeline_view');
assert.equal(view.status, 'timeline_view_ready');
assert.equal(view.meeting.duration_ms, 600_000);
assert.equal(view.viewport.start_ms, 60_000);
assert.equal(view.viewport.end_ms, 300_000);
assert.equal(view.diagnostics.marker_count, 8);
assert.equal(view.diagnostics.visible_marker_count, 4);
assert.equal(view.diagnostics.uncalibrated_marker_count, 1);
assert.equal(view.diagnostics.warning_count, 3);
assert.equal(view.diagnostics.rail_counts.annotations, 3);
assert.equal(view.diagnostics.rail_counts.speaker, 1);
assert.equal(view.diagnostics.rail_counts.participant, 1);
assert.equal(view.diagnostics.rail_counts.artifacts, 1);
assert.equal(view.diagnostics.rail_counts.transcript, 1);
assert.equal(view.markers.find((marker) => marker.id === 'mark-why').x_ratio, 0.125);
assert.equal(view.markers.find((marker) => marker.id === 'mark-after-end').warnings.includes('after_meeting_end'), true);
assert.equal(view.markers.find((marker) => marker.id === 'mark-uncalibrated').warnings.includes('missing_relative_time'), true);
assert.equal(view.markers.find((marker) => marker.id.startsWith('artifact_handoff')).rail, 'artifacts');
assert.equal(view.visible_markers.some((marker) => marker.id === 'speaker-ada'), true);
assert.equal(view.visible_markers.some((marker) => marker.id === 'mark-after-end'), false);
assert.equal(view.ticks[0].ms, 60_000);
assert.equal(view.ticks.at(-1).ms, 300_000);

const zoomed = zoomMeetingPlatformTimelineViewport(view.viewport, 2, 0.5);
assert.equal(zoomed.duration_ms, 120_000);
assert.equal(zoomed.start_ms, 120_000);
assert.equal(zoomed.end_ms, 240_000);

const emptyView = buildMeetingPlatformTimelineView('zoom', {
  meeting: {
    platform: 'zoom',
    meeting_id: '987654321',
    start_time_ms: startMs,
  },
});
assert.equal(emptyView.status, 'empty_timeline_view');
assert.equal(emptyView.diagnostics.marker_count, 0);

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
assert.equal(kit.platformTimelineViewPlan('webex').schema, 'meeting_platform_timeline_view_plan');
assert.equal(kit.platformTimelineViewMatrix({ platforms: ['zoom'] }).platform_count, 1);
assert.equal(kit.platformTimelineView('google-meet', {
  meeting,
  annotations: [{ id: 'a1', label: 'now', captured_at_ms: startMs + 1_000 }],
}).diagnostics.marker_count, 1);
assert.equal(kit.zoomPlatformTimelineViewport({ full_duration_ms: 600_000, start_ms: 0, duration_ms: 600_000 }, 2).duration_ms, 300_000);
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_timeline_view_matrix.platform_count, 1);

console.log('ok meeting platform timeline view');
