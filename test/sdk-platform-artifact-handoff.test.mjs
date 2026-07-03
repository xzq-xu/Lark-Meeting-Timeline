import assert from 'node:assert/strict';

import {
  buildMeetingPlatformArtifactHandoff,
  buildMeetingPlatformArtifactHandoffMatrix,
  buildMeetingPlatformArtifactHandoffPlan,
} from '../packages/meeting-timeline-sdk/adapters/platform-artifact-handoff.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const startMs = 1_782_614_400_000;

const googlePlan = buildMeetingPlatformArtifactHandoffPlan('google-meet', {
  baseUrl: 'https://timeline.example.com',
});
assert.equal(googlePlan.schema, 'meeting_platform_artifact_handoff_plan');
assert.equal(googlePlan.platform, 'google_meet');
assert.equal(googlePlan.post_meeting_only, true);
assert.equal(googlePlan.provider_events_block_realtime, false);
assert.equal(googlePlan.transcript_blocks_realtime, false);
assert.equal(googlePlan.transcript_supported, true);
assert.equal(googlePlan.smart_notes_supported, true);
assert.equal(googlePlan.artifacts.find((row) => row.artifact_kind === 'transcript').normalizer, 'normalizeGoogleMeetTranscriptEntries');
assert.equal(googlePlan.import_endpoint, 'https://timeline.example.com/api/import/transcript');
assert.equal(googlePlan.handoff_contract.realtime_annotations_do_not_wait_for_artifacts, true);

const matrix = buildMeetingPlatformArtifactHandoffMatrix({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(matrix.schema, 'meeting_platform_artifact_handoff_matrix');
assert.equal(matrix.platform_count, 5);
assert.equal(matrix.transcript_supported_count, 5);
assert.equal(matrix.recording_supported_count, 5);
assert.equal(matrix.smart_notes_supported_count, 1);
assert.equal(matrix.realtime_blocking_count, 0);
assert.equal(matrix.rows.find((row) => row.platform === 'lark').supported_artifact_kinds.includes('transcript'), true);

const googleHandoff = buildMeetingPlatformArtifactHandoff('google-meet', {
  signals: [
    {
      type: 'artifact_ready',
      meeting: {
        platform: 'google_meet',
        meeting_id: 'google-record-001',
      },
      occurred_at_ms: startMs,
      artifact_kind: 'transcript',
      artifact_id: 'transcript-1',
      source_event_id: 'google-transcript-1',
    },
  ],
}, {
  baseUrl: 'https://timeline.example.com',
  fetchOptions: { accessToken: 'secret-google-token', pageSize: 100 },
});
assert.equal(googleHandoff.schema, 'meeting_platform_artifact_handoff');
assert.equal(googleHandoff.status, 'artifact_handoff_ready');
assert.equal(googleHandoff.import_plan_count, 1);
assert.equal(googleHandoff.fetch_request_count, 1);
assert.equal(googleHandoff.transcript_import_count, 1);
assert.equal(googleHandoff.rows[0].import_endpoint, 'https://timeline.example.com/api/import/transcript');
assert.equal(googleHandoff.rows[0].fetch_request.headers.authorization, '<redacted>');
assert.match(googleHandoff.rows[0].fetch_request.url, /conferenceRecords\/google-record-001\/transcripts\/transcript-1\/entries/);
assert.equal(googleHandoff.post_meeting_only, true);
assert.equal(googleHandoff.transcript_blocks_realtime, false);

const zoomHandoff = buildMeetingPlatformArtifactHandoff('zoom', [
  {
    type: 'artifact_ready',
    meeting: {
      platform: 'zoom',
      meeting_id: '987654321',
    },
    occurred_at_ms: startMs + 60_000,
    artifact_kind: 'transcript',
    artifact_url: 'https://zoom.us/recording/transcript.vtt',
  },
  {
    type: 'artifact_ready',
    meeting: {
      platform: 'zoom',
      meeting_id: '987654321',
    },
    occurred_at_ms: startMs + 70_000,
    artifact_kind: 'whiteboard',
  },
], {
  fetchOptions: { accessToken: 'zoom-token' },
});
assert.equal(zoomHandoff.import_plan_count, 2);
assert.equal(zoomHandoff.fetch_request_count, 1);
assert.equal(zoomHandoff.warning_count, 1);
assert.equal(zoomHandoff.rows.find((row) => row.artifact_kind === 'whiteboard').issue_codes.includes('unsupported_artifact_kind'), true);
assert.equal(zoomHandoff.rows.find((row) => row.artifact_kind === 'transcript').fetch_request.headers.authorization, '<redacted>');
assert.equal(zoomHandoff.next_actions.includes('import_transcripts_without_rebinding_live_annotations'), true);

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
assert.equal(kit.platformArtifactHandoffPlan('webex').schema, 'meeting_platform_artifact_handoff_plan');
assert.equal(kit.platformArtifactHandoffMatrix({ platforms: ['zoom'] }).platform_count, 1);
assert.equal(kit.platformArtifactHandoff('zoom', {
  signals: [{
    type: 'artifact_ready',
    meeting: { platform: 'zoom', meeting_id: '987654321' },
    occurred_at_ms: startMs,
    artifact_kind: 'transcript',
    artifact_url: 'https://zoom.us/transcript.vtt',
  }],
}).fetch_request_count, 1);
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_artifact_handoff_matrix.platform_count, 1);

console.log('ok meeting platform artifact handoff');
