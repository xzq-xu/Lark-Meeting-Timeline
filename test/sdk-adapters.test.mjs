import assert from 'node:assert/strict';
import {
  applyMeetingSignal,
  applyMeetingSignals,
  normalizeMeetingSignal,
} from '../packages/meeting-timeline-sdk/adapters/core.mjs';
import { normalizeGoogleMeetEvent } from '../packages/meeting-timeline-sdk/adapters/google-meet.mjs';
import { normalizeMicrosoftTeamsEvent } from '../packages/meeting-timeline-sdk/adapters/microsoft-teams.mjs';
import { normalizeZoomEvent } from '../packages/meeting-timeline-sdk/adapters/zoom.mjs';

const startMs = 1_782_442_800_000;
const startIso = new Date(startMs).toISOString();
const endIso = new Date(startMs + 60_000).toISOString();

const normalized = normalizeMeetingSignal({
  type: 'meeting_started',
  meeting: {
    platform: 'google_meet',
    meetingId: 'gm-001',
    meetingUrl: 'https://meet.google.com/abc-defg-hij',
  },
  occurredAtMs: startMs,
  source: 'webhook',
});
assert.equal(normalized.type, 'meeting_started');
assert.equal(normalized.meeting.platform, 'google_meet');
assert.equal(normalized.meeting.meeting_id, 'gm-001');
assert.equal(normalized.occurred_at_ms, startMs);

const calls = [];
const client = {
  async startMeeting(input) {
    calls.push({ method: 'startMeeting', input });
    return { ok: true, input };
  },
  async endMeeting(input) {
    calls.push({ method: 'endMeeting', input });
    return { ok: true, input };
  },
  async insertMark(input) {
    calls.push({ method: 'insertMark', input });
    return { ok: true, input };
  },
};

const startResult = await applyMeetingSignal(client, normalized);
assert.equal(startResult.applied, true);
assert.equal(startResult.action, 'startMeeting');
assert.equal(calls.at(-1).input.platform, 'google_meet');
assert.equal(calls.at(-1).input.meeting_id, 'gm-001');
assert.equal(calls.at(-1).input.start_time_ms, startMs);

const endResult = await applyMeetingSignal(client, {
  type: 'meeting_ended',
  meeting: { platform: 'google_meet', meeting_id: 'gm-001' },
  occurred_at_ms: startMs + 60_000,
  source: 'webhook',
});
assert.equal(endResult.action, 'endMeeting');
assert.equal(calls.at(-1).method, 'endMeeting');
assert.equal(calls.at(-1).input.end_time_ms, startMs + 60_000);

const skippedParticipant = await applyMeetingSignal(client, {
  type: 'participant_joined',
  meeting: { platform: 'google_meet', meeting_id: 'gm-001' },
  occurred_at_ms: startMs + 10_000,
  participant_name: 'Ada',
});
assert.equal(skippedParticipant.applied, false);
assert.equal(skippedParticipant.reason, 'participant_track_not_configured');

const participantMark = await applyMeetingSignal(client, {
  type: 'participant_joined',
  meeting: { platform: 'google_meet', meeting_id: 'gm-001' },
  occurred_at_ms: startMs + 10_000,
  participant_name: 'Ada',
}, { participantAsAnnotation: true });
assert.equal(participantMark.applied, true);
assert.equal(calls.at(-1).method, 'insertMark');
assert.equal(calls.at(-1).input.kind, 'participant_joined');
assert.equal(calls.at(-1).input.captured_at_ms, startMs + 10_000);

const batchResults = await applyMeetingSignals(client, [
  {
    type: 'meeting_started',
    meeting: { platform: 'zoom', meeting_id: 'zoom-001' },
    occurred_at_ms: startMs,
  },
  {
    type: 'artifact_ready',
    meeting: { platform: 'zoom', meeting_id: 'zoom-001' },
    occurred_at_ms: startMs + 120_000,
    artifact_kind: 'recording',
  },
]);
assert.equal(batchResults.length, 2);
assert.equal(batchResults[0].action, 'startMeeting');
assert.equal(batchResults[1].applied, false);
assert.equal(batchResults[1].reason, 'artifact_import_not_configured');

const googleStart = normalizeGoogleMeetEvent({
  id: 'g-start-1',
  type: 'google.workspace.meet.conference.v2.started',
  time: startIso,
  data: {
    conferenceRecord: { name: 'conferenceRecords/google-record-001' },
    title: 'Google weekly review',
  },
});
assert.equal(googleStart.length, 1);
assert.equal(googleStart[0].type, 'meeting_started');
assert.equal(googleStart[0].meeting.platform, 'google_meet');
assert.equal(googleStart[0].meeting.meeting_id, 'google-record-001');
assert.equal(googleStart[0].occurred_at_ms, startMs);

const googleParticipant = normalizeGoogleMeetEvent({
  id: 'g-join-1',
  type: 'google.workspace.meet.participant.v2.joined',
  time: startIso,
  data: {
    participantSession: {
      name: 'conferenceRecords/google-record-001/participants/user-123/participantSessions/session-1',
      participant: { displayName: 'Ada Lovelace' },
    },
  },
});
assert.equal(googleParticipant[0].type, 'participant_joined');
assert.equal(googleParticipant[0].participant_id, 'user-123');
assert.equal(googleParticipant[0].participant_name, 'Ada Lovelace');

const googleTranscript = normalizeGoogleMeetEvent({
  id: 'g-transcript-1',
  type: 'google.workspace.meet.transcript.v2.fileGenerated',
  time: endIso,
  data: {
    transcript: {
      name: 'conferenceRecords/google-record-001/transcripts/transcript-1',
      docsDestination: { document: 'https://docs.google.com/document/d/transcript-1' },
    },
  },
});
assert.equal(googleTranscript[0].type, 'artifact_ready');
assert.equal(googleTranscript[0].artifact_kind, 'transcript');
assert.equal(googleTranscript[0].artifact_id, 'transcript-1');

const teamsStart = normalizeMicrosoftTeamsEvent({
  id: 'teams-notification-1',
  resource: 'communications/onlineMeetings(joinWebUrl=https%3A%2F%2Fteams.example%2Fjoin)/meetingCallEvents',
  resourceData: {
    eventType: 'callStarted',
    eventDateTime: startIso,
    onlineMeetingId: 'teams-meeting-001',
    joinWebUrl: 'https://teams.example/join',
    subject: 'Teams review',
  },
});
assert.equal(teamsStart[0].type, 'meeting_started');
assert.equal(teamsStart[0].meeting.platform, 'microsoft_teams');
assert.equal(teamsStart[0].meeting.meeting_id, 'teams-meeting-001');

const teamsRoster = normalizeMicrosoftTeamsEvent({
  id: 'teams-roster-1',
  resourceData: {
    eventType: 'rosterUpdated',
    eventDateTime: startIso,
    onlineMeetingId: 'teams-meeting-001',
    'participants@delta': [
      { id: 'p1', displayName: 'Ada', joinDateTime: startIso },
      { id: 'p2', displayName: 'Grace', removedState: { reason: 'left' }, leaveDateTime: endIso },
    ],
  },
});
assert.deepEqual(teamsRoster.map((item) => item.type), ['participant_joined', 'participant_left']);
assert.equal(teamsRoster[1].participant_name, 'Grace');

const zoomStart = normalizeZoomEvent({
  event: 'meeting.started',
  event_ts: startMs,
  payload: {
    object: {
      uuid: 'zoom-uuid-001',
      id: 987654321,
      topic: 'Zoom review',
      join_url: 'https://zoom.us/j/987654321',
      start_time: startIso,
    },
  },
});
assert.equal(zoomStart[0].type, 'meeting_started');
assert.equal(zoomStart[0].meeting.platform, 'zoom');
assert.equal(zoomStart[0].meeting.meeting_id, 'zoom-uuid-001');
assert.equal(zoomStart[0].meeting.external_meeting_id, '987654321');

const zoomParticipantLeft = normalizeZoomEvent({
  event: 'meeting.participant_left',
  event_ts: startMs + 30_000,
  payload: {
    object: {
      uuid: 'zoom-uuid-001',
      id: 987654321,
      participant: {
        user_id: 'zoom-user-1',
        user_name: 'Lin',
        leave_time: new Date(startMs + 30_000).toISOString(),
      },
    },
  },
});
assert.equal(zoomParticipantLeft[0].type, 'participant_left');
assert.equal(zoomParticipantLeft[0].participant_id, 'zoom-user-1');

const zoomRecording = normalizeZoomEvent({
  event: 'recording.completed',
  event_ts: startMs + 120_000,
  payload: {
    object: {
      uuid: 'zoom-uuid-001',
      id: 987654321,
      recording_files: [{ download_url: 'https://zoom.us/recording/download/1' }],
    },
  },
});
assert.equal(zoomRecording[0].type, 'artifact_ready');
assert.equal(zoomRecording[0].artifact_kind, 'recording');
assert.equal(zoomRecording[0].artifact_url, 'https://zoom.us/recording/download/1');

console.log('ok meeting timeline SDK adapters');
