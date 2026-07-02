import assert from 'node:assert/strict';
import {
  applyMeetingSignal,
  applyMeetingSignals,
  normalizeMeetingSignal,
} from '../packages/meeting-timeline-sdk/adapters/core.mjs';
import { normalizeGoogleMeetEvent, unwrapGooglePubSubEvent } from '../packages/meeting-timeline-sdk/adapters/google-meet.mjs';
import { normalizeLarkEvent } from '../packages/meeting-timeline-sdk/adapters/lark.mjs';
import { normalizeLocalDetectorEvent } from '../packages/meeting-timeline-sdk/adapters/local-detector.mjs';
import { normalizeMicrosoftTeamsEvent } from '../packages/meeting-timeline-sdk/adapters/microsoft-teams.mjs';
import { normalizeWebexEvent } from '../packages/meeting-timeline-sdk/adapters/webex.mjs';
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

const skippedSpeaker = await applyMeetingSignal(client, {
  type: 'speaker_started',
  meeting: { platform: 'google_meet', meeting_id: 'gm-001' },
  occurred_at_ms: startMs + 20_000,
  speaker_name: 'Ada',
});
assert.equal(skippedSpeaker.applied, false);
assert.equal(skippedSpeaker.reason, 'speaker_track_not_configured');

const speakerMark = await applyMeetingSignal(client, {
  type: 'speaker_started',
  meeting: { platform: 'google_meet', meeting_id: 'gm-001' },
  occurred_at_ms: startMs + 20_000,
  speaker_id: 'speaker-ada',
  speaker_name: 'Ada',
}, { speakerAsAnnotation: true });
assert.equal(speakerMark.applied, true);
assert.equal(speakerMark.action, 'insertSpeakerMark');
assert.equal(calls.at(-1).method, 'insertMark');
assert.equal(calls.at(-1).input.kind, 'speaker_started');
assert.equal(calls.at(-1).input.label, 'Ada speaking');
assert.equal(calls.at(-1).input.captured_at_ms, startMs + 20_000);
assert.equal(calls.at(-1).input.payload.speaker_id, 'speaker-ada');

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

const lifecycleSignal = normalizeMeetingSignal({
  type: 'subscription_lifecycle',
  platform: 'google_meet',
  occurred_at_ms: startMs,
  lifecycle_type: 'expirationReminder',
  subscription_name: 'subscriptions/google-sub-001',
  expires_at: endIso,
  source: 'webhook',
});
assert.equal(lifecycleSignal.meeting, undefined);
assert.equal(lifecycleSignal.platform, 'google_meet');
assert.equal(lifecycleSignal.lifecycle_type, 'expiration_reminder');
assert.equal(lifecycleSignal.subscription_name, 'subscriptions/google-sub-001');
assert.equal(lifecycleSignal.expires_at_ms, startMs + 60_000);

const skippedLifecycle = await applyMeetingSignal(client, lifecycleSignal);
assert.equal(skippedLifecycle.applied, false);
assert.equal(skippedLifecycle.reason, 'subscription_lifecycle_not_configured');

const appliedLifecycle = await applyMeetingSignal(client, lifecycleSignal, {
  onSubscriptionLifecycleSignal(signal) {
    return { ok: true, action: 'renew', subscription: signal.subscription_name };
  },
});
assert.equal(appliedLifecycle.applied, true);
assert.equal(appliedLifecycle.action, 'onSubscriptionLifecycleSignal');
assert.equal(appliedLifecycle.response.subscription, 'subscriptions/google-sub-001');

const localDetectorStart = normalizeLocalDetectorEvent({
  id: 'local-start-1',
  type: 'meeting_started',
  detected_platform: 'google_meet',
  meeting_id: 'local-google-meet-001',
  meeting_url: 'https://meet.google.com/local-demo',
  title: 'Local detector Google Meet',
  start_time_ms: startMs,
});
assert.equal(localDetectorStart.length, 1);
assert.equal(localDetectorStart[0].type, 'meeting_started');
assert.equal(localDetectorStart[0].meeting.platform, 'google_meet');
assert.equal(localDetectorStart[0].meeting.meeting_id, 'local-google-meet-001');
assert.equal(localDetectorStart[0].source, 'local_detector');
assert.equal(localDetectorStart[0].occurred_at_ms, startMs);

const localDetectorStartResult = await applyMeetingSignal(client, localDetectorStart[0]);
assert.equal(localDetectorStartResult.action, 'startMeeting');
assert.equal(calls.at(-1).input.platform, 'google_meet');
assert.equal(calls.at(-1).input.detector_source, 'google_meet_local_detector');
assert.equal(calls.at(-1).input.start_time_ms, startMs);

const localDetectorEnd = normalizeLocalDetectorEvent({
  type: 'meeting_ended',
  meeting: {
    platform: 'google_meet',
    id: 'local-google-meet-001',
    endTimeMs: startMs + 300_000,
  },
});
assert.equal(localDetectorEnd[0].type, 'meeting_ended');
assert.equal(localDetectorEnd[0].occurred_at_ms, startMs + 300_000);

const localDetectorSpeaker = normalizeLocalDetectorEvent({
  id: 'local-speaker-1',
  type: 'active_speaker',
  detected_platform: 'google_meet',
  meeting_id: 'local-google-meet-001',
  occurred_at_ms: startMs + 45_000,
  speaker: {
    id: 'local-speaker-ada',
    name: 'Ada Lovelace',
  },
});
assert.equal(localDetectorSpeaker[0].type, 'speaker_started');
assert.equal(localDetectorSpeaker[0].speaker_id, 'local-speaker-ada');
assert.equal(localDetectorSpeaker[0].speaker_name, 'Ada Lovelace');
assert.equal(localDetectorSpeaker[0].occurred_at_ms, startMs + 45_000);

const larkStart = normalizeLarkEvent({
  header: {
    event_id: 'lark-start-1',
    event_type: 'vc.meeting.all_meeting_started_v1',
    create_time: String(startMs),
  },
  event: {
    meeting: {
      id: 'lark-meeting-001',
      meeting_no: '123456789',
      topic: 'Lark product review',
      url: 'https://vc.feishu.cn/j/lark-meeting-001',
      start_time: String(Math.round(startMs / 1000)),
    },
    minute_token: 'minute-token-001',
  },
});
assert.equal(larkStart.length, 1);
assert.equal(larkStart[0].type, 'meeting_started');
assert.equal(larkStart[0].meeting.platform, 'lark');
assert.equal(larkStart[0].meeting.meeting_id, 'lark-meeting-001');
assert.equal(larkStart[0].meeting.external_meeting_id, '123456789');
assert.equal(larkStart[0].meeting.minute_token, 'minute-token-001');
assert.equal(larkStart[0].occurred_at_ms, startMs);

const larkStartResult = await applyMeetingSignal(client, larkStart[0]);
assert.equal(larkStartResult.action, 'startMeeting');
assert.equal(calls.at(-1).input.minute_token, 'minute-token-001');

const larkJoin = normalizeLarkEvent({
  header: {
    event_id: 'lark-join-1',
    event_type: 'vc.meeting.join_meeting_v1',
    create_time: String(startMs + 60_000),
  },
  event: {
    meeting: {
      id: 'lark-meeting-001',
      topic: 'Lark product review',
      start_time: String(Math.round(startMs / 1000)),
    },
    user: {
      open_id: 'ou_001',
      name: 'Ada',
    },
  },
});
assert.equal(larkJoin[0].type, 'participant_joined');
assert.equal(larkJoin[0].participant_id, 'ou_001');
assert.equal(larkJoin[0].participant_name, 'Ada');

const larkEnd = normalizeLarkEvent({
  header: {
    event_id: 'lark-end-1',
    event_type: 'vc.meeting.all_meeting_ended_v1',
  },
  event: {
    meeting: {
      id: 'lark-meeting-001',
      start_time: String(Math.round(startMs / 1000)),
      end_time: String(Math.round((startMs + 600_000) / 1000)),
    },
  },
});
assert.equal(larkEnd[0].type, 'meeting_ended');
assert.equal(larkEnd[0].occurred_at_ms, startMs + 600_000);

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

const googlePubSubCloudEvent = {
  id: 'g-pubsub-start-1',
  type: 'google.workspace.meet.conference.v2.started',
  time: startIso,
  data: {
    conferenceRecord: { name: 'conferenceRecords/google-record-pubsub-001' },
    title: 'Google Pub/Sub weekly review',
  },
};
const googlePubSubWrapped = {
  message: {
    messageId: 'pubsub-message-001',
    publishTime: startIso,
    data: Buffer.from(JSON.stringify(googlePubSubCloudEvent), 'utf8').toString('base64'),
    attributes: { eventType: googlePubSubCloudEvent.type },
  },
  subscription: 'projects/demo-project/subscriptions/meet-events',
};
const googlePubSubUnwrapped = unwrapGooglePubSubEvent(googlePubSubWrapped);
assert.equal(googlePubSubUnwrapped.id, 'g-pubsub-start-1');
assert.equal(googlePubSubUnwrapped.pubsub.message_id, 'pubsub-message-001');
const googlePubSubStart = normalizeGoogleMeetEvent(googlePubSubWrapped);
assert.equal(googlePubSubStart.length, 1);
assert.equal(googlePubSubStart[0].type, 'meeting_started');
assert.equal(googlePubSubStart[0].meeting.meeting_id, 'google-record-pubsub-001');
assert.equal(googlePubSubStart[0].source_event_id, 'g-pubsub-start-1');

const googleLifecycle = normalizeGoogleMeetEvent({
  id: 'g-lifecycle-1',
  type: 'google.workspace.events.subscription.v1.expirationReminder',
  time: startIso,
  data: {
    subscription: {
      name: 'subscriptions/google-sub-001',
      expire_time: endIso,
    },
  },
});
assert.equal(googleLifecycle.length, 1);
assert.equal(googleLifecycle[0].type, 'subscription_lifecycle');
assert.equal(googleLifecycle[0].platform, 'google_meet');
assert.equal(googleLifecycle[0].lifecycle_type, 'expiration_reminder');
assert.equal(googleLifecycle[0].subscription_id, 'google-sub-001');
assert.equal(googleLifecycle[0].expires_at_ms, startMs + 60_000);

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

const teamsLifecycle = normalizeMicrosoftTeamsEvent({
  value: [
    {
      id: 'teams-lifecycle-1',
      lifecycleEvent: 'reauthorizationRequired',
      subscriptionId: 'teams-sub-001',
      subscriptionExpirationDateTime: endIso,
      resource: 'communications/onlineMeetings/example/meetingCallEvents',
      tenantId: 'tenant-001',
      clientState: 'opaque-secret',
    },
  ],
}, { receivedAtMs: startMs });
assert.equal(teamsLifecycle.length, 1);
assert.equal(teamsLifecycle[0].type, 'subscription_lifecycle');
assert.equal(teamsLifecycle[0].platform, 'microsoft_teams');
assert.equal(teamsLifecycle[0].lifecycle_type, 'reauthorization_required');
assert.equal(teamsLifecycle[0].subscription_id, 'teams-sub-001');
assert.equal(teamsLifecycle[0].expires_at_ms, startMs + 60_000);
assert.equal(teamsLifecycle[0].tenant_id, 'tenant-001');

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

const webexStart = normalizeWebexEvent({
  id: 'webex-hook-1',
  resource: 'meetings',
  event: 'started',
  data: {
    id: 'webex-meeting-001',
    meetingNumber: '123456789',
    title: 'Webex review',
    webLink: 'https://webex.example/meet/webex-meeting-001',
    hostEmail: 'host@example.com',
    startTime: startIso,
  },
});
assert.equal(webexStart[0].type, 'meeting_started');
assert.equal(webexStart[0].meeting.platform, 'webex');
assert.equal(webexStart[0].meeting.meeting_id, 'webex-meeting-001');
assert.equal(webexStart[0].meeting.external_meeting_id, '123456789');

const webexParticipant = normalizeWebexEvent({
  id: 'webex-hook-2',
  resource: 'meetingParticipants',
  event: 'joined',
  data: {
    meetingId: 'webex-meeting-001',
    id: 'participant-1',
    displayName: 'Ada',
    joinTime: startIso,
  },
});
assert.equal(webexParticipant[0].type, 'participant_joined');
assert.equal(webexParticipant[0].participant_id, 'participant-1');
assert.equal(webexParticipant[0].participant_name, 'Ada');

const webexTranscript = normalizeWebexEvent({
  id: 'webex-hook-3',
  resource: 'meetingTranscripts',
  event: 'created',
  data: {
    meetingId: 'webex-meeting-001',
    id: 'transcript-1',
    txtDownloadLink: 'https://webex.example/transcript-1.vtt',
  },
}, { receivedAtMs: startMs + 120_000 });
assert.equal(webexTranscript[0].type, 'artifact_ready');
assert.equal(webexTranscript[0].artifact_kind, 'transcript');
assert.equal(webexTranscript[0].artifact_url, 'https://webex.example/transcript-1.vtt');

console.log('ok meeting timeline SDK adapters');
