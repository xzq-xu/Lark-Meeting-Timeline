import assert from 'node:assert/strict';

import { normalizeGoogleMeetEvent } from '../packages/meeting-timeline-sdk/adapters/google-meet.mjs';
import { normalizeMicrosoftTeamsEvent } from '../packages/meeting-timeline-sdk/adapters/microsoft-teams.mjs';
import { normalizeZoomEvent } from '../packages/meeting-timeline-sdk/adapters/zoom.mjs';
import { normalizeWebexEvent } from '../packages/meeting-timeline-sdk/adapters/webex.mjs';
import { normalizeLarkEvent } from '../packages/meeting-timeline-sdk/adapters/lark.mjs';

const startMs = 1_782_442_800_000;
const startIso = new Date(startMs).toISOString();

const google = normalizeGoogleMeetEvent({
  id: 'google-direct-start',
  type: 'google.workspace.meet.conference.v2.started',
  time: startIso,
  data: {
    conferenceRecord: { name: 'conferenceRecords/google-direct-record' },
    title: 'Google direct import',
  },
});
assert.equal(google[0].type, 'meeting_started');
assert.equal(google[0].meeting.platform, 'google_meet');
assert.equal(google[0].meeting.meeting_id, 'google-direct-record');

const teams = normalizeMicrosoftTeamsEvent({
  id: 'teams-direct-start',
  resource: "communications/onlineMeetings(joinWebUrl='https%3A%2F%2Fteams.microsoft.com%2Fl%2Fmeetup-join%2Fdirect')/meetingCallEvents",
  resourceData: {
    id: 'teams-direct-call',
    eventType: 'callStarted',
    eventDateTime: startIso,
    subject: 'Teams direct import',
  },
});
assert.equal(teams[0].type, 'meeting_started');
assert.equal(teams[0].meeting.platform, 'microsoft_teams');
assert.equal(teams[0].meeting.meeting_id, 'teams-direct-call');

const zoom = normalizeZoomEvent({
  event: 'meeting.participant_joined',
  event_ts: startMs + 30_000,
  payload: {
    object: {
      uuid: 'zoom-direct-uuid',
      id: 987654321,
      participant: {
        user_id: 'zoom-user-1',
        user_name: 'Lin',
        join_time: new Date(startMs + 30_000).toISOString(),
      },
    },
  },
});
assert.equal(zoom[0].type, 'participant_joined');
assert.equal(zoom[0].meeting.platform, 'zoom');
assert.equal(zoom[0].participant_name, 'Lin');

const webex = normalizeWebexEvent({
  id: 'webex-direct-transcript',
  resource: 'meetingTranscripts',
  event: 'created',
  data: {
    meetingId: 'webex-direct-meeting',
    id: 'webex-transcript-1',
    txtDownloadLink: 'https://webex.example/transcript-1.vtt',
  },
});
assert.equal(webex[0].type, 'artifact_ready');
assert.equal(webex[0].meeting.platform, 'webex');
assert.equal(webex[0].artifact_kind, 'transcript');

const lark = normalizeLarkEvent({
  header: {
    event_id: 'lark-direct-start',
    event_type: 'vc.meeting.all_meeting_started_v1',
    create_time: String(startMs),
  },
  event: {
    meeting: {
      id: 'lark-direct-meeting',
      topic: 'Lark direct import',
      start_time: String(Math.round(startMs / 1000)),
    },
  },
});
assert.equal(lark[0].type, 'meeting_started');
assert.equal(lark[0].meeting.platform, 'lark');
assert.equal(lark[0].meeting.meeting_id, 'lark-direct-meeting');

console.log('ok provider normalizer direct imports');
