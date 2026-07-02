import assert from 'node:assert/strict';

import {
  buildPlatformTranscriptImportPayload,
  normalizeGoogleMeetTranscriptEntries,
  normalizeMicrosoftTeamsTranscript,
  normalizeZoomTranscript,
  parseTimedTextTranscript,
} from '../packages/meeting-timeline-sdk/adapters/transcript.mjs';

const startMs = 1_782_442_800_000;
const startIso = new Date(startMs).toISOString();

const googleEntries = normalizeGoogleMeetTranscriptEntries({
  transcriptEntries: [
    {
      name: 'conferenceRecords/google-record-001/transcripts/transcript-1/entries/entry-1',
      participant: 'conferenceRecords/google-record-001/participants/user-123',
      startTime: new Date(startMs + 1000).toISOString(),
      endTime: new Date(startMs + 4000).toISOString(),
      text: 'Google transcript line',
      languageCode: 'en-US',
    },
  ],
});
assert.equal(googleEntries.length, 1);
assert.equal(googleEntries[0].id, 'entry-1');
assert.equal(googleEntries[0].speaker_id, 'user-123');
assert.equal(googleEntries[0].source, 'google_meet_transcript');
assert.equal(googleEntries[0].text, 'Google transcript line');

const timedText = `WEBVTT

00:00:01.000 --> 00:00:04.500
<v Ada>why does this matter?

00:00:05.000 --> 00:00:07.000
Grace: follow up later
`;
const parsedTimedText = parseTimedTextTranscript(timedText, { source: 'test_vtt', language: 'en-US' });
assert.equal(parsedTimedText.length, 2);
assert.equal(parsedTimedText[0].start_ms, 1000);
assert.equal(parsedTimedText[0].end_ms, 4500);
assert.equal(parsedTimedText[0].speaker_name, 'Ada');
assert.equal(parsedTimedText[0].text, 'why does this matter?');
assert.equal(parsedTimedText[1].speaker_name, 'Grace');

const teamsSegments = normalizeMicrosoftTeamsTranscript({
  content: timedText,
}, { language: 'en-US' });
assert.equal(teamsSegments[0].source, 'microsoft_teams_transcript');
assert.equal(teamsSegments[0].speaker_name, 'Ada');

const zoomSegments = normalizeZoomTranscript({
  vtt: timedText,
});
assert.equal(zoomSegments[0].source, 'zoom_transcript_vtt');
assert.equal(zoomSegments[1].text, 'follow up later');

const googlePayload = buildPlatformTranscriptImportPayload({
  platform: 'google_meet',
  meeting: {
    platform: 'google_meet',
    meetingId: 'google-record-001',
    startTimeMs: startMs,
    title: 'Google transcript import',
  },
  raw: {
    transcriptEntries: [
      {
        name: 'conferenceRecords/google-record-001/transcripts/transcript-1/entries/entry-2',
        participant: 'conferenceRecords/google-record-001/participants/user-456',
        startTime: new Date(startMs + 8000).toISOString(),
        endTime: new Date(startMs + 10_000).toISOString(),
        text: 'Second line',
      },
    ],
  },
});
assert.equal(googlePayload.meeting.platform, 'google_meet');
assert.equal(googlePayload.meeting.meeting_id, 'google-record-001');
assert.equal(googlePayload.meeting.start_time, startIso);
assert.equal(googlePayload.transcript[0].source, 'google_meet_transcript');
assert.equal(googlePayload.transcript[0].speaker_id, 'user-456');

const teamsPayload = buildPlatformTranscriptImportPayload({
  platform: 'microsoft_teams',
  meeting: {
    platform: 'microsoft_teams',
    meetingId: 'teams-meeting-001',
    startTimeMs: startMs,
  },
  raw: timedText,
});
assert.equal(teamsPayload.transcript.length, 2);
assert.equal(teamsPayload.transcript[0].source, 'microsoft_teams_transcript');

console.log('ok meeting transcript adapters');
