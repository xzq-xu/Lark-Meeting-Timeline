import assert from 'node:assert/strict';

import {
  PLATFORM_ARTIFACT_IMPORTERS,
  buildArtifactImportPlan,
  buildArtifactImportPlans,
} from '../packages/meeting-timeline-sdk/adapters/artifact-plan.mjs';
import { normalizeGoogleMeetEvent } from '../packages/meeting-timeline-sdk/adapters/google-meet.mjs';
import { normalizeMicrosoftTeamsEvent } from '../packages/meeting-timeline-sdk/adapters/microsoft-teams.mjs';
import { normalizeWebexEvent } from '../packages/meeting-timeline-sdk/adapters/webex.mjs';
import { normalizeZoomEvent } from '../packages/meeting-timeline-sdk/adapters/zoom.mjs';
import { platformCapabilityContract } from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';

const startMs = 1_782_442_800_000;
const startIso = new Date(startMs).toISOString();

assert.equal(PLATFORM_ARTIFACT_IMPORTERS.google_meet.transcript.normalizer, 'normalizeGoogleMeetTranscriptEntries');
assert.equal(
  platformCapabilityContract('google-meet', { baseUrl: 'https://timeline.example.com' }).sdk_modules.artifact_plan,
  '@ai-annotation/meeting-timeline-sdk/adapters/artifact-plan',
);

const googleTranscriptSignal = normalizeGoogleMeetEvent({
  id: 'g-transcript-plan-1',
  type: 'google.workspace.meet.transcript.v2.fileGenerated',
  time: startIso,
  data: {
    transcript: {
      name: 'conferenceRecords/google-record-001/transcripts/transcript-1',
      docsDestination: { document: 'https://docs.google.com/document/d/transcript-1' },
    },
  },
})[0];
const googlePlan = buildArtifactImportPlan(googleTranscriptSignal);
assert.equal(googlePlan.status, 'requires_provider_fetch');
assert.equal(googlePlan.action, 'fetch_and_import_transcript');
assert.equal(googlePlan.platform, 'google_meet');
assert.equal(googlePlan.transcript_import.normalizer, 'normalizeGoogleMeetTranscriptEntries');
assert.equal(
  googlePlan.fetch.resource,
  'conferenceRecords/google-record-001/transcripts/transcript-1/entries',
);

const zoomRecordingSignal = normalizeZoomEvent({
  event: 'recording.completed',
  event_ts: startMs + 120_000,
  payload: {
    object: {
      uuid: 'zoom-uuid-001',
      id: 987654321,
      recording_files: [{ download_url: 'https://zoom.us/recording/download/1' }],
    },
  },
})[0];
const zoomPlan = buildArtifactImportPlan(zoomRecordingSignal);
assert.equal(zoomPlan.status, 'metadata_ready');
assert.equal(zoomPlan.action, 'store_recording_artifact');
assert.equal(zoomPlan.fetch.url, 'https://zoom.us/recording/download/1');
assert.equal(zoomPlan.transcript_import, undefined);

const teamsTranscriptSignal = normalizeMicrosoftTeamsEvent({
  id: 'teams-transcript-plan-1',
  resourceData: {
    eventType: 'callTranscriptCreated',
    eventDateTime: startIso,
    onlineMeetingId: 'teams-meeting-001',
    id: 'teams-transcript-001',
  },
})[0];
const teamsPlan = buildArtifactImportPlan(teamsTranscriptSignal, {
  importEndpoint: 'https://timeline.example.com/api/import/transcript',
});
assert.equal(teamsPlan.status, 'requires_provider_fetch');
assert.equal(teamsPlan.transcript_import.normalizer, 'normalizeMicrosoftTeamsTranscript');
assert.equal(teamsPlan.transcript_import.endpoint, 'https://timeline.example.com/api/import/transcript');
assert.match(teamsPlan.fetch.resource, /onlineMeetings\/teams-meeting-001\/artifacts\/teams-transcript-001/);

const webexTranscriptSignal = normalizeWebexEvent({
  id: 'webex-transcript-plan-1',
  resource: 'meetingTranscripts',
  event: 'created',
  data: {
    meetingId: 'webex-meeting-001',
    id: 'transcript-1',
    txtDownloadLink: 'https://webex.example/transcript-1.vtt',
  },
}, { receivedAtMs: startMs + 120_000 })[0];
const plans = buildArtifactImportPlans({ signals: [webexTranscriptSignal, { type: 'meeting_started' }] });
assert.equal(plans.length, 1);
assert.equal(plans[0].platform, 'webex');
assert.equal(plans[0].transcript_import.normalizer, 'normalizeWebexTranscript');
assert.equal(plans[0].fetch.url, 'https://webex.example/transcript-1.vtt');

const unsupportedPlan = buildArtifactImportPlan({
  type: 'artifact_ready',
  meeting: { platform: 'google_meet', meeting_id: 'google-record-001' },
  occurred_at_ms: startMs,
  artifact_kind: 'whiteboard',
});
assert.equal(unsupportedPlan.status, 'unsupported_artifact');
assert.equal(unsupportedPlan.issues[0].code, 'unsupported_artifact_kind');

const ignoredPlan = buildArtifactImportPlans([{ type: 'meeting_started' }], { includeIgnored: true });
assert.equal(ignoredPlan[0].status, 'ignored');

console.log('ok meeting artifact import plans');
