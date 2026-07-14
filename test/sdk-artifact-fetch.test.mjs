import assert from 'node:assert/strict';

import {
  buildArtifactFetchRequest,
  fetchAndImportArtifactTranscript,
  fetchArtifactContent,
} from '../packages/meeting-timeline-sdk/adapters/artifact-fetch.mjs';
import { buildArtifactImportPlan } from '../packages/meeting-timeline-sdk/adapters/artifact-plan.mjs';
import { normalizeGoogleMeetEvent } from '../packages/meeting-timeline-sdk/adapters/google-meet.mjs';
import { normalizeMicrosoftTeamsEvent } from '../packages/meeting-timeline-sdk/adapters/microsoft-teams.mjs';
import { normalizeWebexEvent } from '../packages/meeting-timeline-sdk/adapters/webex.mjs';
import { normalizeZoomEvent } from '../packages/meeting-timeline-sdk/adapters/zoom.mjs';
import { platformCapabilityContract } from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';

const startMs = 1_782_442_800_000;
const startIso = new Date(startMs).toISOString();

function response(body, options = {}) {
  return {
    ok: options.ok ?? true,
    status: options.status ?? 200,
    headers: new Headers(options.headers ?? { 'content-type': options.contentType ?? 'application/json' }),
    async json() {
      return body;
    },
    async text() {
      return typeof body === 'string' ? body : JSON.stringify(body);
    },
    async arrayBuffer() {
      return new TextEncoder().encode(typeof body === 'string' ? body : JSON.stringify(body)).buffer;
    },
  };
}

const googleTranscriptSignal = normalizeGoogleMeetEvent({
  id: 'google-fetch-transcript-1',
  type: 'google.workspace.meet.transcript.v2.fileGenerated',
  time: startIso,
  data: {
    transcript: {
      name: 'conferenceRecords/google-fetch-001/transcripts/transcript-1',
      docsDestination: { document: 'https://docs.google.com/document/d/transcript-1' },
    },
  },
})[0];
const googlePlan = buildArtifactImportPlan(googleTranscriptSignal);
const googleRequest = buildArtifactFetchRequest(googlePlan, {
  accessToken: 'google-token',
  pageSize: 100,
});
assert.equal(googleRequest.method, 'GET');
assert.equal(
  googleRequest.url,
  'https://meet.googleapis.com/v2/conferenceRecords/google-fetch-001/transcripts/transcript-1/entries?pageSize=100',
);
assert.equal(googleRequest.headers.authorization, 'Bearer google-token');
assert.equal(googleRequest.response_type, 'json');
assert.equal(
  platformCapabilityContract('google-meet', { baseUrl: 'https://timeline.example.com' }).sdk_modules.artifact_fetch,
  '@ai-annotation/meeting-timeline-sdk/adapters/artifact-fetch',
);

const googleFetched = await fetchArtifactContent(googlePlan, {
  accessToken: 'google-token',
  pageSize: 100,
  fetchImpl: async (url, init) => {
    assert.equal(url, googleRequest.url);
    assert.equal(init.headers.authorization, 'Bearer google-token');
    return response({
      transcriptEntries: [{
        name: 'conferenceRecords/google-fetch-001/transcripts/transcript-1/entries/entry-1',
        startTime: new Date(startMs + 1000).toISOString(),
        endTime: new Date(startMs + 4000).toISOString(),
        text: 'Google fetched transcript line',
      }],
    });
  },
});
assert.equal(googleFetched.platform, 'google_meet');
assert.equal(googleFetched.body.transcriptEntries[0].text, 'Google fetched transcript line');

const teamsSignal = normalizeMicrosoftTeamsEvent({
  id: 'teams-fetch-transcript-1',
  resourceData: {
    eventType: 'callTranscriptCreated',
    eventDateTime: startIso,
    onlineMeetingId: 'teams-meeting-001',
    id: 'teams-transcript-001',
  },
})[0];
const teamsPlan = buildArtifactImportPlan(teamsSignal);
const teamsRequest = buildArtifactFetchRequest(teamsPlan, {
  userId: 'user-1',
  accessToken: 'graph-token',
});
assert.equal(
  teamsRequest.url,
  'https://graph.microsoft.com/v1.0/users/user-1/onlineMeetings/teams-meeting-001/transcripts/teams-transcript-001/content',
);
assert.equal(teamsRequest.headers.accept, 'text/vtt');
assert.equal(teamsRequest.headers.authorization, 'Bearer graph-token');

const vtt = `WEBVTT

00:00:01.000 --> 00:00:04.000
<v Ada>Teams fetched transcript line
`;
const importCalls = [];
const client = {
  async importTranscript(input, options) {
    importCalls.push({ input, options });
    return { ok: true, imported: input.transcript.length };
  },
};
const importedTeams = await fetchAndImportArtifactTranscript(client, teamsPlan, {
  userId: 'user-1',
  accessToken: 'graph-token',
  fetchImpl: async (url, init) => {
    assert.equal(url, teamsRequest.url);
    assert.equal(init.headers.accept, 'text/vtt');
    return response(vtt, { contentType: 'text/vtt' });
  },
});
assert.equal(importedTeams.import_result.segment_count, 1);
assert.equal(importCalls.at(-1).input.transcript[0].speaker_name, 'Ada');
assert.equal(importCalls.at(-1).input.transcript[0].text, 'Teams fetched transcript line');

const zoomSignal = normalizeZoomEvent({
  event: 'recording.completed',
  event_ts: startMs,
  payload: {
    object: {
      uuid: 'zoom-fetch-001',
      id: 987654321,
      recording_files: [{ file_type: 'TRANSCRIPT', download_url: 'https://zoom.us/rec/transcript.vtt' }],
    },
  },
})[0];
const zoomPlan = buildArtifactImportPlan({ ...zoomSignal, artifact_kind: 'transcript' });
const zoomRequest = buildArtifactFetchRequest(zoomPlan, { accessToken: 'zoom-token' });
assert.equal(zoomRequest.url, 'https://zoom.us/rec/transcript.vtt');
assert.equal(zoomRequest.headers.authorization, 'Bearer zoom-token');
assert.equal(zoomRequest.response_type, 'text');

const webexSignal = normalizeWebexEvent({
  id: 'webex-fetch-transcript-1',
  resource: 'meetingTranscripts',
  event: 'created',
  data: {
    meetingId: 'webex-meeting-001',
    id: 'webex-transcript-001',
    txtDownloadLink: 'https://webex.example/transcript.txt',
  },
}, { receivedAtMs: startMs })[0];
const webexPlan = buildArtifactImportPlan(webexSignal);
const webexFetched = await fetchArtifactContent(webexPlan, {
  accessToken: 'webex-token',
  fetchImpl: async (url, init) => {
    assert.equal(url, 'https://webex.example/transcript.txt');
    assert.equal(init.headers.authorization, 'Bearer webex-token');
    return response(vtt, { contentType: 'text/vtt' });
  },
});
assert.equal(webexFetched.body.includes('WEBVTT'), true);

assert.throws(
  () => buildArtifactFetchRequest({ status: 'unsupported_artifact' }),
  /not available/,
);

await assert.rejects(
  () => fetchArtifactContent(googlePlan, {
    fetchImpl: async () => response({ error: 'forbidden' }, { ok: false, status: 403 }),
  }),
  /HTTP 403/,
);

console.log('ok meeting artifact fetch helpers');
