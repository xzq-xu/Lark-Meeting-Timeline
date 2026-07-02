import assert from 'node:assert/strict';

import {
  createReconciledPlatformEventIngestor,
} from '../packages/meeting-timeline-sdk/adapters/platform-ingest.mjs';
import {
  createMeetingSignalReconciler,
  meetingSignalFingerprint,
  reconcileMeetingSignals,
} from '../packages/meeting-timeline-sdk/adapters/signal-reconciler.mjs';

const startMs = 1_782_442_800_000;
const meetingUrl = 'https://meet.google.com/abc-defg-hij';

const localStart = {
  id: 'local-start-1',
  type: 'meeting_started',
  detected_platform: 'google_meet',
  meeting_id: 'local-google-001',
  meeting_url: meetingUrl,
  title: 'Local Google Meet',
  start_time_ms: startMs,
};

const googleStart = {
  id: 'google-start-1',
  type: 'google.workspace.meet.conference.v2.started',
  time: new Date(startMs).toISOString(),
  data: {
    conferenceRecord: { name: 'conferenceRecords/google-record-001' },
    meetingUri: meetingUrl,
    title: 'Official Google Meet',
  },
};

const zoomJoinUrl = 'https://us06web.zoom.us/j/987654321';
const zoomLocalStart = {
  id: 'zoom-local-start-1',
  type: 'meeting_started',
  detected_platform: 'zoom',
  meeting_id: '987654321',
  meeting_url: zoomJoinUrl,
  start_time_ms: startMs,
};
const zoomProviderStart = {
  event: 'meeting.started',
  event_id: 'zoom-start-1',
  event_ts: startMs + 2_000,
  payload: {
    object: {
      uuid: 'zoom-uuid-001',
      id: 987654321,
      topic: 'Zoom product review',
      start_time: new Date(startMs + 2_000).toISOString(),
    },
  },
};

const teamsJoinUrl = 'https://teams.microsoft.com/l/meetup-join/19%3Ameeting_TEAMS%40thread.v2/0';
const teamsLocalStart = {
  id: 'teams-local-start-1',
  type: 'meeting_started',
  url: teamsJoinUrl,
  start_time_ms: startMs,
};
const teamsProviderStart = {
  id: 'teams-start-1',
  resource: `/communications/onlineMeetings(joinWebUrl='${encodeURIComponent(teamsJoinUrl)}')/meetingCallEvents`,
  resourceData: {
    id: 'teams-call-event-1',
    eventType: 'callStarted',
    eventDateTime: new Date(startMs + 2_000).toISOString(),
  },
};

const fingerprint = meetingSignalFingerprint({
  type: 'meeting_started',
  meeting: {
    platform: 'google_meet',
    meeting_id: 'google-record-001',
  },
  occurred_at_ms: startMs,
  source_event_id: 'google-start-1',
  source: 'webhook',
});
assert.equal(fingerprint, 'event:webhook:google-start-1');

let reconciliation = reconcileMeetingSignals([
  {
    type: 'meeting_started',
    meeting: {
      platform: 'google_meet',
      meeting_id: 'local-google-001',
      meeting_url: meetingUrl,
    },
    occurred_at_ms: startMs,
    source_event_id: 'local-start-1',
    source: 'local_detector',
  },
  {
    type: 'meeting_started',
    meeting: {
      platform: 'google_meet',
      meeting_id: 'local-google-001',
      meeting_url: meetingUrl,
    },
    occurred_at_ms: startMs,
    source_event_id: 'local-start-1',
    source: 'local_detector',
  },
]);
assert.equal(reconciliation.signals.length, 1);
assert.equal(reconciliation.skipped[0].reason, 'duplicate_signal_fingerprint');

const reconciler = createMeetingSignalReconciler();
const localAxis = reconciler.reconcile({
  type: 'meeting_started',
  meeting: {
    platform: 'google_meet',
    meeting_id: 'local-google-001',
    meeting_url: meetingUrl,
  },
  occurred_at_ms: startMs,
  source_event_id: 'local-start-1',
  source: 'local_detector',
});
assert.equal(localAxis.signals.length, 1);

const providerAxis = reconciler.reconcile({
  type: 'meeting_started',
  meeting: {
    platform: 'google_meet',
    meeting_id: 'google-record-001',
    meeting_url: meetingUrl,
  },
  occurred_at_ms: startMs + 2_000,
  source_event_id: 'google-start-1',
  source: 'webhook',
});
assert.equal(providerAxis.signals.length, 1);
assert.equal(providerAxis.decisions[0].reason, 'provider_reconciles_local_axis');

const localAfterProvider = reconciler.reconcile({
  type: 'meeting_started',
  meeting: {
    platform: 'google_meet',
    meeting_id: 'local-google-002',
    meeting_url: meetingUrl,
  },
  occurred_at_ms: startMs + 3_000,
  source_event_id: 'local-start-2',
  source: 'local_detector',
});
assert.equal(localAfterProvider.signals.length, 0);
assert.equal(localAfterProvider.skipped[0].reason, 'provider_axis_already_active');

const zoomReconciler = createMeetingSignalReconciler();
assert.equal(zoomReconciler.reconcile({
  type: 'meeting_started',
  meeting: {
    platform: 'zoom',
    meeting_id: '987654321',
    meeting_url: zoomJoinUrl,
  },
  occurred_at_ms: startMs,
  source_event_id: 'zoom-local-start-1',
  source: 'local_detector',
}).signals.length, 1);
const zoomProviderAxis = zoomReconciler.reconcile({
  type: 'meeting_started',
  meeting: {
    platform: 'zoom',
    meeting_id: 'zoom-uuid-001',
    external_meeting_id: '987654321',
  },
  occurred_at_ms: startMs + 2_000,
  source_event_id: 'zoom-start-1',
  source: 'webhook',
});
assert.equal(zoomProviderAxis.signals.length, 1);
assert.equal(zoomProviderAxis.decisions[0].reason, 'provider_reconciles_local_axis');

const speakerStart = reconciler.reconcile({
  type: 'speaker_started',
  meeting: {
    platform: 'google_meet',
    meeting_id: 'google-record-001',
  },
  occurred_at_ms: startMs + 10_000,
  speaker_id: 'speaker-ada',
  speaker_name: 'Ada',
  source: 'local_detector',
});
assert.equal(speakerStart.signals.length, 1);
const speakerDuplicate = reconciler.reconcile({
  type: 'speaker_started',
  meeting: {
    platform: 'google_meet',
    meeting_id: 'google-record-001',
  },
  occurred_at_ms: startMs + 11_000,
  speaker_id: 'speaker-ada',
  speaker_name: 'Ada',
  source: 'local_detector',
});
assert.equal(speakerDuplicate.signals.length, 0);
assert.equal(speakerDuplicate.skipped[0].reason, 'duplicate_speaker_signal');

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

const ingestor = createReconciledPlatformEventIngestor(client);
const localIngest = await ingestor.ingest('local-detector', localStart);
assert.equal(localIngest.results.length, 1);
assert.equal(localIngest.results[0].action, 'startMeeting');
assert.equal(calls.at(-1).input.meeting_id, 'local-google-001');

const duplicateLocalIngest = await ingestor.ingest('local-detector', localStart);
assert.equal(duplicateLocalIngest.results.length, 0);
assert.equal(duplicateLocalIngest.reconciliation.skipped[0].reason, 'duplicate_signal_fingerprint');

const googleIngest = await ingestor.ingest('google-meet', googleStart);
assert.equal(googleIngest.rawSignals.length, 1);
assert.equal(googleIngest.results.length, 1);
assert.equal(googleIngest.reconciliation.decisions[0].reason, 'provider_reconciles_local_axis');
assert.equal(calls.at(-1).input.meeting_id, 'google-record-001');

const localAfterOfficial = await ingestor.ingest('local-detector', {
  ...localStart,
  id: 'local-start-2',
  meeting_id: 'local-google-002',
  start_time_ms: startMs + 4_000,
});
assert.equal(localAfterOfficial.results.length, 0);
assert.equal(localAfterOfficial.reconciliation.skipped[0].reason, 'provider_axis_already_active');

const zoomIngestor = createReconciledPlatformEventIngestor(client);
const zoomLocalIngest = await zoomIngestor.ingest('local-detector', zoomLocalStart);
assert.equal(zoomLocalIngest.results.length, 1);
assert.equal(calls.at(-1).input.meeting_id, '987654321');
const zoomProviderIngest = await zoomIngestor.ingest('zoom', zoomProviderStart);
assert.equal(zoomProviderIngest.results.length, 1);
assert.equal(zoomProviderIngest.reconciliation.decisions[0].reason, 'provider_reconciles_local_axis');
assert.equal(calls.at(-1).input.meeting_id, 'zoom-uuid-001');
assert.equal(calls.at(-1).input.external_meeting_id, '987654321');

const teamsIngestor = createReconciledPlatformEventIngestor(client);
const teamsLocalIngest = await teamsIngestor.ingest('local-detector', teamsLocalStart);
assert.equal(teamsLocalIngest.results.length, 1);
assert.equal(teamsLocalIngest.results[0].signal.meeting.platform, 'microsoft_teams');
const teamsProviderIngest = await teamsIngestor.ingest('teams', teamsProviderStart);
assert.equal(teamsProviderIngest.results.length, 1);
assert.equal(teamsProviderIngest.reconciliation.decisions[0].reason, 'provider_reconciles_local_axis');
assert.equal(teamsProviderIngest.results[0].signal.meeting.meeting_url, teamsJoinUrl);

console.log('ok meeting signal reconciler');
