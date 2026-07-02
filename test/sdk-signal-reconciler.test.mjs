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

console.log('ok meeting signal reconciler');
