import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import net from 'node:net';

async function freePort() {
  const server = net.createServer();
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  const address = server.address();
  server.close();
  await once(server, 'close');
  return address.port;
}

async function waitForServer(baseUrl, timeoutMs = 10_000) {
  const deadline = Date.now() + timeoutMs;
  let lastError = null;
  while (Date.now() < deadline) {
    try {
      const response = await fetch(`${baseUrl}/api/state`);
      if (response.ok) return;
      lastError = new Error(`HTTP ${response.status}`);
    } catch (error) {
      lastError = error;
    }
    await new Promise((resolve) => setTimeout(resolve, 80));
  }
  throw lastError ?? new Error('server did not start');
}

async function postJson(baseUrl, path, body) {
  const response = await fetch(`${baseUrl}${path}`, {
    method: 'POST',
    headers: { 'content-type': 'application/json; charset=utf-8' },
    body: JSON.stringify(body),
  });
  const json = await response.json();
  assert.equal(response.ok, true, `${path} failed: ${JSON.stringify(json)}`);
  return json;
}

async function getJson(baseUrl, path) {
  const response = await fetch(`${baseUrl}${path}`);
  const json = await response.json();
  assert.equal(response.ok, true, `${path} failed: ${JSON.stringify(json)}`);
  return json;
}

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'meeting-platform-events-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    TIMELINE_DATA_DIR: tempDir,
  },
  stdio: ['ignore', 'pipe', 'pipe'],
});

let output = '';
child.stdout.on('data', (chunk) => {
  output += chunk.toString();
});
child.stderr.on('data', (chunk) => {
  output += chunk.toString();
});

try {
  await waitForServer(baseUrl);

  const startMs = 1_782_442_800_000;
  const startIso = new Date(startMs).toISOString();
  const endIso = new Date(startMs + 60_000).toISOString();

  const info = await getJson(baseUrl, '/api/annotation-ingest-info');
  assert.equal(info.platform_events.supported, true);
  assert.equal(typeof info.platform_events.setup_endpoint, 'string');
  assert.equal(info.platform_events.platforms.some((item) => item.platform === 'google_meet'), true);

  const setup = await getJson(baseUrl, '/api/platform-events/setup');
  assert.equal(setup.setup.length, 3);
  assert.equal(setup.setup.some((item) => item.platform === 'google_meet' && item.endpoint === `${baseUrl}/api/platform-events/google-meet`), true);
  assert.equal(setup.security_env_configured.GOOGLE_PUBSUB_OIDC_AUDIENCE, false);
  assert.equal(setup.readiness.google_meet.ready, false);
  assert.equal(setup.readiness.google_meet.checks.some((item) => item.id === 'google_pubsub_auth' && item.ok === false), true);

  const teamsSetup = await getJson(baseUrl, `/api/platform-events/teams/setup?join_web_url=${encodeURIComponent('https://teams.microsoft.com/l/meetup-join/example')}&client_state=test-state&subscription_id=graph-sub-1&subscription_expires_at=${encodeURIComponent('2026-06-26T02:30:00.000Z')}&now=${encodeURIComponent('2026-06-26T02:00:00.000Z')}`);
  assert.equal(teamsSetup.setup.platform, 'microsoft_teams');
  assert.equal(teamsSetup.readiness.platform, 'microsoft_teams');
  assert.equal(teamsSetup.readiness.ready, false);
  assert.equal(teamsSetup.maintenance.status, 'renewal_due');
  assert.equal(teamsSetup.maintenance.renewal_request.path, '/subscriptions/graph-sub-1');
  assert.equal(teamsSetup.setup.graph_subscription_request.notificationUrl, `${baseUrl}/api/platform-events/teams`);
  assert.equal(teamsSetup.setup.graph_subscription_request.lifecycleNotificationUrl, `${baseUrl}/api/platform-events/teams`);
  assert.equal(teamsSetup.setup.graph_subscription_request.clientState, 'test-state');
  assert.match(teamsSetup.setup.graph_subscription_request.resource, /meetingCallEvents$/);

  const googleStartEvent = {
    id: 'google-start-001',
    type: 'google.workspace.meet.conference.v2.started',
    time: startIso,
    data: {
      conferenceRecord: { name: 'conferenceRecords/google-record-001' },
      meetingUri: 'https://meet.google.com/abc-defg-hij',
      title: 'Google platform event test',
    },
  };
  const googleStart = await postJson(baseUrl, '/api/platform-events/google-meet', {
    message: {
      messageId: 'google-pubsub-message-001',
      publishTime: startIso,
      data: Buffer.from(JSON.stringify(googleStartEvent), 'utf8').toString('base64'),
      attributes: { eventType: googleStartEvent.type },
    },
    subscription: 'projects/demo-project/subscriptions/google-meet-events',
  });
  assert.equal(googleStart.ok, true);
  assert.equal(googleStart.signal_count, 1);
  assert.equal(googleStart.results[0].action, 'startMeeting');
  assert.equal(googleStart.state.meeting.source, 'google_meet_webhook');
  assert.equal(googleStart.state.meeting.platform, 'google_meet');
  assert.equal(googleStart.state.meeting.meeting_id, 'google-record-001');
  assert.equal(googleStart.state.events.some((event) => event.source === 'google_meet_webhook' && event.type === 'meeting_start'), true);

  const googleJoin = await postJson(baseUrl, '/api/platform-events/google-meet', {
    id: 'google-join-001',
    type: 'google.workspace.meet.participant.v2.joined',
    time: new Date(startMs + 15_000).toISOString(),
    data: {
      participantSession: {
        name: 'conferenceRecords/google-record-001/participants/user-123/participantSessions/session-1',
        participant: { displayName: 'Ada Lovelace' },
      },
    },
  });
  assert.equal(googleJoin.results[0].action, 'onParticipantSignal');
  assert.equal(googleJoin.results[0].response.skipped, false);
  const googleParticipantEvents = googleJoin.state.events.filter((event) => (
    event.source === 'google_meet_webhook'
      && event.type === 'participant_join'
      && event.metadata?.participant_id === 'user-123'
  ));
  assert.equal(googleParticipantEvents.length, 1);
  assert.equal(googleParticipantEvents[0].time_ms, 15_000);

  const duplicateGoogleJoin = await postJson(baseUrl, '/api/platform-events/google-meet', {
    id: 'google-join-duplicate',
    type: 'google.workspace.meet.participant.v2.joined',
    time: new Date(startMs + 16_000).toISOString(),
    data: {
      participantSession: {
        name: 'conferenceRecords/google-record-001/participants/user-123/participantSessions/session-2',
        participant: { displayName: 'Ada Lovelace' },
      },
    },
  });
  assert.equal(duplicateGoogleJoin.results[0].response.skipped, true);
  assert.equal(duplicateGoogleJoin.results[0].response.reason, 'duplicate_participant_signal_filtered');
  assert.equal(duplicateGoogleJoin.state.events.filter((event) => (
    event.source === 'google_meet_webhook'
      && event.type === 'participant_join'
      && event.metadata?.participant_id === 'user-123'
  )).length, 1);

  const annotation = await postJson(baseUrl, '/api/annotations', {
    id: 'google-platform-ann-1',
    source: 'hanwang_epaper',
    captured_at_ms: startMs + 12_000,
    kind: 'handwriting_trigger',
    label: 'why?',
  });
  assert.equal(annotation.ack.on_real_axis, true);
  assert.equal(annotation.ack.meeting_source, 'google_meet_webhook');
  assert.equal(annotation.ack.normalized_time_ms, 12_000);

  const googleEnd = await postJson(baseUrl, '/api/platform-events/google-meet', {
    id: 'google-end-001',
    type: 'google.workspace.meet.conference.v2.ended',
    time: endIso,
    data: {
      conferenceRecord: { name: 'conferenceRecords/google-record-001' },
    },
  });
  assert.equal(googleEnd.results[0].action, 'endMeeting');
  assert.equal(googleEnd.state.meeting.end_time, endIso);
  assert.equal(googleEnd.state.events.some((event) => event.source === 'google_meet_webhook' && event.type === 'meeting_end'), true);

  const googleTranscript = await postJson(baseUrl, '/api/platform-events/google-meet', {
    id: 'google-transcript-001',
    type: 'google.workspace.meet.transcript.v2.fileGenerated',
    time: new Date(startMs + 90_000).toISOString(),
    data: {
      transcript: {
        name: 'conferenceRecords/google-record-001/transcripts/transcript-1',
        docsDestination: { document: 'https://docs.google.com/document/d/transcript-1' },
      },
    },
  });
  assert.equal(googleTranscript.results[0].action, 'onArtifactSignal');
  assert.equal(googleTranscript.results[0].response.skipped, false);
  const googleTranscriptEvents = googleTranscript.state.events.filter((event) => (
    event.source === 'google_meet_webhook'
      && event.type === 'transcript_ready'
      && event.metadata?.artifact_id === 'transcript-1'
  ));
  assert.equal(googleTranscriptEvents.length, 1);
  assert.equal(googleTranscriptEvents[0].time_ms, 90_000);
  assert.equal(googleTranscriptEvents[0].metadata.artifact_url, 'https://docs.google.com/document/d/transcript-1');

  const duplicateGoogleTranscript = await postJson(baseUrl, '/api/platform-events/google-meet', {
    id: 'google-transcript-duplicate',
    type: 'google.workspace.meet.transcript.v2.fileGenerated',
    time: new Date(startMs + 91_000).toISOString(),
    data: {
      transcript: {
        name: 'conferenceRecords/google-record-001/transcripts/transcript-1',
        docsDestination: { document: 'https://docs.google.com/document/d/transcript-1' },
      },
    },
  });
  assert.equal(duplicateGoogleTranscript.results[0].response.skipped, true);
  assert.equal(duplicateGoogleTranscript.results[0].response.reason, 'duplicate_artifact_signal_filtered');
  assert.equal(duplicateGoogleTranscript.state.events.filter((event) => (
    event.source === 'google_meet_webhook'
      && event.type === 'transcript_ready'
      && event.metadata?.artifact_id === 'transcript-1'
  )).length, 1);

  const googleLifecycle = await postJson(baseUrl, '/api/platform-events/google-meet', {
    id: 'google-lifecycle-001',
    type: 'google.workspace.events.subscription.v1.expirationReminder',
    time: new Date(startMs + 92_000).toISOString(),
    data: {
      subscription: {
        name: 'subscriptions/google-sub-001',
        expire_time: new Date(startMs + 12 * 60 * 60 * 1000).toISOString(),
      },
    },
  });
  assert.equal(googleLifecycle.results[0].action, 'onSubscriptionLifecycleSignal');
  assert.equal(googleLifecycle.results[0].response.skipped, false);
  assert.equal(googleLifecycle.results[0].response.lifecycle.type, 'expiration_reminder');
  assert.equal(googleLifecycle.results[0].response.lifecycle.action, 'renew_subscription');
  assert.equal(googleLifecycle.status.lifecycle_event_count, 1);
  assert.equal(googleLifecycle.status.last_lifecycle.subscription_id, 'google-sub-001');
  assert.equal(googleLifecycle.state.events.filter((event) => event.source === 'google_meet_webhook').length, googleTranscriptEvents.length + googleParticipantEvents.length + 2);

  const status = await getJson(baseUrl, '/api/platform-events/google-meet/status');
  assert.equal(status.status.platform, 'google_meet');
  assert.equal(status.status.received_count, 7);
  assert.equal(status.status.applied_signal_count, 5);
  assert.equal(status.status.skipped_signal_count, 2);
  assert.equal(status.status.last_signal_type, 'subscription_lifecycle');
  assert.equal(status.status.lifecycle_event_count, 1);
  assert.equal(status.status.last_lifecycle.action, 'renew_subscription');

  const teamsStart = await postJson(baseUrl, '/api/platform-events/teams', {
    id: 'teams-start-001',
    resourceData: {
      eventType: 'callStarted',
      eventDateTime: startIso,
      onlineMeetingId: 'teams-meeting-001',
      joinWebUrl: 'https://teams.example/join',
      subject: 'Teams platform event test',
    },
  });
  assert.equal(teamsStart.state.meeting.source, 'microsoft_teams_webhook');
  assert.equal(teamsStart.state.meeting.platform, 'microsoft_teams');
  assert.equal(teamsStart.state.meeting.meeting_id, 'teams-meeting-001');

  const teamsRoster = await postJson(baseUrl, '/api/platform-events/teams', {
    id: 'teams-roster-001',
    resourceData: {
      eventType: 'rosterUpdated',
      eventDateTime: new Date(startMs + 20_000).toISOString(),
      onlineMeetingId: 'teams-meeting-001',
      'participants@delta': [
        { id: 'teams-user-1', displayName: 'Grace Hopper', joinDateTime: new Date(startMs + 20_000).toISOString() },
        { id: 'teams-user-2', displayName: 'Alan Turing', removedState: { reason: 'left' }, leaveDateTime: new Date(startMs + 30_000).toISOString() },
      ],
    },
  });
  assert.deepEqual(teamsRoster.results.map((item) => item.action), ['onParticipantSignal', 'onParticipantSignal']);
  assert.equal(teamsRoster.state.events.some((event) => (
    event.source === 'microsoft_teams_webhook'
      && event.type === 'participant_join'
      && event.metadata?.participant_id === 'teams-user-1'
  )), true);
  assert.equal(teamsRoster.state.events.some((event) => (
    event.source === 'microsoft_teams_webhook'
      && event.type === 'participant_leave'
      && event.metadata?.participant_id === 'teams-user-2'
  )), true);

  const teamsLifecycle = await postJson(baseUrl, '/api/platform-events/teams', {
    value: [
      {
        id: 'teams-lifecycle-001',
        lifecycleEvent: 'subscriptionRemoved',
        subscriptionId: 'teams-sub-001',
        subscriptionExpirationDateTime: new Date(startMs + 120_000).toISOString(),
        resource: 'communications/onlineMeetings(joinWebUrl=https%3A%2F%2Fteams.example%2Fjoin)/meetingCallEvents',
        tenantId: 'tenant-001',
        clientState: 'test-state',
      },
    ],
  });
  assert.equal(teamsLifecycle.results[0].action, 'onSubscriptionLifecycleSignal');
  assert.equal(teamsLifecycle.results[0].response.lifecycle.type, 'subscription_removed');
  assert.equal(teamsLifecycle.results[0].response.lifecycle.action, 'recreate_subscription');
  assert.equal(teamsLifecycle.status.lifecycle_event_count, 1);

  const zoomStart = await postJson(baseUrl, '/api/platform-events/zoom', {
    force: true,
    event: 'meeting.started',
    event_ts: startMs,
    payload: {
      object: {
        uuid: 'zoom-uuid-001',
        id: 987654321,
        topic: 'Zoom platform event test',
        join_url: 'https://zoom.us/j/987654321',
        start_time: startIso,
      },
    },
  });
  assert.equal(zoomStart.state.meeting.source, 'zoom_webhook');
  assert.equal(zoomStart.state.meeting.platform, 'zoom');
  assert.equal(zoomStart.state.meeting.meeting_id, 'zoom-uuid-001');

  const zoomRecording = await postJson(baseUrl, '/api/platform-events/zoom', {
    event: 'recording.completed',
    event_ts: startMs + 120_000,
    payload: {
      object: {
        uuid: 'zoom-uuid-001',
        id: 987654321,
        topic: 'Zoom platform event test',
        recording_files: [{ download_url: 'https://zoom.us/recording/download/1' }],
      },
    },
  });
  assert.equal(zoomRecording.results[0].action, 'onArtifactSignal');
  assert.equal(zoomRecording.results[0].response.skipped, false);
  assert.equal(zoomRecording.state.events.some((event) => (
    event.source === 'zoom_webhook'
      && event.type === 'recording_completed'
      && event.metadata?.artifact_url === 'https://zoom.us/recording/download/1'
  )), true);

  const allStatus = await getJson(baseUrl, '/api/platform-events/status');
  assert.equal(allStatus.status.google_meet.received_count, 7);
  assert.equal(allStatus.status.microsoft_teams.received_count, 3);
  assert.equal(allStatus.status.zoom.received_count, 2);

  console.log('ok meeting platform event endpoints');
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
  if (process.exitCode) {
    console.error(output);
  }
}
