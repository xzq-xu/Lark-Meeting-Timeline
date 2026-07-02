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
  assert.equal(info.platform_events.platforms.some((item) => item.platform === 'google_meet'), true);

  const googleStart = await postJson(baseUrl, '/api/platform-events/google-meet', {
    id: 'google-start-001',
    type: 'google.workspace.meet.conference.v2.started',
    time: startIso,
    data: {
      conferenceRecord: { name: 'conferenceRecords/google-record-001' },
      meetingUri: 'https://meet.google.com/abc-defg-hij',
      title: 'Google platform event test',
    },
  });
  assert.equal(googleStart.ok, true);
  assert.equal(googleStart.signal_count, 1);
  assert.equal(googleStart.results[0].action, 'startMeeting');
  assert.equal(googleStart.state.meeting.source, 'google_meet_webhook');
  assert.equal(googleStart.state.meeting.platform, 'google_meet');
  assert.equal(googleStart.state.meeting.meeting_id, 'google-record-001');
  assert.equal(googleStart.state.events.some((event) => event.source === 'google_meet_webhook' && event.type === 'meeting_start'), true);

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

  const status = await getJson(baseUrl, '/api/platform-events/google-meet/status');
  assert.equal(status.status.platform, 'google_meet');
  assert.equal(status.status.received_count, 2);
  assert.equal(status.status.applied_signal_count, 2);
  assert.equal(status.status.last_signal_type, 'meeting_ended');

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

  const allStatus = await getJson(baseUrl, '/api/platform-events/status');
  assert.equal(allStatus.status.google_meet.received_count, 2);
  assert.equal(allStatus.status.microsoft_teams.received_count, 1);
  assert.equal(allStatus.status.zoom.received_count, 1);

  console.log('ok meeting platform event endpoints');
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
  if (process.exitCode) {
    console.error(output);
  }
}
