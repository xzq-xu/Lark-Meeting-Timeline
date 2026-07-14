import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { once } from 'node:events';
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

function flatLarkMeetingEvent(eventType, eventId, meeting) {
  const isEnd = eventType.includes('ended');
  const createTime = isEnd ? meeting.end_time : meeting.start_time;
  return {
    schema: '2.0',
    event_id: eventId,
    event_type: eventType,
    create_time: String(createTime),
    meeting,
  };
}

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'lark-restore-ended-window-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    LARK_APP_ID: 'restore-ended-window-app',
    LARK_APP_SECRET: 'restore-ended-window-secret',
    LARK_VERIFICATION_TOKEN: '',
    LARK_EVENT_CALLBACK_URL: 'https://restore-ended-window.example.com/api/lark/events',
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

  const firstStartSeconds = 1_782_797_754;
  const firstEndSeconds = firstStartSeconds + 40;
  const secondStartSeconds = firstStartSeconds + 181;

  const firstMeeting = {
    id: 'restore-ended-first',
    meeting_no: '694819136',
    topic: 'First meeting should be recovered with end',
    start_time: String(firstStartSeconds),
  };
  const secondMeeting = {
    id: 'restore-ended-second',
    meeting_no: '101107575',
    topic: 'Second meeting has no end yet',
    start_time: String(secondStartSeconds),
  };

  const firstStart = await postJson(baseUrl, '/api/lark/events', flatLarkMeetingEvent(
    'vc.meeting.all_meeting_started_v1',
    'restore-ended-first-start',
    firstMeeting,
  ));
  assert.equal(firstStart.timeline_started, true);
  assert.equal(firstStart.state.meeting.meeting_id, 'restore-ended-first');

  const secondStart = await postJson(baseUrl, '/api/lark/events', flatLarkMeetingEvent(
    'vc.meeting.all_meeting_started_v1',
    'restore-ended-second-start',
    secondMeeting,
  ));
  assert.equal(secondStart.timeline_started, true);
  assert.equal(secondStart.state.meeting.meeting_id, 'restore-ended-second');

  const firstEnd = await postJson(baseUrl, '/api/lark/events', flatLarkMeetingEvent(
    'vc.meeting.all_meeting_ended_v1',
    'restore-ended-first-end',
    {
      ...firstMeeting,
      end_time: String(firstEndSeconds),
    },
  ));
  assert.equal(firstEnd.ok, false);
  assert.equal(firstEnd.ignored_reason, 'meeting_end_without_known_start_axis');

  const restored = await postJson(baseUrl, '/api/lark/restore-latest-real-meeting-axis', {});
  assert.equal(restored.restored, true);
  assert.equal(restored.restore_mode, 'latest_ended_meeting_window');
  assert.equal(restored.start_result.ok, true);
  assert.equal(restored.end_result.ok, true);
  assert.equal(restored.state.meeting.meeting_id, 'restore-ended-first');
  assert.equal(restored.state.meeting.external_meeting_id, '694819136');
  assert.equal(restored.state.meeting.start_time, '2026-06-30T05:35:54.000Z');
  assert.equal(restored.state.meeting.end_time, '2026-06-30T05:36:34.000Z');
  assert.equal(restored.state.events.some((event) => event.id === 'restore-ended-first-start'), true);
  assert.equal(restored.state.events.some((event) => event.id === 'restore-ended-first-end'), true);

  const activeStartSeconds = secondStartSeconds + 420;
  const activeMeeting = {
    id: 'restore-ended-active',
    meeting_no: '695824794',
    topic: 'Latest active meeting should win restore',
    start_time: String(activeStartSeconds),
  };
  const activeStart = await postJson(baseUrl, '/api/lark/events', flatLarkMeetingEvent(
    'vc.meeting.all_meeting_started_v1',
    'restore-ended-active-start',
    activeMeeting,
  ));
  assert.equal(activeStart.timeline_started, true);
  assert.equal(activeStart.state.meeting.meeting_id, 'restore-ended-active');

  const restoredActive = await postJson(baseUrl, '/api/lark/restore-latest-real-meeting-axis', {});
  assert.equal(restoredActive.restored, true);
  assert.equal(restoredActive.restore_mode, 'latest_start_event');
  assert.equal(restoredActive.start_result.ok, true);
  assert.equal(restoredActive.end_result, null);
  assert.equal(restoredActive.state.meeting.meeting_id, 'restore-ended-active');
  assert.equal(restoredActive.state.meeting.external_meeting_id, '695824794');
  assert.equal(restoredActive.state.meeting.end_time, null);
  assert.equal(restoredActive.state.events.some((event) => event.id === 'restore-ended-active-start'), true);

  await postJson(baseUrl, '/api/demo/reset', {});
  const probeWithRecentRestore = await postJson(baseUrl, '/api/lark/real-meeting-probe/start', {
    timeout_ms: 120_000,
    reset_temporary_axis: true,
    auto_search: true,
  });
  assert.equal(probeWithRecentRestore.recent_event_restore.restored, true);
  assert.equal(probeWithRecentRestore.recent_event_restore.reason, 'recent_start_event_before_probe');
  assert.equal(probeWithRecentRestore.recent_event_restore.state.meeting.meeting_id, 'restore-ended-active');
  assert.equal(probeWithRecentRestore.recent_event_restore.state.events.some((event) => event.id === 'restore-ended-active-start'), true);

  console.log('ok restore latest ended meeting window');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
