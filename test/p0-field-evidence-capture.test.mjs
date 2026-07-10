import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { once } from 'node:events';
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
  while (Date.now() < deadline) {
    try {
      const response = await fetch(`${baseUrl}/api/state`);
      if (response.ok) return;
    } catch {
      // Retry while the test server starts.
    }
    await new Promise((resolve) => setTimeout(resolve, 80));
  }
  throw new Error('server did not start');
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

function snapshotRecord(phase, capturedAtMs, meetingId = 'p0-capture-test') {
  return {
    schema: 'meeting_app_snapshot_record',
    schema_version: 1,
    id: `google_meet:${meetingId}:${phase}:${capturedAtMs}`,
    platform: 'google_meet',
    phase,
    captured_at_ms: capturedAtMs,
    source: 'meeting_app_extension_auto_evidence',
    observer_surface: 'browser_extension',
    snapshot: {
      schema: 'meeting_app_dom_capture',
      schema_version: 1,
      platform: 'google_meet',
      meeting_id: meetingId,
      title: phase === 'active' ? 'P0 real capture test' : 'Google Meet ended',
      url: 'https://meet.google.com/abc-defg-hij',
      source: 'meeting_app_extension_auto_evidence',
      page: phase === 'active'
        ? { controls: [{ label: 'Leave call' }], participants: [{ id: 'speaker-a', speaking: true }] }
        : { controls: [{ label: 'Join now' }], participants: [] },
    },
  };
}

const port = await freePort();
const tempDir = await mkdtemp(join(tmpdir(), 'meeting-platform-p0-capture-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    LARK_WS_EVENTS: '0',
    REAL_DEMO_AUTO_ARM: '0',
    TIMELINE_DATA_DIR: tempDir,
  },
  stdio: ['ignore', 'pipe', 'pipe'],
});

let output = '';
child.stdout.on('data', (chunk) => { output += chunk.toString(); });
child.stderr.on('data', (chunk) => { output += chunk.toString(); });

try {
  await waitForServer(baseUrl);
  const candidates = await postJson(baseUrl, '/api/meeting-platform/runtime-events', {
    schema: 'meeting_platform_runtime_event',
    action: 'observe_platform_candidates',
    tabs: [{ active: true, url: 'https://meet.google.com/abc-defg-hij', title: 'P0 real capture test' }],
  });
  assert.equal(candidates.candidate_count, 1);
  assert.equal(candidates.selected_candidate.platform, 'google_meet');
  assert.equal(candidates.mutates_timeline, false);

  const joinAtMs = Date.now();
  const joined = await postJson(baseUrl, '/api/meeting-platform/p0-reference', {
    platform: 'google_meet',
    meeting_id: 'p0-capture-test',
    action: 'join',
    at_ms: joinAtMs,
  });
  assert.equal(joined.queued, true);

  const activeCapturedAtMs = joinAtMs + 20;
  const started = await postJson(baseUrl, '/api/meeting-session/start', {
    platform: 'google_meet',
    meeting_id: 'p0-capture-test',
    title: 'P0 real capture test',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    start_time_ms: activeCapturedAtMs,
    detector_source: 'meeting_app_extension',
    suppress_auto_annotations: true,
    meeting_app_record: snapshotRecord('active', activeCapturedAtMs),
  });
  assert.equal(started.session_evidence.recorded, true);

  const marks = Array.from({ length: 5 }, (_, index) => ({
    id: `p0-device-mark-${index}`,
    source: 'hanwang_epaper',
    captured_at_ms: Date.now(),
    kind: 'handwriting_trigger',
    intent: 'question',
    label: `why ${index}`,
    realtime: true,
    strokes: [],
  }));
  const inserted = await postJson(baseUrl, '/api/annotations/batch', { annotations: marks });
  assert.equal(inserted.annotation_evidence.every((row) => row.recorded), true);

  for (const [index, speaker] of ['speaker-a', 'speaker-b'].entries()) {
    await postJson(baseUrl, '/api/annotations', {
      id: `p0-speaker-${index}`,
      source: 'google_meet_speaker',
      captured_at_ms: Date.now(),
      kind: 'speaker_started',
      intent: 'speaker_track',
      label: `${speaker} speaking`,
      speaker_id: speaker,
    });
  }

  const leaveAtMs = Date.now();
  const left = await postJson(baseUrl, '/api/meeting-platform/p0-reference', {
    platform: 'google_meet',
    meeting_id: 'p0-capture-test',
    action: 'leave',
    at_ms: leaveAtMs,
  });
  assert.equal(left.updated, true);

  const ended = await postJson(baseUrl, '/api/meeting-session/end', {
    meeting_id: 'p0-capture-test',
    end_time_ms: leaveAtMs + 20,
    detector_source: 'meeting_app_extension',
    meeting_app_record: snapshotRecord('ended', leaveAtMs + 20),
  });
  assert.equal(ended.session_evidence.recorded, true);

  const evidencePath = join(tempDir, 'meeting-platform-field-evidence', 'google_meet-p0-capture-test.json');
  const evidence = JSON.parse(await readFile(evidencePath, 'utf8'));
  assert.equal(evidence.schema, 'meeting_platform_field_evidence_input');
  assert.equal(evidence.run_id, 'google_meet:p0-capture-test');
  assert.equal(evidence.observer_surface, 'browser_extension');
  assert.equal(evidence.measurements.observer_surface, 'browser_extension');
  assert.equal(evidence.meetingAppRecords.length, 2);
  assert.equal(evidence.annotations.length, 5);
  assert.equal(evidence.speaker_markers.length, 2);
  assert.equal(evidence.measurements.previous_meeting_annotation_count_on_new_axis, 0);
  assert.equal(evidence.measurements.start_detection_latency_ms >= 0, true);
  assert.equal(evidence.measurements.end_detection_latency_ms >= 0, true);
  assert.equal(evidence.measurements.axis_started_at_ms >= joinAtMs, true);
  assert.equal(evidence.measurements.axis_ended_at_ms >= leaveAtMs, true);
  assert.equal(evidence.annotations.every((row) => Math.abs(row.timeline_error_ms) <= 1), true);

  const secondStartMs = Date.now();
  const secondStarted = await postJson(baseUrl, '/api/meeting-session/start', {
    platform: 'google_meet',
    meeting_id: 'p0-capture-test-2',
    title: 'P0 isolation test',
    meeting_url: 'https://meet.google.com/klm-nopq-rst',
    start_time_ms: secondStartMs,
    detector_source: 'meeting_app_extension',
    suppress_auto_annotations: true,
    meeting_app_record: snapshotRecord('active', secondStartMs, 'p0-capture-test-2'),
  });
  assert.equal(secondStarted.session_evidence.recorded, true);

  const secondEvidencePath = join(tempDir, 'meeting-platform-field-evidence', 'google_meet-p0-capture-test-2.json');
  const secondEvidence = JSON.parse(await readFile(secondEvidencePath, 'utf8'));
  assert.equal(secondEvidence.annotations.length, 0);
  assert.equal(secondEvidence.speaker_markers.length, 0);
  assert.equal(secondEvidence.measurements.previous_meeting_annotation_count_on_new_axis, 0);

  const firstEvidenceAfterSecondStart = JSON.parse(await readFile(evidencePath, 'utf8'));
  assert.equal(firstEvidenceAfterSecondStart.annotations.length, 5);

  console.log('ok p0 field evidence capture');
} catch (error) {
  console.error(output);
  throw error;
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(tempDir, { recursive: true, force: true });
}
