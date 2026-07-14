#!/usr/bin/env node

import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { mkdir, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import net from 'node:net';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  createMeetingTimelineAnnotationProducer,
  createMeetingTimelineClient,
} from '../packages/meeting-timeline-sdk/index.mjs';
import { normalizeMeetingPlatform } from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';
import { parseCliArgs, unique } from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const platforms = unique(String(args.get('platforms') || 'google-meet,teams,zoom,webex,lark')
  .split(',')
  .map((value) => value.trim())
  .filter(Boolean));
const runsPerPlatform = positiveInteger(args.get('runs') || args.get('runs-per-platform'), 3);
const annotationsPerRun = positiveInteger(args.get('annotations') || args.get('annotations-per-run'), 5);
const reportFile = resolve(repoRoot, String(args.get('report-file') || 'data/meeting-timeline-sdk-release-acceptance-report.json'));
const jsonOutput = args.get('json') === 'true';
const keepEvidence = args.get('keep-evidence') === 'true';
const failOnIncomplete = args.get('fail-on-incomplete') !== 'false';

const thresholds = Object.freeze({
  meeting_start_detection_latency_ms: 1_500,
  annotation_visible_p95_ms: 300,
  annotation_visible_hard_max_ms: 800,
  annotation_timeline_error_max_ms: 500,
  speaker_marker_visible_max_ms: 1_500,
  meeting_end_detection_latency_ms: 6_500,
  lost_annotation_count: 0,
  duplicate_visible_annotation_count: 0,
  previous_meeting_annotation_count_on_new_axis: 0,
});

function positiveInteger(value, fallback) {
  const number = Number(value);
  return Number.isInteger(number) && number > 0 ? number : fallback;
}

function percentile95(values = []) {
  const sorted = values.map(Number).filter(Number.isFinite).sort((left, right) => left - right);
  if (!sorted.length) return null;
  return sorted[Math.max(0, Math.ceil(sorted.length * 0.95) - 1)];
}

function maximum(values = []) {
  const rows = values.map(Number).filter(Number.isFinite);
  return rows.length ? Math.max(...rows) : null;
}

function absoluteMaximum(values = []) {
  return maximum(values.map((value) => Math.abs(Number(value))));
}

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
      // Retry while the isolated acceptance server starts.
    }
    await new Promise((resolveDelay) => setTimeout(resolveDelay, 50));
  }
  throw new Error('SDK release acceptance server did not start');
}

async function postJson(baseUrl, path, body) {
  const response = await fetch(`${baseUrl}${path}`, {
    method: 'POST',
    headers: { 'content-type': 'application/json; charset=utf-8' },
    body: JSON.stringify(body),
  });
  const value = await response.json();
  if (!response.ok) throw new Error(`${path} failed (${response.status}): ${JSON.stringify(value)}`);
  return value;
}

function snapshotRecord(platform, meetingId, phase, capturedAtMs) {
  const active = phase === 'active';
  return {
    schema: 'meeting_app_snapshot_record',
    schema_version: 1,
    id: `${platform}:${meetingId}:${phase}:${capturedAtMs}`,
    platform: normalizeMeetingPlatform(platform),
    phase,
    captured_at_ms: capturedAtMs,
    source: 'sdk_release_acceptance_harness',
    observer_surface: 'browser_extension',
    snapshot: {
      schema: 'meeting_app_dom_capture',
      schema_version: 1,
      platform: normalizeMeetingPlatform(platform),
      meeting_id: meetingId,
      title: active ? `SDK acceptance ${platform}` : `SDK acceptance ${platform} ended`,
      url: `https://acceptance.invalid/${platform}/${meetingId}`,
      source: 'sdk_release_acceptance_harness',
      page: active
        ? { controls: [{ label: 'Leave call' }], participants: [] }
        : { controls: [{ label: 'Join' }], participants: [] },
    },
  };
}

async function evidenceForMeeting(dataDir, meetingId) {
  const dir = join(dataDir, 'meeting-platform-field-evidence');
  const files = await readdir(dir);
  const file = files.find((name) => name.endsWith(`-${meetingId}.json`));
  if (!file) throw new Error(`missing SDK acceptance evidence for ${meetingId}`);
  return JSON.parse(await readFile(join(dir, file), 'utf8'));
}

function evaluateEvidence(platform, meetingId, evidence, options = {}) {
  const annotations = evidence.annotations ?? [];
  const annotationIds = new Set(annotations.map((row) => String(row.id)));
  const speakerStarts = (evidence.speaker_markers ?? []).filter((row) => row.kind === 'speaker_started');
  const speakers = new Set(speakerStarts.map((row) => String(row.speaker_id ?? row.speaker_name ?? row.label ?? '')).filter(Boolean));
  const latencies = annotations.map((row) => row.visible_latency_ms);
  const timelineErrors = annotations.map((row) => row.timeline_error_ms);
  const speakerLatencies = speakerStarts.map((row) => row.visible_latency_ms);
  const p95 = percentile95(latencies);
  const latencyMax = maximum(latencies);
  const timelineErrorMax = absoluteMaximum(timelineErrors);
  const speakerLatencyMax = maximum(speakerLatencies);
  const startLatency = Number(evidence.measurements?.start_detection_latency_ms);
  const endLatency = Number(evidence.measurements?.end_detection_latency_ms);
  const leakageCount = Number(evidence.measurements?.previous_meeting_annotation_count_on_new_axis);
  const lostCount = Math.max(0, annotationsPerRun - annotationIds.size);
  const duplicateVisibleCount = annotations.reduce((count, row) => (
    count + Math.max(0, Number(row.visible_instance_count ?? 1) - 1)
  ), 0);
  const retryRow = options.expectRetry
    ? annotations.find((row) => row.id === `${meetingId}:mark-1`)
    : null;
  const checks = [
    { id: 'source_is_device_neutral', passed: annotations.every((row) => row.source === 'sdk_annotation_producer') },
    { id: 'annotation_count', passed: annotationIds.size === annotationsPerRun },
    { id: 'annotation_visible_p95', passed: p95 != null && p95 <= thresholds.annotation_visible_p95_ms },
    { id: 'annotation_visible_hard_max', passed: latencyMax != null && latencyMax <= thresholds.annotation_visible_hard_max_ms },
    { id: 'annotation_timeline_error', passed: timelineErrorMax != null && timelineErrorMax <= thresholds.annotation_timeline_error_max_ms },
    { id: 'speaker_markers', passed: speakerStarts.length >= 2 && speakers.size >= 2 },
    { id: 'speaker_marker_visible_latency', passed: speakerLatencyMax != null && speakerLatencyMax <= thresholds.speaker_marker_visible_max_ms },
    { id: 'meeting_start_detection_latency', passed: Number.isFinite(startLatency) && startLatency <= thresholds.meeting_start_detection_latency_ms },
    { id: 'meeting_end_detection_latency', passed: Number.isFinite(endLatency) && endLatency <= thresholds.meeting_end_detection_latency_ms },
    { id: 'lost_annotations', passed: lostCount === thresholds.lost_annotation_count },
    { id: 'duplicate_visible_annotations', passed: duplicateVisibleCount === thresholds.duplicate_visible_annotation_count },
    { id: 'cross_meeting_isolation', passed: leakageCount === thresholds.previous_meeting_annotation_count_on_new_axis },
    { id: 'retry_keeps_stable_id', passed: !options.expectRetry || Number(retryRow?.delivery_attempt_count) >= 2 },
    { id: 'retry_does_not_duplicate_visible_mark', passed: !options.expectRetry || Number(retryRow?.visible_instance_count) === 1 },
  ];
  return {
    platform: normalizeMeetingPlatform(platform),
    meeting_id: meetingId,
    accepted: checks.every((check) => check.passed),
    annotation_count: annotationIds.size,
    speaker_count: speakers.size,
    annotation_visible_p95_ms: p95,
    annotation_visible_max_ms: latencyMax,
    annotation_timeline_error_max_ms: timelineErrorMax,
    speaker_marker_visible_max_ms: speakerLatencyMax,
    start_detection_latency_ms: startLatency,
    end_detection_latency_ms: endLatency,
    lost_annotation_count: lostCount,
    duplicate_visible_annotation_count: duplicateVisibleCount,
    previous_meeting_annotation_count_on_new_axis: leakageCount,
    checks,
    failed_check_ids: checks.filter((check) => !check.passed).map((check) => check.id),
  };
}

async function runMeeting(baseUrl, dataDir, platform, runNumber, options = {}) {
  const normalizedPlatform = normalizeMeetingPlatform(platform);
  const meetingId = `sdk-accept-${normalizedPlatform}-${runNumber}-${options.runToken}`;
  const client = createMeetingTimelineClient({
    baseUrl,
    source: 'sdk_annotation_producer',
    detectorSource: 'sdk_release_acceptance_harness',
  });
  const joinAtMs = Date.now();
  await postJson(baseUrl, '/api/meeting-platform/p0-reference', {
    platform: normalizedPlatform,
    meeting_id: meetingId,
    action: 'join',
    at_ms: joinAtMs,
  });
  const activeAtMs = Date.now();
  await client.startMeeting({
    platform: normalizedPlatform,
    meetingId,
    title: `SDK acceptance ${normalizedPlatform} ${runNumber}`,
    meetingUrl: `https://acceptance.invalid/${normalizedPlatform}/${meetingId}`,
    startTimeMs: activeAtMs,
    detectorSource: 'sdk_release_acceptance_harness',
    observerSurface: 'browser_extension',
    runId: `${normalizedPlatform}:${meetingId}`,
    suppressAutoAnnotations: true,
    meetingAppRecord: snapshotRecord(normalizedPlatform, meetingId, 'active', activeAtMs),
  });

  let simulateLostResponse = options.expectRetry;
  const producerClient = simulateLostResponse
    ? createMeetingTimelineClient({
      baseUrl,
      source: 'sdk_annotation_producer',
      fetch: async (url, init) => {
        const response = await fetch(url, init);
        if (simulateLostResponse && new URL(url).pathname === '/api/annotations') {
          simulateLostResponse = false;
          await response.arrayBuffer();
          throw new TypeError('simulated response loss after server commit');
        }
        return response;
      },
    })
    : client;
  const producer = createMeetingTimelineAnnotationProducer({
    client: producerClient,
    source: 'sdk_annotation_producer',
    producerId: `${normalizedPlatform}-${runNumber}`,
    retryDelayMs: 0,
  });

  for (let index = 1; index <= annotationsPerRun; index += 1) {
    await producer.publish({
      id: `${meetingId}:mark-${index}`,
      capturedAtMs: Date.now(),
      kind: 'annotation',
      label: `SDK mark ${index}`,
      payload: { producer: 'standard_sdk_reference_producer', ordinal: index },
    });
  }

  for (const [index, speakerId] of ['speaker-a', 'speaker-b'].entries()) {
    await client.insertMark({
      id: `${meetingId}:speaker-${index + 1}`,
      capturedAtMs: Date.now(),
      source: `${normalizedPlatform}_speaker_observer`,
      kind: 'speaker_started',
      intent: 'speaker_track',
      label: `${speakerId} speaking`,
      payload: { speaker_id: speakerId, stable_for_ms: 3_000 },
    });
  }

  const leaveAtMs = Date.now();
  await postJson(baseUrl, '/api/meeting-platform/p0-reference', {
    platform: normalizedPlatform,
    meeting_id: meetingId,
    action: 'leave',
    at_ms: leaveAtMs,
  });
  const endedAtMs = Date.now();
  await client.endMeeting({
    meetingId,
    endTimeMs: endedAtMs,
    detectorSource: 'sdk_release_acceptance_harness',
    observerSurface: 'browser_extension',
    runId: `${normalizedPlatform}:${meetingId}`,
    meetingAppRecord: snapshotRecord(normalizedPlatform, meetingId, 'ended', endedAtMs),
  });
  const evidence = await evidenceForMeeting(dataDir, meetingId);
  return evaluateEvidence(normalizedPlatform, meetingId, evidence, options);
}

const dataDir = await mkdtemp(join(tmpdir(), 'meeting-timeline-sdk-release-'));
const port = await freePort();
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: repoRoot,
  env: {
    ...process.env,
    PORT: String(port),
    TIMELINE_DATA_DIR: dataDir,
    LARK_WS_EVENTS: '0',
    REAL_DEMO_AUTO_ARM: '0',
    REAL_DEMO_AUTO_ANNOTATION: '0',
    REAL_DEMO_DEVICE_SIMULATOR: '0',
    REAL_DEMO_DEVICE_STREAM: '0',
  },
  stdio: ['ignore', 'pipe', 'pipe'],
});
let serverOutput = '';
child.stdout.on('data', (chunk) => { serverOutput += chunk.toString(); });
child.stderr.on('data', (chunk) => { serverOutput += chunk.toString(); });

let report;
try {
  await waitForServer(baseUrl);
  const runToken = `${Date.now()}`;
  const runs = [];
  for (const platform of platforms) {
    for (let runNumber = 1; runNumber <= runsPerPlatform; runNumber += 1) {
      runs.push(await runMeeting(baseUrl, dataDir, platform, runNumber, {
        runToken,
        expectRetry: platform === platforms[0] && runNumber === 1,
      }));
    }
  }
  const platformRows = platforms.map((platform) => {
    const normalized = normalizeMeetingPlatform(platform);
    const platformRuns = runs.filter((row) => row.platform === normalized);
    return {
      platform: normalized,
      accepted: platformRuns.length === runsPerPlatform && platformRuns.every((row) => row.accepted),
      accepted_run_count: platformRuns.filter((row) => row.accepted).length,
      required_run_count: runsPerPlatform,
      runs: platformRuns,
    };
  });
  report = {
    type: 'meeting_timeline_sdk_release_acceptance_report',
    schema: 'meeting_timeline_sdk_release_acceptance_report',
    schema_version: 1,
    generated_at: new Date().toISOString(),
    scope: 'deterministic_sdk_contract_and_transport',
    field_pilot_claimed: false,
    device_required: false,
    annotation_source: 'sdk_annotation_producer',
    requirements: {
      platform_count: platforms.length,
      runs_per_platform: runsPerPlatform,
      annotations_per_run: annotationsPerRun,
      stable_speakers_per_run: 2,
      thresholds,
    },
    accepted: platformRows.every((row) => row.accepted),
    accepted_platform_count: platformRows.filter((row) => row.accepted).length,
    platform_count: platformRows.length,
    accepted_run_count: runs.filter((row) => row.accepted).length,
    run_count: runs.length,
    platforms: platformRows,
  };
} catch (error) {
  report = {
    type: 'meeting_timeline_sdk_release_acceptance_report',
    schema: 'meeting_timeline_sdk_release_acceptance_report',
    schema_version: 1,
    generated_at: new Date().toISOString(),
    scope: 'deterministic_sdk_contract_and_transport',
    field_pilot_claimed: false,
    device_required: false,
    accepted: false,
    error: String(error?.stack ?? error),
    server_output: serverOutput,
  };
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
}

await mkdir(dirname(reportFile), { recursive: true });
await writeFile(reportFile, `${JSON.stringify(report, null, 2)}\n`, 'utf8');
if (!keepEvidence) await rm(dataDir, { recursive: true, force: true });

if (jsonOutput) {
  console.log(JSON.stringify(report, null, 2));
} else {
  console.log(`meeting_timeline_sdk_release_acceptance | accepted=${report.accepted ? 'yes' : 'no'} | platforms=${report.accepted_platform_count ?? 0}/${report.platform_count ?? platforms.length} | runs=${report.accepted_run_count ?? 0}/${report.run_count ?? platforms.length * runsPerPlatform}`);
  for (const platform of report.platforms ?? []) {
    const failures = platform.runs.flatMap((run) => run.failed_check_ids);
    console.log(`${platform.platform}: accepted=${platform.accepted ? 'yes' : 'no'} runs=${platform.accepted_run_count}/${platform.required_run_count} failed=${unique(failures).join(',') || '-'}`);
  }
  if (report.error) console.error(report.error);
}

if (!report.accepted && failOnIncomplete) process.exitCode = 2;
