#!/usr/bin/env node

import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformRealtimeAnnotation,
  buildMeetingPlatformRealtimeAnnotationMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-realtime-annotation.mjs';
import {
  normalizeMeetingPlatform,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const inputFile = String(args.get('input') || args.get('annotation-input') || args.get('pipeline-input') || '');
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const outDir = String(args.get('out-dir') || args.get('outDir') || '');
const writePipelines = args.get('write-pipelines') === 'true' || Boolean(outDir);
const includePipelines = args.get('include-pipelines') === 'true';
const jsonOutput = args.get('json') === 'true';
const requiredPlatforms = unique(String(
  args.get('platforms') || args.get('required-platforms') || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

function normalizePlatformKey(platform) {
  try {
    return normalizeMeetingPlatform(platform);
  } catch {
    return String(platform);
  }
}

async function readJson(file) {
  if (!file) return undefined;
  return JSON.parse(await readFile(resolve(file), 'utf8'));
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function pipelineFile(platform) {
  return resolve(outDir || 'data/meeting-platform-realtime-annotation', `${platform}.json`);
}

function inputForPlatform(input, platform) {
  if (!input) return {};
  const key = normalizePlatformKey(platform);
  const direct = input[key] ?? input[platform];
  if (direct) return direct;
  const inputPlatform = normalizePlatformKey(
    input.current_meeting?.platform
      ?? input.currentMeeting?.platform
      ?? input.local_observer?.platform
      ?? input.localObserver?.platform
      ?? input.annotation?.platform
      ?? input.mark?.platform
      ?? input.platform
      ?? key,
  );
  return inputPlatform === key ? input : {};
}

function optionBool(name, fallback) {
  const value = args.get(name);
  if (value == null) return fallback;
  return value !== 'false';
}

function optionNumber(name, fallback) {
  const value = args.get(name);
  if (value == null || value === '') return fallback;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function pipelineOptions() {
  return {
    allowUnsyncedCapturedAtMs: optionBool('allow-unsynced-captured-at-ms', false),
    clock: {
      maxRecommendedSkewMs: optionNumber('max-recommended-skew-ms', 500),
      maxRecommendedRttMs: optionNumber('max-recommended-rtt-ms', 1000),
      maxUncertaintyMs: optionNumber('max-uncertainty-ms', 500),
    },
    binding: {
      acceptScore: optionNumber('accept-score', 70),
      conflictScore: optionNumber('conflict-score', 20),
      maxStartDeltaMs: optionNumber('max-start-delta-ms', 10 * 60_000),
    },
    intake: {
      allowOpenSessionWhenNoAxis: optionBool('allow-open-session-when-no-axis', true),
      allowPendingWhenNoAxis: optionBool('allow-pending-when-no-axis', true),
      allowLocalSimulationAxis: optionBool('allow-local-simulation-axis', false),
      allowedBeforeStartMs: optionNumber('allowed-before-start-ms', 0),
      openSessionSource: args.get('open-session-source') ?? undefined,
      defaultTitle: args.get('default-title') ?? undefined,
    },
  };
}

function rowSummary(pipeline = {}) {
  return {
    platform: pipeline.platform,
    status: pipeline.status,
    accepted_for_realtime: pipeline.accepted_for_realtime,
    actions: pipeline.actions ?? [],
    clock_status: pipeline.clock_sync?.status,
    binding_status: pipeline.session_binding?.status,
    intake_status: pipeline.annotation_intake?.status,
    selected_meeting_id: pipeline.selected_meeting?.meeting_id,
    captured_at_ms: pipeline.insert_payload?.captured_at_ms ?? pipeline.calibrated_annotation?.captured_at_ms,
    normalized_time_ms: pipeline.insert_payload?.time_ms,
    warning_count: pipeline.warnings?.length ?? 0,
    next_actions: pipeline.next_actions ?? [],
  };
}

async function buildReport() {
  const errors = [];
  let input;
  try {
    input = await readJson(inputFile);
  } catch (error) {
    errors.push({ file: resolve(inputFile), error: String(error?.message ?? error) });
  }
  const options = pipelineOptions();
  const matrix = buildMeetingPlatformRealtimeAnnotationMatrix({
    platforms: requiredPlatforms,
    ...options,
  });
  const pipelines = matrix.platforms.map((platform) => buildMeetingPlatformRealtimeAnnotation(
    platform,
    inputForPlatform(input, platform),
    options,
  ));
  const writtenFiles = [];
  if (writePipelines) {
    for (const pipeline of pipelines) {
      const file = pipelineFile(pipeline.platform);
      await writeJson(file, pipeline);
      writtenFiles.push(file);
    }
  }
  return {
    type: 'meeting_platform_realtime_annotation_report',
    ok: errors.length === 0,
    requirement: 'clock_sync_bind_meeting_identity_and_route_realtime_annotation',
    input_file: inputFile ? resolve(inputFile) : undefined,
    platform_count: matrix.platform_count,
    accepted_count: pipelines.filter((pipeline) => pipeline.accepted_for_realtime).length,
    ready_count: pipelines.filter((pipeline) => pipeline.status === 'ready_to_insert').length,
    start_axis_count: pipelines.filter((pipeline) => pipeline.status === 'start_axis_then_insert').length,
    start_open_session_count: pipelines.filter((pipeline) => pipeline.status === 'start_open_session_then_insert').length,
    pending_count: pipelines.filter((pipeline) => pipeline.status === 'pending_real_meeting').length,
    needs_clock_count: pipelines.filter((pipeline) => pipeline.status === 'needs_clock_sync').length,
    conflict_count: pipelines.filter((pipeline) => pipeline.status === 'binding_conflict').length,
    insufficient_count: pipelines.filter((pipeline) => pipeline.status === 'insufficient_identity').length,
    warning_count: pipelines.reduce((total, pipeline) => total + (pipeline.warnings?.length ?? 0), 0),
    provider_blocking_count: matrix.provider_blocking_count,
    transcript_blocking_count: matrix.transcript_blocking_count,
    rows: pipelines.map((pipeline) => rowSummary(pipeline)),
    plan_rows: matrix.rows,
    written_files: writtenFiles,
    pipelines: includePipelines ? pipelines : undefined,
    next_actions: unique(pipelines.flatMap((pipeline) => pipeline.next_actions ?? [])),
    errors,
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_realtime_annotation_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | accepted=${report.accepted_count} | ready=${report.ready_count} | start_axis=${report.start_axis_count} | open_session=${report.start_open_session_count} | pending=${report.pending_count} | needs_clock=${report.needs_clock_count} | conflicts=${report.conflict_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: status=${row.status} accepted=${boolLabel(row.accepted_for_realtime)} actions=${row.actions.join('+') || 'none'} clock=${row.clock_status} binding=${row.binding_status} intake=${row.intake_status} meeting=${row.selected_meeting_id ?? 'none'} captured=${row.captured_at_ms ?? 'missing'}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
    for (const error of report.errors) console.error(`error ${error.file}: ${error.error}`);
  }
  if (!report.ok) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
