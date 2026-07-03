#!/usr/bin/env node

import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAnnotationIntake,
  buildMeetingPlatformAnnotationIntakeMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-annotation-intake.mjs';
import {
  normalizeMeetingPlatform,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const inputFile = String(args.get('input') || args.get('annotation-input') || args.get('intake-input') || '');
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const outDir = String(args.get('out-dir') || args.get('outDir') || '');
const writeDecisions = args.get('write-decisions') === 'true' || Boolean(outDir);
const includeDecisions = args.get('include-decisions') === 'true';
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

function decisionFile(platform) {
  return resolve(outDir || 'data/meeting-platform-annotation-intake', `${platform}.json`);
}

function inputForPlatform(input, platform) {
  if (!input) return {};
  const key = normalizePlatformKey(platform);
  const direct = input[key] ?? input[platform];
  if (direct) return direct;
  const meetingPlatform = normalizePlatformKey(
    input.current_meeting?.platform
      ?? input.currentMeeting?.platform
      ?? input.meeting?.platform
      ?? input.annotation?.platform
      ?? input.mark?.platform
      ?? input.platform
      ?? key,
  );
  return meetingPlatform === key ? input : {};
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

function intakeOptions() {
  return {
    allowOpenSessionWhenNoAxis: optionBool('allow-open-session-when-no-axis', true),
    allowPendingWhenNoAxis: optionBool('allow-pending-when-no-axis', true),
    allowLocalSimulationAxis: optionBool('allow-local-simulation-axis', false),
    allowedBeforeStartMs: optionNumber('allowed-before-start-ms', 0),
    openSessionSource: args.get('open-session-source') ?? undefined,
    defaultTitle: args.get('default-title') ?? undefined,
  };
}

function rowSummary(decision = {}) {
  return {
    platform: decision.platform,
    status: decision.status,
    action: decision.action,
    accepted_for_realtime: decision.accepted_for_realtime,
    timing_reliable: decision.timing_reliable,
    axis_state: decision.axis_state,
    captured_at_ms: decision.captured_at_ms,
    captured_at_source: decision.captured_at_source,
    normalized_time_ms: decision.normalized_time_ms,
    warning_count: decision.warnings?.length ?? 0,
    next_actions: decision.next_actions ?? [],
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
  const options = intakeOptions();
  const matrix = buildMeetingPlatformAnnotationIntakeMatrix({
    platforms: requiredPlatforms,
    ...options,
  });
  const decisions = matrix.platforms.map((platform) => buildMeetingPlatformAnnotationIntake(
    platform,
    inputForPlatform(input, platform),
    options,
  ));
  const writtenFiles = [];
  if (writeDecisions) {
    for (const decision of decisions) {
      const file = decisionFile(decision.platform);
      await writeJson(file, decision);
      writtenFiles.push(file);
    }
  }
  return {
    type: 'meeting_platform_annotation_intake_report',
    ok: errors.length === 0,
    requirement: 'route_realtime_annotations_by_captured_at_ms_without_provider_or_transcript_blocking',
    input_file: inputFile ? resolve(inputFile) : undefined,
    platform_count: matrix.platform_count,
    accepted_count: decisions.filter((decision) => decision.accepted_for_realtime).length,
    ready_count: decisions.filter((decision) => decision.status === 'ready_to_insert_current_axis').length,
    open_session_count: decisions.filter((decision) => decision.status === 'start_open_session_then_insert').length,
    pending_count: decisions.filter((decision) => decision.status === 'pending_real_meeting').length,
    missing_captured_count: decisions.filter((decision) => decision.status === 'needs_device_captured_at').length,
    after_end_count: decisions.filter((decision) => decision.status === 'after_meeting_end').length,
    warning_count: decisions.reduce((total, decision) => total + (decision.warnings?.length ?? 0), 0),
    rows: decisions.map((decision) => rowSummary(decision)),
    plan_rows: matrix.rows,
    written_files: writtenFiles,
    decisions: includeDecisions ? decisions : undefined,
    next_actions: unique(decisions.flatMap((decision) => decision.next_actions ?? [])),
    errors,
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_annotation_intake_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | accepted=${report.accepted_count} | ready=${report.ready_count} | open_session=${report.open_session_count} | pending=${report.pending_count} | missing_time=${report.missing_captured_count} | after_end=${report.after_end_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: status=${row.status} action=${row.action} axis=${row.axis_state} captured=${row.captured_at_ms ?? 'missing'} normalized=${row.normalized_time_ms ?? 'none'} warnings=${row.warning_count}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
    for (const error of report.errors) console.error(`error ${error.file}: ${error.error}`);
  }
  if (!report.ok) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
