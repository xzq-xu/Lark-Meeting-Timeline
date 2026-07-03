#!/usr/bin/env node

import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformArtifactHandoff,
  buildMeetingPlatformArtifactHandoffMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-artifact-handoff.mjs';
import {
  normalizeMeetingPlatform,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const includeHandoffs = args.get('include-handoffs') === 'true';
const outDir = String(args.get('out-dir') || args.get('outDir') || '');
const writeHandoffs = args.get('write-handoffs') === 'true' || Boolean(outDir);
const signalsFile = String(args.get('signals-file') || args.get('input') || '');
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
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

function handoffFile(platform) {
  return resolve(outDir || 'data/meeting-platform-artifact-handoffs', `${platform}.json`);
}

function signalArray(input) {
  if (!input) return [];
  if (Array.isArray(input)) return input;
  return input.signals ?? input.artifactSignals ?? input.artifact_signals ?? input.events ?? input.rawSignals ?? input.raw_signals ?? [];
}

function inputForPlatform(input, platform) {
  if (!input) return {};
  const key = normalizePlatformKey(platform);
  const direct = input[key] ?? input[platform];
  if (direct) return direct;
  return {
    signals: signalArray(input).filter((signal) => normalizePlatformKey(signal.platform ?? signal.meeting?.platform ?? key) === key),
  };
}

function rowSummary(handoff = {}) {
  return {
    platform: handoff.platform,
    status: handoff.status,
    input_signal_count: handoff.input_signal_count,
    import_plan_count: handoff.import_plan_count,
    fetch_request_count: handoff.fetch_request_count,
    transcript_import_count: handoff.transcript_import_count,
    recording_count: handoff.recording_count,
    warning_count: handoff.warning_count,
    post_meeting_only: handoff.post_meeting_only,
    transcript_blocks_realtime: handoff.transcript_blocks_realtime,
    next_actions: handoff.next_actions ?? [],
  };
}

async function buildReport() {
  const errors = [];
  let input;
  try {
    input = await readJson(signalsFile);
  } catch (error) {
    errors.push({ file: resolve(signalsFile), error: String(error?.message ?? error) });
  }
  const matrix = buildMeetingPlatformArtifactHandoffMatrix({
    baseUrl,
    platforms: requiredPlatforms,
  });
  const handoffs = matrix.platforms.map((platform) => buildMeetingPlatformArtifactHandoff(platform, inputForPlatform(input, platform), {
    baseUrl,
  }));
  const writtenFiles = [];
  if (writeHandoffs) {
    for (const handoff of handoffs) {
      const file = handoffFile(handoff.platform);
      await writeJson(file, handoff);
      writtenFiles.push(file);
    }
  }
  const fetchRequestCount = handoffs.reduce((total, handoff) => total + handoff.fetch_request_count, 0);
  const transcriptImportCount = handoffs.reduce((total, handoff) => total + handoff.transcript_import_count, 0);
  return {
    type: 'meeting_platform_artifact_handoff_report',
    ok: errors.length === 0,
    requirement: 'post_meeting_artifacts_do_not_block_realtime_annotations',
    base_url: baseUrl,
    signals_file: signalsFile ? resolve(signalsFile) : undefined,
    platform_count: matrix.platform_count,
    fetch_request_count: fetchRequestCount,
    transcript_import_count: transcriptImportCount,
    recording_count: handoffs.reduce((total, handoff) => total + handoff.recording_count, 0),
    warning_count: handoffs.reduce((total, handoff) => total + handoff.warning_count, 0),
    transcript_supported_count: matrix.transcript_supported_count,
    recording_supported_count: matrix.recording_supported_count,
    smart_notes_supported_count: matrix.smart_notes_supported_count,
    realtime_blocking_count: matrix.realtime_blocking_count,
    rows: handoffs.map((handoff) => rowSummary(handoff)),
    plan_rows: matrix.rows,
    written_files: writtenFiles,
    handoffs: includeHandoffs ? handoffs : undefined,
    next_actions: unique(handoffs.flatMap((handoff) => handoff.next_actions ?? [])),
    errors,
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_artifact_handoff_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | fetch_requests=${report.fetch_request_count} | transcript_imports=${report.transcript_import_count} | realtime_blocking=${report.realtime_blocking_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: status=${row.status} signals=${row.input_signal_count} imports=${row.import_plan_count} fetch_requests=${row.fetch_request_count} transcripts=${row.transcript_import_count} recordings=${row.recording_count} warnings=${row.warning_count}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
    for (const error of report.errors) console.error(`error ${error.file}: ${error.error}`);
  }
  if (!report.ok) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
