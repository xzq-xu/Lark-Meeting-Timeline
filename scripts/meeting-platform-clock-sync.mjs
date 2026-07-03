#!/usr/bin/env node

import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformClockSyncMatrix,
  buildMeetingPlatformClockSyncReport,
} from '../packages/meeting-timeline-sdk/adapters/platform-clock-sync.mjs';
import {
  normalizeMeetingPlatform,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const inputFile = String(args.get('input') || args.get('clock-input') || args.get('samples-file') || '');
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const outDir = String(args.get('out-dir') || args.get('outDir') || '');
const writeReports = args.get('write-reports') === 'true' || Boolean(outDir);
const includeReports = args.get('include-reports') === 'true';
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

function reportPath(platform) {
  return resolve(outDir || 'data/meeting-platform-clock-sync', `${platform}.json`);
}

function inputForPlatform(input, platform) {
  if (!input) return {};
  const key = normalizePlatformKey(platform);
  const direct = input[key] ?? input[platform];
  if (direct) return direct;
  const inputPlatform = normalizePlatformKey(input.platform ?? input.annotation?.platform ?? input.mark?.platform ?? key);
  return inputPlatform === key ? input : {};
}

function optionNumber(name, fallback) {
  const value = args.get(name);
  if (value == null || value === '') return fallback;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function optionBool(name, fallback) {
  const value = args.get(name);
  if (value == null) return fallback;
  return value !== 'false';
}

function clockOptions() {
  return {
    endpoint: args.get('endpoint') ?? args.get('time-sync-endpoint') ?? undefined,
    sampleCount: optionNumber('sample-count', 3),
    maxRecommendedSkewMs: optionNumber('max-recommended-skew-ms', 500),
    maxRecommendedRttMs: optionNumber('max-recommended-rtt-ms', 1000),
    maxUncertaintyMs: optionNumber('max-uncertainty-ms', 500),
    applyOffsetToAnnotation: optionBool('apply-offset-to-annotation', true),
  };
}

function rowSummary(report = {}) {
  return {
    platform: report.platform,
    status: report.status,
    accepted_for_realtime: report.accepted_for_realtime,
    sample_count: report.sample_count,
    usable_sample_count: report.usable_sample_count,
    recommended_offset_ms: report.recommended_offset_ms,
    selected_rtt_ms: report.selected_rtt_ms,
    selected_uncertainty_ms: report.selected_uncertainty_ms,
    warning_count: report.warnings?.length ?? 0,
    operator_attention_required: report.operator_attention_required,
    next_actions: report.next_actions ?? [],
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
  const options = clockOptions();
  const matrix = buildMeetingPlatformClockSyncMatrix({
    platforms: requiredPlatforms,
    ...options,
  });
  const reports = matrix.platforms.map((platform) => buildMeetingPlatformClockSyncReport(
    platform,
    inputForPlatform(input, platform),
    options,
  ));
  const writtenFiles = [];
  if (writeReports) {
    for (const report of reports) {
      const file = reportPath(report.platform);
      await writeJson(file, report);
      writtenFiles.push(file);
    }
  }
  return {
    type: 'meeting_platform_clock_sync_report_bundle',
    ok: errors.length === 0,
    requirement: 'calibrate_device_captured_at_ms_before_realtime_annotation_intake',
    input_file: inputFile ? resolve(inputFile) : undefined,
    platform_count: matrix.platform_count,
    accepted_count: reports.filter((report) => report.accepted_for_realtime).length,
    ready_count: reports.filter((report) => report.status === 'clock_sync_ready').length,
    requires_offset_count: reports.filter((report) => report.status === 'clock_sync_requires_offset').length,
    unstable_count: reports.filter((report) => report.status === 'clock_sync_unstable').length,
    missing_count: reports.filter((report) => report.status === 'clock_sync_missing').length,
    warning_count: reports.reduce((total, report) => total + (report.warnings?.length ?? 0), 0),
    rows: reports.map((report) => rowSummary(report)),
    plan_rows: matrix.rows,
    written_files: writtenFiles,
    reports: includeReports ? reports : undefined,
    next_actions: unique(reports.flatMap((report) => report.next_actions ?? [])),
    errors,
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_clock_sync_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | accepted=${report.accepted_count} | ready=${report.ready_count} | requires_offset=${report.requires_offset_count} | unstable=${report.unstable_count} | missing=${report.missing_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: status=${row.status} accepted=${boolLabel(row.accepted_for_realtime)} samples=${row.usable_sample_count}/${row.sample_count} offset=${row.recommended_offset_ms ?? 'missing'} rtt=${row.selected_rtt_ms ?? 'missing'} warnings=${row.warning_count}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
    for (const error of report.errors) console.error(`error ${error.file}: ${error.error}`);
  }
  if (!report.ok) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
