#!/usr/bin/env node

import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformTimelineView,
  buildMeetingPlatformTimelineViewMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-timeline-view.mjs';
import {
  normalizeMeetingPlatform,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const inputFile = String(args.get('input') || args.get('view-input') || args.get('timeline-input') || '');
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const outDir = String(args.get('out-dir') || args.get('outDir') || '');
const writeViews = args.get('write-views') === 'true' || Boolean(outDir);
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

function inputForPlatform(input, platform) {
  if (!input) return {};
  const key = normalizePlatformKey(platform);
  const direct = input[key] ?? input[platform];
  if (direct) return direct;
  const meetingPlatform = normalizePlatformKey(input.meeting?.platform ?? input.platform ?? key);
  return meetingPlatform === key ? input : {};
}

function viewFile(platform) {
  return resolve(outDir || 'data/meeting-platform-timeline-views', `${platform}.json`);
}

function viewSummary(view = {}) {
  return {
    platform: view.platform,
    status: view.status,
    meeting_id: view.meeting?.meeting_id,
    marker_count: view.diagnostics?.marker_count ?? 0,
    visible_marker_count: view.diagnostics?.visible_marker_count ?? 0,
    uncalibrated_marker_count: view.diagnostics?.uncalibrated_marker_count ?? 0,
    warning_count: view.diagnostics?.warning_count ?? 0,
    viewport: view.viewport,
    rail_counts: view.diagnostics?.rail_counts ?? {},
    next_actions: view.next_actions ?? [],
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
  const matrix = buildMeetingPlatformTimelineViewMatrix({
    platforms: requiredPlatforms,
  });
  const views = matrix.platforms.map((platform) => buildMeetingPlatformTimelineView(platform, inputForPlatform(input, platform), {
    viewportStartMs: args.get('viewport-start-ms') ?? args.get('viewportStartMs') ?? undefined,
    viewportDurationMs: args.get('viewport-duration-ms') ?? args.get('viewportDurationMs') ?? undefined,
    minViewMs: args.get('min-view-ms') ?? args.get('minViewMs') ?? undefined,
    fullDurationMs: args.get('full-duration-ms') ?? args.get('fullDurationMs') ?? undefined,
  }));
  const writtenFiles = [];
  if (writeViews) {
    for (const view of views) {
      const file = viewFile(view.platform);
      await writeJson(file, view);
      writtenFiles.push(file);
    }
  }
  return {
    type: 'meeting_platform_timeline_view_report',
    ok: errors.length === 0,
    requirement: 'renderer_agnostic_multi_platform_timeline_view',
    input_file: inputFile ? resolve(inputFile) : undefined,
    platform_count: matrix.platform_count,
    marker_count: views.reduce((total, view) => total + view.diagnostics.marker_count, 0),
    visible_marker_count: views.reduce((total, view) => total + view.diagnostics.visible_marker_count, 0),
    uncalibrated_marker_count: views.reduce((total, view) => total + view.diagnostics.uncalibrated_marker_count, 0),
    warning_count: views.reduce((total, view) => total + view.diagnostics.warning_count, 0),
    rows: views.map((view) => viewSummary(view)),
    plans: matrix.rows,
    written_files: writtenFiles,
    views: args.get('include-views') === 'true' ? views : undefined,
    next_actions: unique(views.flatMap((view) => view.next_actions ?? [])),
    errors,
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_timeline_view_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | markers=${report.marker_count} | visible=${report.visible_marker_count} | uncalibrated=${report.uncalibrated_marker_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: status=${row.status} meeting=${row.meeting_id ?? 'none'} markers=${row.marker_count} visible=${row.visible_marker_count} warnings=${row.warning_count}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
    for (const error of report.errors) console.error(`error ${error.file}: ${error.error}`);
  }
  if (!report.ok) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
