#!/usr/bin/env node

import { mkdir, readFile, readdir, writeFile } from 'node:fs/promises';
import { dirname, extname, join, resolve } from 'node:path';
import {
  runMeetingPlatformRuntimeHostReplayMatrix,
} from '../packages/meeting-timeline-sdk/adapters/meeting-platform-runtime-host-verifier.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const includeReports = args.get('include-reports') === 'true';
const failOnFailure = args.get('fail-on-failure') === 'true'
  || args.get('fail-on-not-ready') === 'true'
  || args.get('fail-on-incomplete') === 'true';
const inputFiles = unique(String(
  args.get('input')
    || args.get('inputs')
    || '',
).split(',').map((item) => item.trim()).filter(Boolean));
const inputDir = String(args.get('dir') || args.get('input-dir') || '');
const platforms = unique(String(
  args.get('platforms')
    || args.get('platform')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

async function jsonFilesFromDir(dir) {
  if (!dir) return [];
  const absolute = resolve(dir);
  let entries = [];
  try {
    entries = await readdir(absolute, { withFileTypes: true });
  } catch (error) {
    if (error?.code === 'ENOENT') return [];
    throw error;
  }
  return entries
    .filter((entry) => entry.isFile() && extname(entry.name).toLowerCase() === '.json')
    .map((entry) => join(absolute, entry.name))
    .sort();
}

async function readJson(file) {
  return JSON.parse(await readFile(resolve(file), 'utf8'));
}

async function loadInputs() {
  const files = unique([
    ...inputFiles,
    ...(await jsonFilesFromDir(inputDir)),
  ]);
  const inputs = [];
  for (const file of files) {
    inputs.push({
      file: resolve(file),
      data: await readJson(file),
    });
  }
  return inputs;
}

function summarizeRow(row = {}) {
  return {
    platform: row.platform,
    accepted: row.accepted,
    input_record_count: row.input_record_count,
    replay_row_count: row.replay_row_count,
    active_at_ms: row.active_at_ms,
    end_at_ms: row.end_at_ms,
    call_count: row.call_count,
    actions: row.actions,
    signal_types: row.signal_types,
    missing: row.missing,
  };
}

async function buildReport() {
  const loaded = await loadInputs();
  const input = loaded.length === 1
    ? loaded[0].data
    : {
      items: loaded.map((item) => item.data),
    };
  const matrix = await runMeetingPlatformRuntimeHostReplayMatrix({
    platforms,
    input,
  });
  const ok = matrix.platform_count > 0 && matrix.accepted_count === matrix.platform_count;
  return {
    type: 'meeting_platform_runtime_host_replay_report',
    ok,
    requirement: 'runtime_host_can_replay_live_meeting_app_snapshots_into_realtime_axis_events',
    input_file_count: loaded.length,
    input_files: loaded.map((item) => item.file),
    platform_count: matrix.platform_count,
    accepted_count: matrix.accepted_count,
    platforms: matrix.platforms,
    rows: matrix.rows.map((row) => summarizeRow(row)),
    matrix: includeReports ? matrix : undefined,
    next_actions: matrix.next_actions,
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_runtime_host_replay_report | ok=${boolLabel(report.ok)} | accepted=${report.accepted_count}/${report.platform_count} | inputs=${report.input_file_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: accepted=${boolLabel(row.accepted)} records=${row.input_record_count} rows=${row.replay_row_count} calls=${row.call_count} actions=${row.actions.join(',') || '-'} signals=${row.signal_types.join(',') || '-'} missing=${row.missing.join(',') || '-'}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
  }
  if (!report.ok && failOnFailure) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
