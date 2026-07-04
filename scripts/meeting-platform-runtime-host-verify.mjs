#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  runMeetingPlatformRuntimeHostVerificationMatrix,
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
const startMs = args.get('start-ms') || args.get('startMs');
const endOffsetMs = args.get('end-offset-ms') || args.get('endOffsetMs');
const platforms = unique(String(
  args.get('platforms')
    || args.get('platform')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function summarizeRow(row = {}) {
  return {
    platform: row.platform,
    accepted: row.accepted,
    call_count: row.call_count,
    actions: row.actions,
    signal_types: row.signal_types,
    missing: row.missing,
    start_capture_profile: row.start_capture_profile,
    end_capture_profile: row.end_capture_profile,
    participant_count: row.participant_count,
  };
}

async function buildReport() {
  const matrix = await runMeetingPlatformRuntimeHostVerificationMatrix({
    platforms,
    startMs,
    endOffsetMs,
  });
  const ok = matrix.platform_count > 0 && matrix.accepted_count === matrix.platform_count;
  return {
    type: 'meeting_platform_runtime_host_verify_report',
    ok,
    requirement: 'runtime_host_can_drive_realtime_meeting_axis_from_browser_runtime_for_each_platform',
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
    console.log(`meeting_platform_runtime_host_verify_report | ok=${boolLabel(report.ok)} | accepted=${report.accepted_count}/${report.platform_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: accepted=${boolLabel(row.accepted)} calls=${row.call_count} actions=${row.actions.join(',')} signals=${row.signal_types.join(',')} missing=${row.missing.join(',') || '-'}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
  }
  if (!report.ok && failOnFailure) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
