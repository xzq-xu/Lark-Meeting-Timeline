#!/usr/bin/env node

import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingAppAdapterCapabilityMatrix,
  buildMeetingAppAdapterExecutionPlanMatrix,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-adapter-capability.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const includeReports = args.get('include-reports') === 'true';
const failOnRejected = args.get('fail-on-rejected') === 'true' || args.get('fail-on-failed') === 'true';
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()).filter(Boolean));

async function readJsonArg(name) {
  const value = args.get(name);
  if (!value) return undefined;
  return JSON.parse(await readFile(resolve(String(value)), 'utf8'));
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

const inputs = await readJsonArg('input-file') ?? await readJsonArg('snapshot-file');
const evidenceByAdapter = await readJsonArg('evidence-file');
const matrix = buildMeetingAppAdapterCapabilityMatrix({
  platforms: requiredPlatforms,
  inputs,
  evidenceByAdapter,
  includeVerification: evidenceByAdapter != null || args.get('include-verification') === 'true',
});
const executionPlanMatrix = buildMeetingAppAdapterExecutionPlanMatrix({
  capabilityMatrix: matrix,
});

const report = {
  type: 'meeting_app_adapter_capability_cli_report',
  ok: matrix.accepted_count === matrix.platform_count,
  platform_count: matrix.platform_count,
  accepted_count: matrix.accepted_count,
  static_ready_count: matrix.static_ready_count,
  pilot_ready_count: matrix.pilot_ready_count,
  production_ready_count: matrix.production_ready_count,
  local_axis_ready_count: matrix.local_axis_ready_count,
  local_speaker_ready_count: matrix.local_speaker_ready_count,
  provider_axis_declared_count: matrix.provider_axis_declared_count,
  required_platforms: requiredPlatforms,
  rows: matrix.rows,
  execution_plan_rows: executionPlanMatrix.rows,
  matrix: includeReports ? matrix : {
    ...matrix,
    reports: undefined,
  },
  execution_plan_matrix: includeReports ? executionPlanMatrix : {
    ...executionPlanMatrix,
    plans: undefined,
  },
  next_actions: matrix.next_actions,
};

if (reportFile) await writeJson(resolve(reportFile), report);

if (jsonOutput) {
  console.log(JSON.stringify(report, null, 2));
} else {
  console.log(`meeting_app_adapter_capability_report | ok=${boolLabel(report.ok)} | accepted=${report.accepted_count}/${report.platform_count} | pilot=${report.pilot_ready_count} | production=${report.production_ready_count} | local_axis=${report.local_axis_ready_count} | provider_axis=${report.provider_axis_declared_count}`);
  for (const row of report.rows) {
    const planRow = report.execution_plan_rows.find((item) => item.platform === row.platform);
    console.log(`${row.platform}: mode=${row.recommended_mode} pilot=${boolLabel(row.pilot_ready)} production=${boolLabel(row.production_ready)} realtime=${row.realtime_axis_status} speaker=${row.speaker_track_status} blocked=${planRow?.first_blocked_step ?? 'none'} next=${row.first_next_action ?? 'none'}`);
  }
  if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
}

if (!report.ok && failOnRejected) process.exitCode = 2;
