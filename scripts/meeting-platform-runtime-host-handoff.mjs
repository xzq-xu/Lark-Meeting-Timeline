#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformRuntimeHostHandoffMatrix,
} from '../packages/meeting-timeline-sdk/adapters/meeting-platform-runtime-host.mjs';
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
const failOnNotReady = args.get('fail-on-not-ready') === 'true'
  || args.get('fail-on-incomplete') === 'true'
  || args.get('fail-on-unaccepted') === 'true';
const outDir = String(args.get('out-dir') || args.get('outDir') || '');
const writeHandoffs = args.get('write-handoffs') === 'true' || Boolean(outDir);
const platforms = unique(String(
  args.get('platforms')
    || args.get('platform')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function handoffFile(platform) {
  return resolve(outDir || 'data/meeting-platform-runtime-host-handoffs', `${platform}.json`);
}

function summarizeRow(row = {}) {
  return {
    platform: row.platform,
    display_name: row.display_name,
    accepted: row.accepted,
    host_ready: row.host_ready,
    browser_match_count: row.browser_match_count,
    host_hook_count: row.host_hook_count,
    timestamp_field: row.timestamp_field,
    provider_events_role: row.provider_events_role,
    transcript_role: row.transcript_role,
  };
}

async function buildReport() {
  const matrix = buildMeetingPlatformRuntimeHostHandoffMatrix({
    baseUrl,
    platforms,
  });
  const writtenFiles = [];
  if (writeHandoffs) {
    for (const handoff of matrix.handoffs) {
      const file = handoffFile(handoff.platform);
      await writeJson(file, handoff);
      writtenFiles.push(file);
    }
  }
  const ok = matrix.platform_count > 0
    && matrix.accepted_count === matrix.platform_count
    && matrix.provider_blocking_count === 0
    && matrix.transcript_blocking_count === 0;
  return {
    type: 'meeting_platform_runtime_host_handoff_report',
    ok,
    requirement: 'runtime_host_handoffs_ready_for_external_project_embedding',
    base_url: baseUrl,
    platforms,
    platform_count: matrix.platform_count,
    accepted_count: matrix.accepted_count,
    host_ready_count: matrix.host_ready_count,
    provider_blocking_count: matrix.provider_blocking_count,
    transcript_blocking_count: matrix.transcript_blocking_count,
    written_files: writtenFiles,
    rows: matrix.rows.map((row) => summarizeRow(row)),
    matrix: includeHandoffs ? matrix : undefined,
    next_actions: unique(matrix.next_actions ?? []),
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_runtime_host_handoff_report | ok=${boolLabel(report.ok)} | accepted=${report.accepted_count}/${report.platform_count} | host_ready=${report.host_ready_count}/${report.platform_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: accepted=${boolLabel(row.accepted)} host=${boolLabel(row.host_ready)} hooks=${row.host_hook_count} matches=${row.browser_match_count} timestamp=${row.timestamp_field}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
  }
  if (!report.ok && failOnNotReady) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
