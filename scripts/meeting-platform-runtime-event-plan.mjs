#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformRuntimeEventPlanMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-runtime-event.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const outDir = resolve(String(args.get('out-dir') || args.get('outDir') || 'data/meeting-platform-runtime-event-plans'));
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const includePlans = args.get('include-plans') === 'true';
const writePlans = args.get('write-plans') !== 'false';
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function planFile(platform) {
  return resolve(outDir, `${platform}.json`);
}

function rowWithFile(row = {}) {
  return {
    ...row,
    plan_file: planFile(row.platform),
  };
}

async function buildReport() {
  const matrix = buildMeetingPlatformRuntimeEventPlanMatrix({
    baseUrl,
    env: process.env,
    platforms: requiredPlatforms,
  });
  const writtenFiles = [];
  if (writePlans) {
    for (const plan of matrix.plans) {
      const file = planFile(plan.platform);
      await writeJson(file, plan);
      writtenFiles.push(file);
    }
  }
  return {
    type: 'meeting_platform_runtime_event_plan_report',
    ok: matrix.platform_count > 0
      && matrix.realtime_provider_dependency_count === 0
      && matrix.transcript_realtime_dependency_count === 0,
    requirement: 'runtime_event_plans_exported_for_nonblocking_realtime_annotation',
    base_url: baseUrl,
    out_dir: writePlans ? outDir : undefined,
    write_plans: writePlans,
    platform_count: matrix.platform_count,
    realtime_provider_dependency_count: matrix.realtime_provider_dependency_count,
    transcript_realtime_dependency_count: matrix.transcript_realtime_dependency_count,
    required_platforms: requiredPlatforms,
    written_files: writtenFiles,
    rows: matrix.rows.map((row) => rowWithFile(row)),
    matrix: includePlans ? matrix : {
      ...matrix,
      plans: undefined,
    },
    next_actions: matrix.next_actions ?? [],
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_runtime_event_plan_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | provider_blocks=${report.realtime_provider_dependency_count} | transcript_blocks=${report.transcript_realtime_dependency_count} | written=${report.written_files.length}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: action=${row.action} method=${row.client_method} producer=${row.producer} provider=${boolLabel(row.provider_dependency)} transcript=${boolLabel(row.transcript_dependency)} plan=${row.plan_file}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
  }
  if (!report.ok) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
