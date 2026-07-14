#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdapterAuthoringMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-authoring.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const outDir = resolve(String(args.get('out-dir') || args.get('outDir') || 'data/meeting-platform-adapter-authoring'));
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const includePlans = args.get('include-plans') === 'true';
const writePlans = args.get('write-plans') !== 'false';
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()).filter(Boolean));

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function planFile(platform) {
  return resolve(outDir, `${platform}.json`);
}

async function buildReport() {
  const matrix = buildMeetingPlatformAdapterAuthoringMatrix({
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
    type: 'meeting_platform_adapter_authoring_report',
    ok: matrix.platform_count > 0,
    requirement: 'adapter_authoring_plans_available_for_builtin_and_external_meeting_platforms',
    base_url: baseUrl,
    out_dir: writePlans ? outDir : undefined,
    write_plans: writePlans,
    platform_count: matrix.platform_count,
    built_in_count: matrix.built_in_count,
    external_authoring_count: matrix.external_authoring_count,
    browser_surface_ready_count: matrix.browser_surface_ready_count,
    required_platforms: requiredPlatforms,
    written_files: writtenFiles,
    rows: matrix.rows.map((row) => ({
      ...row,
      plan_file: planFile(row.platform),
    })),
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
    console.log(`meeting_platform_adapter_authoring_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | builtin=${report.built_in_count} | external=${report.external_authoring_count} | browser_ready=${report.browser_surface_ready_count} | written=${report.written_files.length}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: builtin=${boolLabel(row.built_in)} surface=${row.recommended_first_surface} provider=${row.provider_path} matches=${row.browser_match_count} next=${row.next_action}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
  }
  if (!report.ok) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
