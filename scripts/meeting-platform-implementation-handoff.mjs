#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformImplementationHandoffMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-implementation-handoff.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const outDir = resolve(String(args.get('out-dir') || args.get('outDir') || 'data/meeting-platform-implementation-handoffs'));
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const includeHandoffs = args.get('include-handoffs') === 'true';
const writeHandoffs = args.get('write-handoffs') !== 'false';
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()).filter(Boolean));

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function handoffFile(platform) {
  return resolve(outDir, `${platform}.json`);
}

function rowWithFile(row = {}) {
  return {
    ...row,
    handoff_file: handoffFile(row.platform),
  };
}

async function buildReport() {
  const matrix = buildMeetingPlatformImplementationHandoffMatrix({
    baseUrl,
    env: process.env,
    platforms: requiredPlatforms,
  });
  const writtenFiles = [];
  if (writeHandoffs) {
    for (const handoff of matrix.handoffs) {
      const file = handoffFile(handoff.platform);
      await writeJson(file, handoff);
      writtenFiles.push(file);
    }
  }
  return {
    type: 'meeting_platform_implementation_handoff_report',
    ok: matrix.platform_count > 0 && matrix.implementation_ready_count === matrix.platform_count,
    requirement: 'external_projects_can_wire_meeting_annotation_timeline_sdk_per_platform',
    base_url: baseUrl,
    out_dir: writeHandoffs ? outDir : undefined,
    write_handoffs: writeHandoffs,
    platform_count: matrix.platform_count,
    implementation_ready_count: matrix.implementation_ready_count,
    pilot_ready_count: matrix.pilot_ready_count,
    production_ready_count: matrix.production_ready_count,
    recommended_first_platform: matrix.recommended_first_platform,
    recommended_first_surface: matrix.recommended_first_surface,
    required_platforms: requiredPlatforms,
    written_files: writtenFiles,
    rows: matrix.rows.map((row) => rowWithFile(row)),
    matrix: includeHandoffs ? matrix : {
      ...matrix,
      handoffs: undefined,
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
    console.log(`meeting_platform_implementation_handoff_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | implementation_ready=${report.implementation_ready_count} | pilot=${report.pilot_ready_count} | production=${report.production_ready_count} | first=${report.recommended_first_platform ?? 'n/a'} | surface=${report.recommended_first_surface ?? 'n/a'} | written=${report.written_files.length}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: ready=${boolLabel(row.implementation_ready)} surface=${row.recommended_first_surface} matches=${row.browser_match_count} provider=${row.provider_path ?? 'none'} gaps=${row.production_gap_count} handoff=${row.handoff_file}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
  }
  if (!report.ok) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
