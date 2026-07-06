#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdapterAcceptanceChecklistMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-acceptance-checklist.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const outDir = resolve(String(args.get('out-dir') || args.get('outDir') || 'data/meeting-platform-adapter-acceptance-checklists'));
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const writeChecklists = args.get('write-checklists') !== 'false';
const target = String(args.get('target') || args.get('acceptance-target') || args.get('acceptanceTarget') || 'pilot');
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()).filter(Boolean));

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function checklistFile(platform) {
  return resolve(outDir, `${platform}.json`);
}

async function buildReport() {
  const matrix = buildMeetingPlatformAdapterAcceptanceChecklistMatrix({
    platforms: requiredPlatforms,
  }, {
    baseUrl,
    env: process.env,
    target,
  });
  const writtenFiles = [];
  if (writeChecklists) {
    for (const checklist of matrix.checklists) {
      const file = checklistFile(checklist.platform);
      await writeJson(file, checklist);
      writtenFiles.push(file);
    }
  }
  return {
    type: 'meeting_platform_adapter_acceptance_checklist_report',
    ok: matrix.platform_count > 0 && matrix.accepted_count === matrix.platform_count,
    requirement: 'meeting_platform_adapter_acceptance_checklists_available_for_static_pilot_or_production_gate',
    base_url: baseUrl,
    target,
    out_dir: writeChecklists ? outDir : undefined,
    write_checklists: writeChecklists,
    platform_count: matrix.platform_count,
    accepted_count: matrix.accepted_count,
    blocked_count: matrix.blocked_count,
    static_ready_count: matrix.static_ready_count,
    pilot_ready_count: matrix.pilot_ready_count,
    production_ready_count: matrix.production_ready_count,
    required_platforms: requiredPlatforms,
    written_files: writtenFiles,
    rows: matrix.rows.map((row) => ({
      ...row,
      checklist_file: checklistFile(row.platform),
    })),
    matrix: {
      ...matrix,
      checklists: undefined,
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
    console.log(`meeting_platform_adapter_acceptance_checklist_report | ok=${boolLabel(report.ok)} | target=${report.target} | accepted=${report.accepted_count}/${report.platform_count} | blocked=${report.blocked_count} | pilot=${report.pilot_ready_count} | production=${report.production_ready_count} | written=${report.written_files.length}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: accepted=${boolLabel(row.accepted)} target=${row.target} blocking=${row.blocking_count} status=${row.handoff_status ?? '-'} next=${row.first_next_action ?? 'none'}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
  }
  if (!report.ok) process.exitCode = report.target === 'static' ? 2 : 0;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
