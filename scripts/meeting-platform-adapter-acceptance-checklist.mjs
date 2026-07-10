#!/usr/bin/env node

import { mkdir, readdir, readFile, stat, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdapterAcceptanceChecklistMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-acceptance-checklist.mjs';
import {
  MEETING_PLATFORM_EVIDENCE_PACKAGE_SCHEMA,
} from '../packages/meeting-timeline-sdk/adapters/platform-evidence-package.mjs';
import {
  normalizeMeetingPlatform,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const outDir = resolve(String(args.get('out-dir') || args.get('outDir') || 'data/meeting-platform-adapter-acceptance-checklists'));
const packageDir = resolve(String(
  args.get('package-dir')
    || args.get('evidence-package-dir')
    || 'data/meeting-platform-evidence-packages',
));
const packageInputList = String(
  args.get('package-inputs')
    || args.get('package-input')
    || args.get('evidence-package-inputs')
    || args.get('evidence-package-input')
    || '',
).split(',').map((item) => item.trim()).filter(Boolean);
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const writeChecklists = args.get('write-checklists') !== 'false';
const failOnIncomplete = args.get('fail-on-incomplete') === 'true' || args.get('fail-on-not-ready') === 'true';
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

function normalizePlatformKey(platform) {
  if (!platform) return undefined;
  try {
    return normalizeMeetingPlatform(platform);
  } catch {
    return String(platform);
  }
}

async function collectJsonFiles(dir) {
  const files = [];
  async function walk(current) {
    let entries = [];
    try {
      entries = await readdir(current, { withFileTypes: true });
    } catch {
      return;
    }
    for (const entry of entries) {
      const path = resolve(current, entry.name);
      if (entry.isDirectory()) await walk(path);
      else if (entry.isFile() && entry.name.endsWith('.json')) files.push(path);
    }
  }
  await walk(dir);
  return files.sort();
}

async function loadEvidencePackages() {
  const files = packageInputList.length > 0
    ? packageInputList.map((item) => resolve(item))
    : await collectJsonFiles(packageDir);
  const packages = new Map();
  const errors = [];
  let ignoredCount = 0;
  for (const file of files) {
    try {
      const input = JSON.parse(await readFile(file, 'utf8'));
      if (input?.schema !== MEETING_PLATFORM_EVIDENCE_PACKAGE_SCHEMA) {
        ignoredCount += 1;
        continue;
      }
      const platform = normalizePlatformKey(input.platform ?? input.rollout_plan?.platform);
      const stats = await stat(file);
      const record = { file, package: input, modified_at_ms: Math.round(stats.mtimeMs) };
      const current = packages.get(platform);
      if (platform && (!current || record.modified_at_ms >= current.modified_at_ms)) packages.set(platform, record);
    } catch (error) {
      errors.push({ file, error: String(error?.message ?? error) });
    }
  }
  return { files, packages, errors, ignoredCount };
}

async function buildReport() {
  const loaded = await loadEvidencePackages();
  const matrixInput = {
    platforms: requiredPlatforms,
  };
  for (const rawPlatform of requiredPlatforms) {
    const platform = normalizePlatformKey(rawPlatform);
    matrixInput[platform] = {
      evidencePackage: loaded.packages.get(platform)?.package,
    };
  }
  const matrix = buildMeetingPlatformAdapterAcceptanceChecklistMatrix(matrixInput, {
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
    ok: matrix.platform_count > 0 && matrix.accepted_count === matrix.platform_count && loaded.errors.length === 0,
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
    package_evidence_dir: packageInputList.length > 0 ? null : packageDir,
    package_input_files: loaded.files,
    loaded_package_count: loaded.packages.size,
    ignored_package_file_count: loaded.ignoredCount,
    written_files: writtenFiles,
    rows: matrix.rows.map((row) => ({
      ...row,
      checklist_file: checklistFile(row.platform),
      evidence_package_file: loaded.packages.get(row.platform)?.file,
    })),
    matrix: {
      ...matrix,
      checklists: undefined,
    },
    next_actions: matrix.next_actions ?? [],
    errors: loaded.errors,
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_adapter_acceptance_checklist_report | ok=${boolLabel(report.ok)} | target=${report.target} | accepted=${report.accepted_count}/${report.platform_count} | blocked=${report.blocked_count} | pilot=${report.pilot_ready_count} | production=${report.production_ready_count} | packages=${report.loaded_package_count} | written=${report.written_files.length}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: accepted=${boolLabel(row.accepted)} target=${row.target} blocking=${row.blocking_count} status=${row.handoff_status ?? '-'} package=${row.evidence_package_file ?? '-'} next=${row.first_next_action ?? 'none'}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
    for (const error of report.errors) console.error(`error ${error.file}: ${error.error}`);
  }
  if (!report.ok && (report.target === 'static' || failOnIncomplete)) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
