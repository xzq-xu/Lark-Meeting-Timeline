#!/usr/bin/env node

import { mkdir, readdir, readFile, stat, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  MEETING_PLATFORM_EVIDENCE_PACKAGE_SCHEMA,
} from '../packages/meeting-timeline-sdk/adapters/platform-evidence-package.mjs';
import {
  runMeetingPlatformHandoffReadinessMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-handoff-readiness.mjs';
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
const packageDir = resolve(String(
  args.get('package-dir')
    || args.get('evidence-package-dir')
    || args.get('handoff-dir')
    || 'data/meeting-platform-evidence-packages',
));
const packageInputList = String(
  args.get('package-inputs')
    || args.get('package-input')
    || args.get('evidence-package-inputs')
    || args.get('evidence-package-input')
    || '',
).split(',').map((item) => item.trim()).filter(Boolean);
const outDir = String(args.get('out-dir') || args.get('outDir') || '');
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const includeReports = args.get('include-reports') === 'true';
const failOnNotReady = args.get('fail-on-not-ready') === 'true';
const writeReports = args.get('write-reports') === 'true' || Boolean(outDir);
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

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

async function evidenceFiles() {
  if (packageInputList.length > 0) return packageInputList.map((item) => resolve(item));
  return collectJsonFiles(packageDir);
}

async function readEvidencePackage(file) {
  const text = await readFile(file, 'utf8');
  const input = JSON.parse(text);
  const stats = await stat(file);
  if (input?.schema !== MEETING_PLATFORM_EVIDENCE_PACKAGE_SCHEMA) {
    return {
      file,
      ignored: true,
      schema: input?.schema,
      modified_at_ms: Math.round(stats.mtimeMs),
    };
  }
  return {
    file,
    ignored: false,
    schema: input.schema,
    platform: normalizePlatformKey(input.platform ?? input.rollout_plan?.platform),
    package: input,
    modified_at_ms: Math.round(stats.mtimeMs),
  };
}

async function loadPackages() {
  const files = await evidenceFiles();
  const packages = new Map();
  const loaded = [];
  const errors = [];
  for (const file of files) {
    try {
      const record = await readEvidencePackage(file);
      loaded.push(record);
      if (record.platform && record.ignored !== true) {
        const current = packages.get(record.platform);
        if (!current || record.modified_at_ms >= current.modified_at_ms) packages.set(record.platform, record);
      }
    } catch (error) {
      errors.push({ file, error: String(error?.message ?? error) });
    }
  }
  return {
    files,
    loaded,
    errors,
    packages,
  };
}

function buildMatrixInput(loaded) {
  const input = {
    platforms: requiredPlatforms,
  };
  for (const rawPlatform of requiredPlatforms) {
    const platform = normalizePlatformKey(rawPlatform);
    input[platform] = {
      evidencePackage: loaded.packages.get(platform)?.package,
    };
  }
  return input;
}

function summarizeRow(row = {}, loaded) {
  return {
    platform: row.platform,
    display_name: row.display_name,
    status: row.status,
    handoff_ready: row.handoff_ready,
    pilot_ready: row.pilot_ready,
    production_ready: row.production_ready,
    runtime_host_replay_accepted: row.runtime_host_replay_accepted,
    runtime_host_replay_missing: row.runtime_host_replay_missing ?? [],
    local_observer_ready: row.local_observer_ready,
    provider_reconcile_ready: row.provider_reconcile_ready,
    provider_missing_env: row.provider_missing_env ?? [],
    provider_record_count: row.provider_record_count,
    meeting_app_record_count: row.meeting_app_record_count,
    dom_record_count: row.dom_record_count,
    evidence_package_file: loaded.packages.get(row.platform)?.file,
    next_actions: row.next_actions ?? [],
  };
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function reportFileFor(platform) {
  return resolve(outDir || 'data/meeting-platform-handoff-readiness', `${platform}.json`);
}

async function buildReport() {
  const loaded = await loadPackages();
  const matrixInput = buildMatrixInput(loaded);
  const matrix = await runMeetingPlatformHandoffReadinessMatrix(matrixInput, {
    baseUrl,
    env: process.env,
  });
  const writtenFiles = [];
  if (writeReports) {
    for (const report of matrix.reports) {
      const file = reportFileFor(report.platform);
      await writeJson(file, report);
      writtenFiles.push(file);
    }
  }
  const rows = matrix.rows.map((row) => summarizeRow(row, loaded));
  const ok = matrix.platform_count > 0
    && matrix.handoff_ready_count === matrix.platform_count
    && loaded.errors.length === 0;
  return {
    type: 'meeting_platform_handoff_readiness_report',
    ok,
    requirement: 'all_required_platforms_handoff_ready',
    base_url: baseUrl,
    package_evidence_dir: packageInputList.length > 0 ? null : packageDir,
    package_input_files: loaded.files,
    loaded_package_count: [...loaded.packages.values()].length,
    ignored_file_count: loaded.loaded.filter((item) => item.ignored).length,
    required_platforms: requiredPlatforms,
    platform_count: matrix.platform_count,
    handoff_ready_count: matrix.handoff_ready_count,
    pilot_ready_count: matrix.pilot_ready_count,
    production_ready_count: matrix.production_ready_count,
    runtime_host_replay_ready_count: matrix.runtime_host_replay_ready_count,
    local_observer_ready_count: matrix.local_observer_ready_count,
    provider_reconcile_ready_count: matrix.provider_reconcile_ready_count,
    provider_setup_needed_count: matrix.provider_setup_needed_count,
    local_evidence_needed_count: matrix.local_evidence_needed_count,
    written_files: writtenFiles,
    rows,
    matrix: includeReports ? matrix : undefined,
    next_actions: matrix.next_actions,
    errors: loaded.errors,
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_handoff_readiness_report | ok=${boolLabel(report.ok)} | handoff_ready=${report.handoff_ready_count}/${report.platform_count} | pilot_ready=${report.pilot_ready_count} | production_ready=${report.production_ready_count} | runtime_replay=${report.runtime_host_replay_ready_count}/${report.platform_count} | packages=${report.loaded_package_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: status=${row.status} handoff=${boolLabel(row.handoff_ready)} pilot=${boolLabel(row.pilot_ready)} production=${boolLabel(row.production_ready)} runtime_replay=${boolLabel(row.runtime_host_replay_accepted)} local=${boolLabel(row.local_observer_ready)} provider=${boolLabel(row.provider_reconcile_ready)} package=${row.evidence_package_file ?? '-'}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
    for (const error of report.errors) console.error(`error ${error.file}: ${error.error}`);
  }
  if (!report.ok && failOnNotReady) process.exitCode = 2;
} catch (error) {
  if (jsonOutput) console.error(JSON.stringify({
    type: 'meeting_platform_handoff_readiness_report',
    ok: false,
    error: String(error?.message ?? error),
  }, null, 2));
  else console.error(`meeting_platform_handoff_readiness_report | ok=no | error=${String(error?.message ?? error)}`);
  process.exitCode = 2;
}
