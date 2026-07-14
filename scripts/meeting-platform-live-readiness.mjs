#!/usr/bin/env node

import { mkdir, readdir, readFile, stat, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  MEETING_PLATFORM_EVIDENCE_PACKAGE_SCHEMA,
} from '../packages/meeting-timeline-sdk/adapters/platform-evidence-package.mjs';
import {
  buildMeetingPlatformLiveAdapterReadinessMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-live-adapter.mjs';
import {
  normalizeMeetingPlatform,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const inputDir = resolve(String(args.get('input-dir') || args.get('dir') || 'data/meeting-platform-evidence-packages'));
const inputList = String(args.get('inputs') || args.get('input') || args.get('evidence-package-inputs') || args.get('evidence-package-input') || '').split(',').map((item) => item.trim()).filter(Boolean);
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const failOnIncomplete = args.get('fail-on-incomplete') === 'true';
const requireProductionReady = args.get('require-production-ready') !== 'false';
const target = String(args.get('target') || (requireProductionReady ? 'production' : 'pilot'));
const requireCorrelation = args.get('require-correlation') !== 'false';
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || '');
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

function normalizePlatformKey(platform) {
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
  if (inputList.length > 0) return inputList.map((item) => resolve(item));
  return collectJsonFiles(inputDir);
}

async function readPackage(file) {
  const text = await readFile(file, 'utf8');
  const input = JSON.parse(text);
  const stats = await stat(file);
  return {
    file,
    input,
    platform: normalizePlatformKey(input.platform ?? input.rollout_plan?.platform),
    schema_ok: input?.schema === MEETING_PLATFORM_EVIDENCE_PACKAGE_SCHEMA,
    bytes: Buffer.byteLength(text, 'utf8'),
    modified_at_ms: Math.round(stats.mtimeMs),
  };
}

function buildPackageIndex(packages = []) {
  const byPlatform = new Map();
  for (const item of packages) {
    if (!item.platform) continue;
    const current = byPlatform.get(item.platform);
    if (!current || item.modified_at_ms >= current.modified_at_ms) byPlatform.set(item.platform, item);
  }
  return byPlatform;
}

function summarizeReadiness(row = {}) {
  return {
    platform: row.platform,
    display_name: row.display_name,
    target: row.target,
    status: row.status,
    passed: row.passed,
    rollout_status: row.rollout_status,
    ready_for_realtime_annotations: row.ready_for_realtime_annotations,
    production_ready: row.production_ready,
    blocking_count: row.blocking_count,
    warning_count: row.warning_count,
    blocking_codes: (row.blocking_checks ?? []).map((item) => item.code),
    warning_codes: (row.warnings ?? []).map((item) => item.code),
    verification: row.verification ? {
      passed: row.verification.passed,
      requirement: row.verification.requirement,
      correlation_passed: row.verification.correlation_passed,
      correlation_status: row.verification.correlation_status,
      embedded_plan_matches: row.verification.embedded_plan_matches,
      provider_record_count: row.verification.provider_record_count,
      meeting_app_record_count: row.verification.meeting_app_record_count,
    } : undefined,
    next_actions: row.next_actions ?? [],
  };
}

async function buildReport() {
  const files = await evidenceFiles();
  const packages = [];
  const errors = [];
  for (const file of files) {
    try {
      packages.push(await readPackage(file));
    } catch (error) {
      errors.push({ file, error: String(error?.message ?? error) });
    }
  }
  const packageIndex = buildPackageIndex(packages);
  const evidencePackage = Object.fromEntries([...packageIndex.entries()].map(([platform, item]) => [platform, item.input]));
  const matrix = buildMeetingPlatformLiveAdapterReadinessMatrix({
    baseUrl: baseUrl || undefined,
    env: process.env,
    platforms: requiredPlatforms,
    evidencePackage,
    target,
    requireCorrelation,
    require_correlation: requireCorrelation,
    requireProductionReady,
    require_production_ready: requireProductionReady,
  });
  const ok = matrix.platform_count > 0 && matrix.reports.every((row) => row.passed) && errors.length === 0;
  return {
    type: 'meeting_platform_live_readiness_report',
    ok,
    target: target === 'production' ? 'production' : 'pilot',
    require_production_ready: requireProductionReady,
    require_correlation: requireCorrelation,
    base_url: baseUrl || undefined,
    input_dir: inputList.length > 0 ? null : inputDir,
    input_files: files,
    file_count: files.length,
    evaluated_file_count: packages.length,
    package_count: packages.length,
    schema_ok_count: packages.filter((item) => item.schema_ok).length,
    required_platforms: requiredPlatforms,
    package_files_by_platform: Object.fromEntries([...packageIndex.entries()].map(([platform, item]) => [platform, item.file])),
    platform_count: matrix.platform_count,
    passed_count: matrix.passed_count,
    ready_count: matrix.ready_count,
    warning_count: matrix.warning_count,
    blocked_count: matrix.blocked_count,
    realtime_ready_count: matrix.realtime_ready_count,
    production_ready_count: matrix.production_ready_count,
    rows: matrix.reports.map((row) => summarizeReadiness(row)),
    matrix,
    next_actions: unique(matrix.reports.flatMap((row) => row.next_actions ?? [])),
    errors,
  };
}

try {
  const report = await buildReport();
  if (reportFile) {
    const targetFile = resolve(reportFile);
    await mkdir(dirname(targetFile), { recursive: true });
    await writeFile(targetFile, `${JSON.stringify(report, null, 2)}\n`, 'utf8');
  }
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_live_readiness_report | ok=${boolLabel(report.ok)} | target=${report.target} | passed=${report.passed_count}/${report.platform_count} | ready=${report.ready_count} | warning=${report.warning_count} | blocked=${report.blocked_count} | packages=${report.evaluated_file_count}/${report.file_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: status=${row.status} passed=${boolLabel(row.passed)} rollout=${row.rollout_status} production_ready=${boolLabel(row.production_ready)} realtime_ready=${boolLabel(row.ready_for_realtime_annotations)} blocking=${row.blocking_count} warning=${row.warning_count}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
    for (const error of report.errors) console.error(`error ${error.file}: ${error.error}`);
  }
  if (!report.ok && failOnIncomplete) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
