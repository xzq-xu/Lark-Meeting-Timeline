#!/usr/bin/env node

import { mkdir, readdir, readFile, stat, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  MEETING_PLATFORM_EVIDENCE_PACKAGE_SCHEMA,
} from '../packages/meeting-timeline-sdk/adapters/platform-evidence-package.mjs';
import {
  buildMeetingPlatformFieldCaptureMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-field-capture.mjs';
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
const failOnIncomplete = args.get('fail-on-incomplete') === 'true' || args.get('fail-on-missing') === 'true';
const requireProductionReady = args.get('require-production-ready') === 'true';
const planOnly = args.get('plan-only') === 'true';
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || '');
const evidenceDir = String(args.get('evidence-dir') || args.get('evidenceDir') || 'data');
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

function summarizePlan(plan = {}) {
  return {
    platform: plan.platform,
    display_name: plan.display_name,
    status: plan.status,
    production_ready: plan.production_ready,
    ready_for_realtime_annotations: plan.ready_for_realtime_annotations,
    local_required_records: plan.local_observer?.minimum_record_count ?? 0,
    provider_required_records: plan.provider_events?.minimum_record_count ?? 0,
    provider_missing_env: plan.provider_events?.security?.missing_env ?? [],
    missing_items: plan.missing_items ?? [],
    output_paths: plan.output_paths,
    checklist: (plan.checklist ?? []).map((item) => ({
      id: item.id,
      channel: item.channel,
      when: item.when,
      output: item.output,
      required_coverage: item.required_coverage,
      required_snapshot: item.required_snapshot,
    })),
    current_evidence: plan.current_evidence ? {
      package_id: plan.current_evidence.package_id,
      status: plan.current_evidence.status,
      production_ready: plan.current_evidence.production_ready,
      ready_for_realtime_annotations: plan.current_evidence.ready_for_realtime_annotations,
      provider_record_count: plan.current_evidence.provider_record_count,
      meeting_app_record_count: plan.current_evidence.meeting_app_record_count,
      correlation_status: plan.current_evidence.correlation_status,
      correlation_passed: plan.current_evidence.correlation_passed,
    } : undefined,
    next_actions: plan.next_actions ?? [],
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
  const matrix = buildMeetingPlatformFieldCaptureMatrix({
    baseUrl: baseUrl || undefined,
    env: process.env,
    evidenceDir,
    platforms: requiredPlatforms,
    evidencePackage,
  });
  const complete = requireProductionReady
    ? matrix.production_ready_count === matrix.platform_count
    : matrix.missing_item_count === 0;
  const ok = matrix.platform_count > 0 && (planOnly || complete) && errors.length === 0;
  return {
    type: 'meeting_platform_field_capture_report',
    ok,
    plan_only: planOnly,
    requirement: requireProductionReady ? 'production_ready' : 'no_missing_capture_items',
    require_production_ready: requireProductionReady,
    base_url: baseUrl || undefined,
    evidence_dir: evidenceDir,
    input_dir: inputList.length > 0 ? null : inputDir,
    input_files: files,
    file_count: files.length,
    evaluated_file_count: packages.length,
    package_count: packages.length,
    schema_ok_count: packages.filter((item) => item.schema_ok).length,
    required_platforms: requiredPlatforms,
    package_files_by_platform: Object.fromEntries([...packageIndex.entries()].map(([platform, item]) => [platform, item.file])),
    platform_count: matrix.platform_count,
    production_ready_count: matrix.production_ready_count,
    realtime_ready_count: matrix.realtime_ready_count,
    missing_item_count: matrix.missing_item_count,
    rows: matrix.plans.map((plan) => summarizePlan(plan)),
    matrix,
    next_actions: unique(matrix.next_actions ?? []),
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
    console.log(`meeting_platform_field_capture_report | ok=${boolLabel(report.ok)} | requirement=${report.requirement} | production_ready=${report.production_ready_count}/${report.platform_count} | realtime_ready=${report.realtime_ready_count} | missing_items=${report.missing_item_count} | packages=${report.evaluated_file_count}/${report.file_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: status=${row.status} production_ready=${boolLabel(row.production_ready)} realtime_ready=${boolLabel(row.ready_for_realtime_annotations)} missing=${row.missing_items.length} local_required=${row.local_required_records} provider_required=${row.provider_required_records}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
    for (const error of report.errors) console.error(`error ${error.file}: ${error.error}`);
  }
  if (!report.ok && failOnIncomplete) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
