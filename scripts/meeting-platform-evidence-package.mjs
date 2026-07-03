#!/usr/bin/env node

import { mkdir, readdir, readFile, stat, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  MEETING_PLATFORM_EVIDENCE_PACKAGE_SCHEMA,
  verifyMeetingPlatformEvidencePackage,
} from '../packages/meeting-timeline-sdk/adapters/platform-evidence-package.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const inputDir = resolve(String(args.get('input-dir') || args.get('dir') || 'data/meeting-platform-evidence-packages'));
const inputList = String(args.get('inputs') || args.get('input') || '').split(',').map((item) => item.trim()).filter(Boolean);
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const failOnIncomplete = args.get('fail-on-incomplete') === 'true';
const requireProductionReady = args.get('require-production-ready') !== 'false';
const requireCorrelation = args.get('require-correlation') !== 'false';
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || '');

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
    bytes: Buffer.byteLength(text, 'utf8'),
    modified_at_ms: Math.round(stats.mtimeMs),
  };
}

function verifyInput(evidence = {}) {
  const options = {
    env: process.env,
    requireProductionReady,
    requireCorrelation,
  };
  if (baseUrl) options.baseUrl = baseUrl;
  return verifyMeetingPlatformEvidencePackage(evidence.input, options);
}

function summarizeRow(evidence = {}, verification = {}) {
  return {
    file: evidence.file,
    package_id: verification.package_id,
    schema: evidence.input?.schema,
    schema_ok: evidence.input?.schema === MEETING_PLATFORM_EVIDENCE_PACKAGE_SCHEMA,
    platform: verification.platform,
    passed: verification.passed,
    requirement: verification.requirement,
    status: verification.status,
    production_ready: verification.production_ready,
    ready_for_realtime_annotations: verification.ready_for_realtime_annotations,
    embedded_plan_matches: verification.embedded_plan_matches,
    provider_record_count: verification.provider_record_count,
    provider_sample_count: verification.provider_sample_count,
    meeting_app_record_count: verification.meeting_app_record_count,
    correlation_required: verification.correlation_required,
    correlation_passed: verification.correlation_passed,
    correlation_status: verification.correlation_status,
    correlation_confidence: verification.correlation_confidence,
    provider_missing_required_coverage: verification.provider_missing_required_coverage,
    local_dom_missing_required_coverage: verification.local_dom_missing_required_coverage,
    next_actions: verification.next_actions,
  };
}

async function buildReport() {
  const files = await evidenceFiles();
  const rows = [];
  const verifications = [];
  const errors = [];
  for (const file of files) {
    try {
      const evidence = await readPackage(file);
      const verification = verifyInput(evidence);
      verifications.push(verification);
      rows.push(summarizeRow(evidence, verification));
    } catch (error) {
      errors.push({ file, error: String(error?.message ?? error) });
    }
  }
  const ok = rows.length > 0 && rows.every((row) => row.passed) && errors.length === 0;
  return {
    type: 'meeting_platform_evidence_package_report',
    ok,
    requirement: requireProductionReady ? 'production_ready' : 'ready_for_realtime_annotations',
    require_production_ready: requireProductionReady,
    require_correlation: requireCorrelation,
    input_dir: inputList.length > 0 ? null : inputDir,
    input_files: files,
    file_count: files.length,
    evaluated_file_count: rows.length,
    passed_count: rows.filter((row) => row.passed).length,
    production_ready_count: rows.filter((row) => row.production_ready).length,
    realtime_ready_count: rows.filter((row) => row.ready_for_realtime_annotations).length,
    correlation_passed_count: rows.filter((row) => row.correlation_passed).length,
    stale_embedded_plan_count: rows.filter((row) => row.embedded_plan_matches === false).length,
    rows,
    verifications,
    next_actions: unique(rows.flatMap((row) => row.next_actions)),
    errors,
  };
}

try {
  const report = await buildReport();
  if (reportFile) {
    const target = resolve(reportFile);
    await mkdir(dirname(target), { recursive: true });
    await writeFile(target, `${JSON.stringify(report, null, 2)}\n`, 'utf8');
  }
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_evidence_package_report | ok=${boolLabel(report.ok)} | requirement=${report.requirement} | passed=${report.passed_count}/${report.file_count} | production_ready=${report.production_ready_count} | realtime_ready=${report.realtime_ready_count} | correlation=${report.correlation_passed_count}/${report.evaluated_file_count} | stale=${report.stale_embedded_plan_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: status=${row.status} passed=${boolLabel(row.passed)} production_ready=${boolLabel(row.production_ready)} realtime_ready=${boolLabel(row.ready_for_realtime_annotations)} correlation=${row.correlation_status}/${row.correlation_confidence} provider_records=${row.provider_record_count} dom_records=${row.meeting_app_record_count}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
    for (const error of report.errors) console.error(`error ${error.file}: ${error.error}`);
  }
  if (!report.ok && failOnIncomplete) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
