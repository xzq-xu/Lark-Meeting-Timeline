#!/usr/bin/env node

import { mkdir, readdir, readFile, stat, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  PLATFORM_EVENT_CAPTURE_SCHEMA,
} from '../packages/meeting-timeline-sdk/adapters/platform-capture.mjs';
import {
  MEETING_PLATFORM_EVIDENCE_PACKAGE_SCHEMA,
} from '../packages/meeting-timeline-sdk/adapters/platform-evidence-package.mjs';
import {
  MEETING_APP_SNAPSHOT_RECORD_SCHEMA,
  meetingAppSnapshotRecords,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';
import {
  runMeetingPlatformRealEvidenceIntakeMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-real-intake.mjs';
import {
  normalizeMeetingPlatform,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';
import {
  asArray,
  boolLabel,
  extractRecordSet,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();

const providerDir = resolve(String(
  args.get('provider-dir')
    || args.get('platform-dir')
    || args.get('event-dir')
    || 'data/provider-evidence',
));
const domDir = resolve(String(
  args.get('dom-dir')
    || args.get('meeting-app-dir')
    || args.get('app-dir')
    || 'data/meeting-app-evidence',
));
const packageDir = resolve(String(
  args.get('package-dir')
    || args.get('evidence-package-dir')
    || args.get('handoff-dir')
    || 'data/meeting-platform-evidence-packages',
));
const providerInputList = String(args.get('provider-inputs') || args.get('provider-input') || '').split(',').map((item) => item.trim()).filter(Boolean);
const domInputList = String(args.get('dom-inputs') || args.get('dom-input') || '').split(',').map((item) => item.trim()).filter(Boolean);
const packageInputList = String(args.get('package-inputs') || args.get('package-input') || args.get('evidence-package-inputs') || args.get('evidence-package-input') || '').split(',').map((item) => item.trim()).filter(Boolean);
const jsonOutput = args.get('json') === 'true';
const includeMatrix = args.get('include-matrix') === 'true';
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const requireProductionReady = args.get('require-production-ready') !== 'false';
const allowFixtureEvidence = args.get('allow-fixture-evidence') === 'true';
const failOnIncomplete = args.get('fail-on-incomplete') === 'true';
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

async function evidenceFiles(inputList, dir) {
  if (inputList.length > 0) return inputList.map((item) => resolve(item));
  return collectJsonFiles(dir);
}

function providerRecordsFromInput(input = {}) {
  const rows = [];
  if (Array.isArray(input)) rows.push(...input);
  else if (input?.schema === PLATFORM_EVENT_CAPTURE_SCHEMA) rows.push(input);
  else {
    rows.push(
      ...asArray(input.records),
      ...asArray(input.capture_records),
      ...asArray(input.captureRecords),
      ...asArray(input.provider_records),
      ...asArray(input.providerRecords),
      ...asArray(input.provider_capture_records),
      ...asArray(input.providerCaptureRecords),
    );
  }
  return rows.filter((record) => record?.schema === PLATFORM_EVENT_CAPTURE_SCHEMA || record?.body || record?.raw_body);
}

function meetingAppRecordsFromInput(input = {}) {
  const rows = [];
  if (Array.isArray(input)) rows.push(...input);
  else if (input?.schema === MEETING_APP_SNAPSHOT_RECORD_SCHEMA) rows.push(input);
  else {
    const recordSet = extractRecordSet(input);
    rows.push(
      ...meetingAppSnapshotRecords(recordSet),
      ...asArray(input.meeting_app_records),
      ...asArray(input.meetingAppRecords),
      ...asArray(input.meeting_app_snapshot_records),
      ...asArray(input.meetingAppSnapshotRecords),
      ...asArray(input.dom_records),
      ...asArray(input.domRecords),
    );
  }
  return rows.filter((record) => record?.schema === MEETING_APP_SNAPSHOT_RECORD_SCHEMA || record?.snapshot);
}

function platformFromRecord(record = {}) {
  return normalizePlatformKey(record.platform ?? record.provider ?? record.adapter ?? record.snapshot?.platform ?? record.snapshot?.provider);
}

function platformsFromInput(input = {}, providerRecords = [], meetingAppRecords = []) {
  return unique([
    ...asArray(input.platforms),
    ...asArray(input.platform),
    input.rollout_plan?.platform,
    ...providerRecords.map((record) => record.platform ?? record.provider ?? record.adapter),
    ...meetingAppRecords.map((record) => record.platform ?? record.provider ?? record.snapshot?.platform ?? record.snapshot?.provider),
  ].map((item) => normalizePlatformKey(item)).filter(Boolean));
}

async function readEvidenceFile(file, kind) {
  const text = await readFile(file, 'utf8');
  const input = JSON.parse(text);
  const stats = await stat(file);
  const providerRecords = kind === 'dom' ? [] : providerRecordsFromInput(input);
  const meetingAppRecords = kind === 'provider' ? [] : meetingAppRecordsFromInput(input);
  const evidencePackage = input?.schema === MEETING_PLATFORM_EVIDENCE_PACKAGE_SCHEMA ? input : undefined;
  return {
    kind,
    file,
    input,
    schema: input?.schema,
    providerRecords,
    meetingAppRecords,
    evidencePackage,
    platforms: platformsFromInput(input, providerRecords, meetingAppRecords),
    bytes: Buffer.byteLength(text, 'utf8'),
    modified_at_ms: Math.round(stats.mtimeMs),
  };
}

function addMapList(map, key, values) {
  if (!key) return;
  if (!map.has(key)) map.set(key, []);
  map.get(key).push(...values);
}

function aggregateEvidence(evaluations = []) {
  const providerRecordsByPlatform = new Map();
  const meetingAppRecordsByPlatform = new Map();
  const packageByPlatform = new Map();
  const sourceFilesByPlatform = new Map();

  for (const evidence of evaluations) {
    for (const record of evidence.providerRecords) {
      const platform = platformFromRecord(record);
      addMapList(providerRecordsByPlatform, platform, [record]);
      addMapList(sourceFilesByPlatform, platform, [evidence.file]);
    }
    for (const record of evidence.meetingAppRecords) {
      const platform = platformFromRecord(record);
      addMapList(meetingAppRecordsByPlatform, platform, [record]);
      addMapList(sourceFilesByPlatform, platform, [evidence.file]);
    }
    if (evidence.evidencePackage) {
      const platform = normalizePlatformKey(evidence.evidencePackage.platform ?? evidence.evidencePackage.rollout_plan?.platform);
      const current = packageByPlatform.get(platform);
      if (!current || evidence.modified_at_ms >= current.modified_at_ms) {
        packageByPlatform.set(platform, evidence);
      }
      addMapList(sourceFilesByPlatform, platform, [evidence.file]);
    }
  }

  return {
    providerRecordsByPlatform,
    meetingAppRecordsByPlatform,
    packageByPlatform,
    sourceFilesByPlatform,
  };
}

async function loadEvidence() {
  const providerFiles = await evidenceFiles(providerInputList, providerDir);
  const domFiles = await evidenceFiles(domInputList, domDir);
  const packageFiles = await evidenceFiles(packageInputList, packageDir);
  const providerEvaluations = [];
  const domEvaluations = [];
  const packageEvaluations = [];
  const errors = [];

  for (const file of providerFiles) {
    try {
      providerEvaluations.push(await readEvidenceFile(file, 'provider'));
    } catch (error) {
      errors.push({ kind: 'provider', file, error: String(error?.message ?? error) });
    }
  }
  for (const file of domFiles) {
    try {
      domEvaluations.push(await readEvidenceFile(file, 'dom'));
    } catch (error) {
      errors.push({ kind: 'dom', file, error: String(error?.message ?? error) });
    }
  }
  for (const file of packageFiles) {
    try {
      packageEvaluations.push(await readEvidenceFile(file, 'package'));
    } catch (error) {
      errors.push({ kind: 'package', file, error: String(error?.message ?? error) });
    }
  }

  return { providerFiles, domFiles, packageFiles, providerEvaluations, domEvaluations, packageEvaluations, errors };
}

function buildMatrixInput(aggregate) {
  const input = {
    platforms: requiredPlatforms,
  };
  for (const rawPlatform of requiredPlatforms) {
    const platform = normalizePlatformKey(rawPlatform);
    const packageEvidence = aggregate.packageByPlatform.get(platform)?.input;
    input[platform] = {
      providerRecords: aggregate.providerRecordsByPlatform.get(platform) ?? [],
      meetingAppRecords: aggregate.meetingAppRecordsByPlatform.get(platform) ?? [],
      evidencePackage: packageEvidence,
    };
  }
  return input;
}

function summarizeRow(row = {}, aggregate) {
  const platform = row.platform;
  return {
    platform,
    accepted: row.accepted,
    readiness_status: row.readiness?.status,
    verification_passed: row.verification?.passed,
    runtime_host_replay_accepted: row.runtime_host_replay_accepted,
    runtime_host_replay_missing: row.runtime_host_replay?.missing ?? [],
    provider_record_count: row.provider_record_count,
    meeting_app_record_count: row.meeting_app_record_count,
    fixture_evidence_count: row.fixture_evidence_count,
    blocking_count: row.blocking_count,
    production_ready: row.readiness?.production_ready ?? false,
    ready_for_realtime_annotations: row.readiness?.ready_for_realtime_annotations ?? false,
    blocking_codes: (row.blocking_checks ?? []).map((item) => item.code),
    next_actions: row.next_actions ?? [],
    source_files: unique(aggregate.sourceFilesByPlatform.get(platform) ?? []),
  };
}

async function buildReport(loaded) {
  const aggregate = aggregateEvidence([
    ...loaded.providerEvaluations,
    ...loaded.domEvaluations,
    ...loaded.packageEvaluations,
  ]);
  const matrixInput = buildMatrixInput(aggregate);
  const matrix = await runMeetingPlatformRealEvidenceIntakeMatrix(matrixInput, {
    baseUrl,
    env: process.env,
    requireProductionReady,
    require_production_ready: requireProductionReady,
    allowFixtureEvidence,
    allow_fixture_evidence: allowFixtureEvidence,
  });
  const rows = matrix.reports.map((report) => summarizeRow(report, aggregate));
  const ok = matrix.platform_count > 0 && matrix.rejected_count === 0 && loaded.errors.length === 0;
  return {
    type: 'meeting_platform_real_intake_report',
    ok,
    require_production_ready: requireProductionReady,
    allow_fixture_evidence: allowFixtureEvidence,
    base_url: baseUrl,
    provider_evidence_dir: providerInputList.length > 0 ? null : providerDir,
    dom_evidence_dir: domInputList.length > 0 ? null : domDir,
    package_evidence_dir: packageInputList.length > 0 ? null : packageDir,
    provider_input_files: loaded.providerFiles,
    dom_input_files: loaded.domFiles,
    package_input_files: loaded.packageFiles,
    provider_file_count: loaded.providerFiles.length,
    dom_file_count: loaded.domFiles.length,
    package_file_count: loaded.packageFiles.length,
    provider_evaluated_file_count: loaded.providerEvaluations.length,
    dom_evaluated_file_count: loaded.domEvaluations.length,
    package_evaluated_file_count: loaded.packageEvaluations.length,
    required_platforms: requiredPlatforms,
    platform_count: matrix.platform_count,
    accepted_count: matrix.accepted_count,
    rejected_count: matrix.rejected_count,
    production_ready_count: matrix.production_ready_count,
    realtime_ready_count: matrix.realtime_ready_count,
    runtime_host_replay_ready_count: matrix.runtime_host_replay_ready_count,
    rows,
    matrix: includeMatrix ? matrix : undefined,
    next_actions: unique(rows.flatMap((row) => row.next_actions)),
    errors: loaded.errors,
  };
}

try {
  const loaded = await loadEvidence();
  const report = await buildReport(loaded);
  if (loaded.errors.length > 0 && loaded.providerEvaluations.length === 0 && loaded.domEvaluations.length === 0 && loaded.packageEvaluations.length === 0) {
    report.ok = false;
  }
  if (reportFile) {
    const target = resolve(reportFile);
    await mkdir(dirname(target), { recursive: true });
    await writeFile(target, `${JSON.stringify(report, null, 2)}\n`, 'utf8');
  }
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_real_intake_report | ok=${boolLabel(report.ok)} | accepted=${report.accepted_count}/${report.platform_count} | production_ready=${report.production_ready_count} | realtime_ready=${report.realtime_ready_count} | runtime_replay=${report.runtime_host_replay_ready_count}/${report.platform_count} | provider_files=${report.provider_evaluated_file_count}/${report.provider_file_count} | dom_files=${report.dom_evaluated_file_count}/${report.dom_file_count} | package_files=${report.package_evaluated_file_count}/${report.package_file_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: accepted=${boolLabel(row.accepted)} readiness=${row.readiness_status} verification=${boolLabel(row.verification_passed)} runtime_replay=${boolLabel(row.runtime_host_replay_accepted)} provider_records=${row.provider_record_count} dom_records=${row.meeting_app_record_count} fixture=${row.fixture_evidence_count} blocking=${row.blocking_count}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
    for (const error of report.errors) console.error(`error ${error.kind} ${error.file}: ${error.error}`);
  }
  if (!report.ok && failOnIncomplete) process.exitCode = 2;
} catch (error) {
  const payload = {
    type: 'meeting_platform_real_intake_report',
    ok: false,
    provider_evidence_dir: providerDir,
    dom_evidence_dir: domDir,
    package_evidence_dir: packageDir,
    error: String(error?.message ?? error),
  };
  if (jsonOutput) console.error(JSON.stringify(payload, null, 2));
  else console.error(`meeting_platform_real_intake_report | ok=no | error=${payload.error}`);
  process.exitCode = 2;
}
