#!/usr/bin/env node

import { mkdir, readdir, readFile, stat, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  PLATFORM_EVENT_CAPTURE_SCHEMA,
} from '../packages/meeting-timeline-sdk/adapters/platform-capture.mjs';
import {
  buildMeetingAppSnapshotRecordSet,
  meetingAppSnapshotRecords,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';
import {
  buildMeetingPlatformRolloutPlan,
} from '../packages/meeting-timeline-sdk/adapters/platform-rollout.mjs';
import {
  normalizeMeetingPlatform,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';
import {
  asArray,
  boolLabel,
  extractRecordSet,
  parseCliArgs,
  platformsFor,
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
const providerInputList = String(args.get('provider-inputs') || args.get('provider-input') || '').split(',').map((item) => item.trim()).filter(Boolean);
const domInputList = String(args.get('dom-inputs') || args.get('dom-input') || '').split(',').map((item) => item.trim()).filter(Boolean);
const jsonOutput = args.get('json') === 'true';
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const requireProductionReady = args.get('require-production-ready') !== 'false';
const failOnIncomplete = args.get('fail-on-incomplete') === 'true';
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

async function evidenceFiles(inputList, dir) {
  if (inputList.length > 0) return inputList.map((item) => resolve(item));
  return collectJsonFiles(dir);
}

function platformFromRecord(record = {}) {
  return normalizePlatformKey(record.platform ?? record.provider ?? record.adapter ?? record.snapshot?.platform);
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
    );
  }
  return rows.filter((record) => record?.schema === PLATFORM_EVENT_CAPTURE_SCHEMA || record?.body || record?.raw_body);
}

function sampleGroupsFromInput(input = {}) {
  const output = new Map();
  const groups = [
    input.samples,
    input.sample_events,
    input.sampleEvents,
    input.provider_samples,
    input.providerSamples,
  ].filter(Boolean);
  for (const group of groups) {
    if (Array.isArray(group)) {
      for (const sample of group) {
        const platform = normalizePlatformKey(sample?.platform ?? sample?.provider ?? sample?.adapter ?? input.platform);
        if (!output.has(platform)) output.set(platform, []);
        output.get(platform).push(sample);
      }
    } else if (typeof group === 'object') {
      for (const [rawPlatform, values] of Object.entries(group)) {
        const platform = normalizePlatformKey(rawPlatform);
        if (!output.has(platform)) output.set(platform, []);
        output.get(platform).push(...asArray(values));
      }
    }
  }
  return output;
}

function addMapList(map, key, values) {
  if (!map.has(key)) map.set(key, []);
  map.get(key).push(...values);
}

async function readProviderEvidence(file) {
  const text = await readFile(file, 'utf8');
  const input = JSON.parse(text);
  const records = providerRecordsFromInput(input);
  const sampleGroups = sampleGroupsFromInput(input);
  const platforms = unique([
    ...asArray(input.platforms),
    ...asArray(input.platform),
    ...records.map((record) => record.platform ?? record.provider ?? record.adapter),
    ...sampleGroups.keys(),
  ]).map((item) => normalizePlatformKey(item));
  const stats = await stat(file);
  return {
    file,
    input,
    records,
    sampleGroups,
    platforms,
    bytes: Buffer.byteLength(text, 'utf8'),
    modified_at_ms: Math.round(stats.mtimeMs),
  };
}

async function readDomEvidence(file) {
  const text = await readFile(file, 'utf8');
  const input = JSON.parse(text);
  const recordSet = extractRecordSet(input);
  const platforms = platformsFor(recordSet, input).map((item) => normalizePlatformKey(item));
  const records = meetingAppSnapshotRecords(recordSet);
  const stats = await stat(file);
  return {
    file,
    input,
    recordSet,
    records,
    platforms,
    bytes: Buffer.byteLength(text, 'utf8'),
    modified_at_ms: Math.round(stats.mtimeMs),
  };
}

function aggregateProvider(evaluations = []) {
  const recordsByPlatform = new Map();
  const samplesByPlatform = new Map();
  const filesByPlatform = new Map();
  for (const evidence of evaluations) {
    for (const record of evidence.records) {
      const platform = platformFromRecord(record);
      addMapList(recordsByPlatform, platform, [record]);
      addMapList(filesByPlatform, platform, [evidence.file]);
    }
    for (const [platform, samples] of evidence.sampleGroups.entries()) {
      addMapList(samplesByPlatform, platform, samples);
      addMapList(filesByPlatform, platform, [evidence.file]);
    }
  }
  return { recordsByPlatform, samplesByPlatform, filesByPlatform };
}

function aggregateDom(evaluations = []) {
  const recordsByPlatform = new Map();
  const filesByPlatform = new Map();
  for (const evidence of evaluations) {
    for (const record of evidence.records) {
      const platform = platformFromRecord(record);
      addMapList(recordsByPlatform, platform, [record]);
      addMapList(filesByPlatform, platform, [evidence.file]);
    }
  }
  return { recordsByPlatform, filesByPlatform };
}

function summarizePlan(plan = {}, providerFiles = [], domFiles = []) {
  return {
    platform: plan.platform,
    status: plan.status,
    production_ready: plan.production_ready,
    ready_for_realtime_annotations: plan.ready_for_realtime_annotations,
    recommended_mode: plan.recommended_mode,
    provider: {
      production_ready: plan.provider_events?.production_ready ?? false,
      setup_ready: plan.provider_events?.setup_ready ?? false,
      evidence_level: plan.provider_events?.evidence_level ?? 'none',
      evidence_count: plan.provider_events?.evidence_count ?? 0,
      missing_required_coverage: plan.provider_events?.missing_required_coverage ?? [],
      source_files: unique(providerFiles),
    },
    local_dom: plan.local_observer ? {
      production_ready: plan.local_observer.production_ready,
      runtime_ready: plan.local_observer.runtime_ready,
      evidence_level: plan.local_observer.evidence_level,
      evidence_count: plan.local_observer.evidence_count,
      missing_required_coverage: plan.local_observer.missing_required_coverage ?? [],
      source_files: unique(domFiles),
    } : undefined,
    next_actions: plan.next_actions ?? [],
  };
}

async function loadEvidence() {
  const providerFiles = await evidenceFiles(providerInputList, providerDir);
  const domFiles = await evidenceFiles(domInputList, domDir);
  const providerEvaluations = [];
  const domEvaluations = [];
  const errors = [];
  for (const file of providerFiles) {
    try {
      providerEvaluations.push(await readProviderEvidence(file));
    } catch (error) {
      errors.push({ kind: 'provider', file, error: String(error?.message ?? error) });
    }
  }
  for (const file of domFiles) {
    try {
      domEvaluations.push(await readDomEvidence(file));
    } catch (error) {
      errors.push({ kind: 'dom', file, error: String(error?.message ?? error) });
    }
  }
  return { providerFiles, domFiles, providerEvaluations, domEvaluations, errors };
}

function buildMatrix(loaded) {
  const provider = aggregateProvider(loaded.providerEvaluations);
  const dom = aggregateDom(loaded.domEvaluations);
  const rows = [];
  const plans = [];
  for (const rawPlatform of requiredPlatforms) {
    const platform = normalizePlatformKey(rawPlatform);
    const providerRecords = provider.recordsByPlatform.get(platform) ?? [];
    const providerSamples = provider.samplesByPlatform.get(platform) ?? [];
    const domRecords = dom.recordsByPlatform.get(platform) ?? [];
    const meetingAppRecordSet = buildMeetingAppSnapshotRecordSet(domRecords, {
      id: `${platform}-rollout-dom-records`,
      source: 'meeting_platform_rollout_matrix',
    });
    const plan = buildMeetingPlatformRolloutPlan(platform, {
      baseUrl,
      env: process.env,
      providerRecords,
      providerSamples: providerSamples.length > 0 ? { [platform]: providerSamples } : undefined,
      meetingAppRecordSet,
      allowProviderFixtureEvidence: false,
      allowMeetingAppFixtureEvidence: false,
    });
    plans.push(plan);
    rows.push(summarizePlan(
      plan,
      provider.filesByPlatform.get(platform) ?? [],
      dom.filesByPlatform.get(platform) ?? [],
    ));
  }
  const ok = rows.length > 0 && (requireProductionReady
    ? rows.every((row) => row.production_ready)
    : rows.every((row) => row.ready_for_realtime_annotations));
  return {
    type: 'meeting_platform_rollout_matrix',
    ok,
    production_ready: rows.length > 0 && rows.every((row) => row.production_ready),
    realtime_ready: rows.length > 0 && rows.every((row) => row.ready_for_realtime_annotations),
    base_url: baseUrl,
    provider_evidence_dir: providerInputList.length > 0 ? null : providerDir,
    dom_evidence_dir: domInputList.length > 0 ? null : domDir,
    provider_input_files: loaded.providerFiles,
    dom_input_files: loaded.domFiles,
    provider_file_count: loaded.providerFiles.length,
    dom_file_count: loaded.domFiles.length,
    provider_evaluated_file_count: loaded.providerEvaluations.length,
    dom_evaluated_file_count: loaded.domEvaluations.length,
    required_platforms: requiredPlatforms,
    require_production_ready: requireProductionReady,
    production_ready_count: rows.filter((row) => row.production_ready).length,
    realtime_ready_count: rows.filter((row) => row.ready_for_realtime_annotations).length,
    pending_platforms: rows.filter((row) => !row.production_ready).map((row) => row.platform),
    rows,
    plans,
    next_actions: unique(rows.flatMap((row) => row.next_actions)),
    errors: loaded.errors,
  };
}

try {
  const loaded = await loadEvidence();
  const summary = buildMatrix(loaded);
  if (loaded.errors.length > 0 && loaded.providerEvaluations.length === 0 && loaded.domEvaluations.length === 0) {
    summary.ok = false;
  }
  if (reportFile) {
    const target = resolve(reportFile);
    await mkdir(dirname(target), { recursive: true });
    await writeFile(target, `${JSON.stringify(summary, null, 2)}\n`, 'utf8');
  }
  if (jsonOutput) {
    console.log(JSON.stringify(summary, null, 2));
  } else {
    console.log(`meeting_platform_rollout_matrix | ok=${boolLabel(summary.ok)} | production_ready=${boolLabel(summary.production_ready)} | realtime_ready=${boolLabel(summary.realtime_ready)} | provider_files=${summary.provider_evaluated_file_count}/${summary.provider_file_count} | dom_files=${summary.dom_evaluated_file_count}/${summary.dom_file_count}`);
    for (const row of summary.rows) {
      console.log(`${row.platform}: status=${row.status} production_ready=${boolLabel(row.production_ready)} realtime_ready=${boolLabel(row.ready_for_realtime_annotations)} provider=${row.provider.evidence_level}/${row.provider.evidence_count} dom=${row.local_dom?.evidence_level ?? 'none'}/${row.local_dom?.evidence_count ?? 0}`);
    }
    if (summary.next_actions.length > 0) console.log(`next_actions=${summary.next_actions.join(',')}`);
  }
  if (!summary.ok && failOnIncomplete) process.exitCode = 2;
} catch (error) {
  const payload = {
    type: 'meeting_platform_rollout_matrix',
    ok: false,
    provider_evidence_dir: providerDir,
    dom_evidence_dir: domDir,
    error: String(error?.message ?? error),
  };
  if (jsonOutput) console.error(JSON.stringify(payload, null, 2));
  else console.error(`meeting_platform_rollout_matrix | ok=no | error=${payload.error}`);
  process.exitCode = 2;
}
