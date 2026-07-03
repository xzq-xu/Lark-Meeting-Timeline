#!/usr/bin/env node

import { mkdir, readdir, readFile, stat, writeFile } from 'node:fs/promises';
import { basename, dirname, resolve } from 'node:path';
import {
  meetingAppSnapshotRecords,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';
import {
  buildMeetingPlatformFieldEvidenceBundle,
} from '../packages/meeting-timeline-sdk/adapters/platform-field-capture.mjs';
import {
  normalizeMeetingPlatform,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';
import {
  asArray,
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const inputDir = resolve(String(args.get('input-dir') || args.get('dir') || 'data/meeting-platform-field-evidence'));
const inputList = String(args.get('inputs') || args.get('input') || '').split(',').map((item) => item.trim()).filter(Boolean);
const bundleDir = resolve(String(args.get('bundle-dir') || args.get('bundleDir') || 'data/meeting-platform-field-evidence-bundles'));
const packageDir = resolve(String(args.get('package-dir') || args.get('packageDir') || args.get('out-dir') || args.get('outDir') || 'data/meeting-platform-evidence-packages'));
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const failOnIncomplete = args.get('fail-on-incomplete') === 'true' || args.get('fail-on-missing') === 'true';
const requireProductionReady = args.get('require-production-ready') !== 'false';
const requireCorrelation = args.get('require-correlation') !== 'false';
const writeBundles = args.get('write-bundles') !== 'false';
const writePackages = args.get('write-packages') !== 'false';
const writeEmpty = args.get('write-empty') === 'true';
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || '');
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function normalizePlatformKey(value) {
  try {
    return normalizeMeetingPlatform(value);
  } catch {
    return String(value ?? '').replace(/\.json$/i, '').replaceAll('-', '_');
  }
}

function matchesPlatform(value, platform) {
  if (!value) return false;
  return normalizePlatformKey(value) === platform;
}

function platformFromInput(input = {}, file = '') {
  return firstNonEmpty(
    input.platform,
    input.provider,
    input.adapter,
    input.rollout_plan?.platform,
    input.evidence_package?.platform,
    input.evidencePackage?.platform,
    file ? basename(file).replace(/\.json$/i, '') : undefined,
  );
}

function platformFromRecord(record = {}) {
  return firstNonEmpty(
    record.platform,
    record.provider,
    record.adapter,
    record.snapshot?.platform,
    record.snapshot?.provider,
    record.payload?.platform,
    record.raw?.event?.type,
  );
}

function recordsForPlatform(platform, records = []) {
  return asArray(records).filter((record) => {
    const recordPlatform = platformFromRecord(record);
    return !recordPlatform || matchesPlatform(recordPlatform, platform);
  });
}

function samplesFromCollection(platform, collection) {
  if (!collection) return [];
  if (Array.isArray(collection)) return recordsForPlatform(platform, collection);
  if (typeof collection === 'object') {
    const samples = [];
    for (const [key, value] of Object.entries(collection)) {
      if (matchesPlatform(key, platform)) samples.push(...asArray(value));
    }
    return samples;
  }
  return [];
}

function snapshotRecordsFrom(platform, input) {
  if (!input) return [];
  if (Array.isArray(input)) return recordsForPlatform(platform, input);
  if (input.schema === 'meeting_app_snapshot_record_set' || input.records) {
    return recordsForPlatform(platform, meetingAppSnapshotRecords(input));
  }
  return [];
}

function candidateInputsForPlatform(platform, source = {}) {
  const input = source.input;
  if (!input) return [];
  if (Array.isArray(input)) {
    return input.flatMap((item) => candidateInputsForPlatform(platform, { ...source, input: item }));
  }
  if (input.schema === 'meeting_platform_field_evidence_bundle') {
    return matchesPlatform(input.platform, platform) ? [{
      file: source.file,
      input: input.evidence_package ?? input,
      source_kind: 'field_evidence_bundle',
    }] : [];
  }
  if (input.schema === 'meeting_platform_evidence_package') {
    return matchesPlatform(input.platform, platform) ? [{
      file: source.file,
      input,
      source_kind: 'evidence_package',
    }] : [];
  }
  const explicitPlatform = platformFromInput(input, source.file);
  const hasEvidenceFields = Boolean(
    input.providerRecords
      || input.provider_records
      || input.providerCaptureRecords
      || input.provider_capture_records
      || input.captureRecords
      || input.capture_records
      || input.meetingAppRecordSet
      || input.meeting_app_record_set
      || input.meetingAppRecords
      || input.meeting_app_records
      || input.snapshotRecords
      || input.snapshot_records
      || input.providerSamples
      || input.provider_samples
  );
  if (hasEvidenceFields && matchesPlatform(explicitPlatform, platform)) {
    return [{ file: source.file, input, source_kind: 'raw_field_evidence' }];
  }
  if (typeof input === 'object') {
    const items = [];
    for (const [key, value] of Object.entries(input)) {
      if (matchesPlatform(key, platform)) {
        items.push({
          file: source.file,
          input: {
            ...(value && typeof value === 'object' && !Array.isArray(value) ? value : { records: value }),
            platform,
          },
          source_kind: 'platform_map',
        });
      }
    }
    return items;
  }
  return [];
}

function mergePlatformEvidence(platform, candidates = []) {
  const providerRecords = [];
  const providerSamples = [];
  const meetingAppRecords = [];
  const metadata = {
    source_files: unique(candidates.map((candidate) => candidate.file)),
    source_kinds: unique(candidates.map((candidate) => candidate.source_kind)),
  };
  for (const candidate of candidates) {
    const input = candidate.input ?? {};
    providerRecords.push(...recordsForPlatform(platform, [
      ...asArray(input.providerRecords),
      ...asArray(input.provider_records),
      ...asArray(input.providerCaptureRecords),
      ...asArray(input.provider_capture_records),
      ...asArray(input.captureRecords),
      ...asArray(input.capture_records),
      ...asArray(input.provider_record_set?.records),
    ]));
    providerSamples.push(
      ...samplesFromCollection(platform, input.providerSamples),
      ...samplesFromCollection(platform, input.provider_samples),
      ...samplesFromCollection(platform, input.provider_samples_by_platform),
      ...samplesFromCollection(platform, input.sampleEvents),
      ...samplesFromCollection(platform, input.sample_events),
      ...samplesFromCollection(platform, input.samples),
    );
    meetingAppRecords.push(
      ...snapshotRecordsFrom(platform, input.meetingAppRecordSet),
      ...snapshotRecordsFrom(platform, input.meeting_app_record_set),
      ...snapshotRecordsFrom(platform, input.recordSet),
      ...snapshotRecordsFrom(platform, input.record_set),
      ...snapshotRecordsFrom(platform, input.meetingAppRecords),
      ...snapshotRecordsFrom(platform, input.meeting_app_records),
      ...snapshotRecordsFrom(platform, input.meetingAppSnapshotRecords),
      ...snapshotRecordsFrom(platform, input.meeting_app_snapshot_records),
      ...snapshotRecordsFrom(platform, input.snapshotRecords),
      ...snapshotRecordsFrom(platform, input.snapshot_records),
      ...snapshotRecordsFrom(platform, input.domSnapshots),
      ...snapshotRecordsFrom(platform, input.dom_snapshots),
    );
  }
  return {
    platform,
    providerRecords,
    providerSamples,
    meetingAppRecords,
    metadata,
  };
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

async function readSource(file) {
  const text = await readFile(file, 'utf8');
  const stats = await stat(file);
  return {
    file,
    input: JSON.parse(text),
    bytes: Buffer.byteLength(text, 'utf8'),
    modified_at_ms: Math.round(stats.mtimeMs),
  };
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function outputPath(root, platform) {
  return resolve(root, `${platform}.json`);
}

function summarizeRow(bundle = {}, input = {}) {
  const packagePath = outputPath(packageDir, bundle.platform);
  const bundlePath = outputPath(bundleDir, bundle.platform);
  return {
    platform: bundle.platform,
    status: bundle.status,
    production_ready: bundle.production_ready,
    ready_for_realtime_annotations: bundle.ready_for_realtime_annotations,
    verified: bundle.verification?.passed,
    package_id: bundle.package_id,
    provider_record_count: bundle.evidence_summary?.provider_record_count ?? 0,
    provider_sample_count: bundle.evidence_summary?.provider_sample_count ?? 0,
    meeting_app_record_count: bundle.evidence_summary?.meeting_app_record_count ?? 0,
    correlation_status: bundle.verification?.correlation_status,
    correlation_passed: bundle.verification?.correlation_passed,
    missing_items: bundle.field_capture_plan?.missing_items ?? [],
    source_files: input.metadata?.source_files ?? [],
    source_kinds: input.metadata?.source_kinds ?? [],
    package_file: packagePath,
    bundle_file: bundlePath,
    next_actions: bundle.next_actions ?? [],
  };
}

async function buildReport() {
  const files = await evidenceFiles();
  const sources = [];
  const errors = [];
  for (const file of files) {
    try {
      sources.push(await readSource(file));
    } catch (error) {
      errors.push({ file, error: String(error?.message ?? error) });
    }
  }
  const platforms = requiredPlatforms.map((platform) => normalizePlatformKey(platform));
  const rows = [];
  const bundles = [];
  const written = [];
  for (const platform of platforms) {
    const candidates = sources.flatMap((source) => candidateInputsForPlatform(platform, source));
    const input = mergePlatformEvidence(platform, candidates);
    const bundle = buildMeetingPlatformFieldEvidenceBundle(platform, input, {
      baseUrl: baseUrl || undefined,
      env: process.env,
      requireProductionReady,
      requireCorrelation,
    });
    const sourceFileCount = input.metadata.source_files.length;
    const shouldWrite = writeEmpty || sourceFileCount > 0;
    if (shouldWrite && writeBundles) {
      const bundleFile = outputPath(bundleDir, platform);
      await writeJson(bundleFile, bundle);
      written.push(bundleFile);
    }
    if (shouldWrite && writePackages) {
      const packageFile = outputPath(packageDir, platform);
      await writeJson(packageFile, bundle.evidence_package);
      written.push(packageFile);
    }
    rows.push(summarizeRow(bundle, input));
    bundles.push(bundle);
  }
  const passedCount = rows.filter((row) => row.verified).length;
  const ok = rows.length > 0 && passedCount === rows.length && errors.length === 0;
  return {
    type: 'meeting_platform_field_evidence_report',
    ok,
    requirement: requireProductionReady ? 'production_ready' : 'ready_for_realtime_annotations',
    require_production_ready: requireProductionReady,
    require_correlation: requireCorrelation,
    input_dir: inputList.length > 0 ? null : inputDir,
    input_files: files,
    file_count: files.length,
    evaluated_file_count: sources.length,
    required_platforms: platforms,
    bundle_dir: writeBundles ? bundleDir : undefined,
    package_dir: writePackages ? packageDir : undefined,
    written_files: written,
    platform_count: rows.length,
    passed_count: passedCount,
    production_ready_count: rows.filter((row) => row.production_ready).length,
    realtime_ready_count: rows.filter((row) => row.ready_for_realtime_annotations).length,
    missing_item_count: rows.reduce((total, row) => total + row.missing_items.length, 0),
    rows,
    bundles,
    next_actions: unique(rows.flatMap((row) => row.next_actions)),
    errors,
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_field_evidence_report | ok=${boolLabel(report.ok)} | requirement=${report.requirement} | passed=${report.passed_count}/${report.platform_count} | production_ready=${report.production_ready_count} | realtime_ready=${report.realtime_ready_count} | missing_items=${report.missing_item_count} | files=${report.evaluated_file_count}/${report.file_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: status=${row.status} verified=${boolLabel(row.verified)} production_ready=${boolLabel(row.production_ready)} realtime_ready=${boolLabel(row.ready_for_realtime_annotations)} provider_records=${row.provider_record_count} dom_records=${row.meeting_app_record_count} sources=${row.source_files.length}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
    for (const error of report.errors) console.error(`error ${error.file}: ${error.error}`);
  }
  if (!report.ok && failOnIncomplete) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
