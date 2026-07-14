#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformFieldCaptureManifestMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-field-capture.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || '');
const evidenceDir = String(args.get('evidence-dir') || args.get('evidenceDir') || 'data');
const outDir = resolve(String(args.get('out-dir') || args.get('outDir') || args.get('manifest-dir') || args.get('manifestDir') || 'data/meeting-platform-field-manifests'));
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const writeManifests = args.get('write-manifests') !== 'false';
const failOnIncomplete = args.get('fail-on-incomplete') === 'true' || args.get('fail-on-missing') === 'true';
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function manifestFile(platform) {
  return resolve(outDir, `${platform}.json`);
}

function summarizeManifest(manifest = {}) {
  return {
    platform: manifest.platform,
    display_name: manifest.display_name,
    status: manifest.status,
    production_ready: manifest.production_ready,
    ready_for_realtime_annotations: manifest.ready_for_realtime_annotations,
    input_file: manifest.file_contract?.files?.field_evidence_input,
    bundle_file: manifest.file_contract?.files?.field_evidence_bundle,
    evidence_package_file: manifest.file_contract?.files?.evidence_package,
    manifest_file: manifestFile(manifest.platform),
    accepted_input_schemas: (manifest.input_contract?.accepted_inputs ?? []).map((item) => item.schema),
    required_provider_coverage: manifest.acceptance?.required_provider_coverage ?? [],
    required_local_snapshots: manifest.acceptance?.required_local_snapshots ?? [],
    missing_items: manifest.acceptance?.current_missing_items ?? [],
    build_field_evidence_command: manifest.automation?.commands?.build_field_evidence,
  };
}

async function buildReport() {
  const matrix = buildMeetingPlatformFieldCaptureManifestMatrix({
    baseUrl: baseUrl || undefined,
    env: process.env,
    evidenceDir,
    platforms: requiredPlatforms,
  });
  const writtenFiles = [];
  if (writeManifests) {
    for (const manifest of matrix.manifests) {
      const file = manifestFile(manifest.platform);
      await writeJson(file, manifest);
      writtenFiles.push(file);
    }
  }
  const rows = matrix.manifests.map((manifest) => summarizeManifest(manifest));
  const ok = matrix.platform_count > 0 && (!failOnIncomplete || matrix.missing_item_count === 0);
  return {
    type: 'meeting_platform_field_manifest_report',
    ok,
    requirement: failOnIncomplete ? 'no_missing_capture_items' : 'manifest_exported',
    base_url: baseUrl || undefined,
    evidence_dir: evidenceDir,
    out_dir: writeManifests ? outDir : undefined,
    write_manifests: writeManifests,
    platform_count: matrix.platform_count,
    production_ready_count: matrix.production_ready_count,
    realtime_ready_count: matrix.realtime_ready_count,
    missing_item_count: matrix.missing_item_count,
    required_platforms: requiredPlatforms,
    written_files: writtenFiles,
    rows,
    matrix,
    next_actions: matrix.next_actions ?? [],
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_field_manifest_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | production_ready=${report.production_ready_count} | realtime_ready=${report.realtime_ready_count} | missing_items=${report.missing_item_count} | written=${report.written_files.length}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: status=${row.status} production_ready=${boolLabel(row.production_ready)} realtime_ready=${boolLabel(row.ready_for_realtime_annotations)} missing=${row.missing_items.length} manifest=${row.manifest_file}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
  }
  if (!report.ok && failOnIncomplete) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
