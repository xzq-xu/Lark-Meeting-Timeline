#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformRuntimeBundleMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-runtime-bundle.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const outDir = resolve(String(args.get('out-dir') || args.get('outDir') || 'data/meeting-platform-runtime-bundles'));
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const includeBundles = args.get('include-bundles') === 'true';
const writeBundles = args.get('write-bundles') !== 'false';
const contentScriptJs = args.get('content-script-js') || args.get('contentScriptJs') || undefined;
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function bundleFile(platform) {
  return resolve(outDir, `${platform}.json`);
}

function rowWithFile(row = {}) {
  return {
    ...row,
    bundle_file: bundleFile(row.platform),
  };
}

async function buildReport() {
  const matrix = buildMeetingPlatformRuntimeBundleMatrix({
    baseUrl,
    env: process.env,
    platforms: requiredPlatforms,
    contentScriptJs,
  });
  const writtenFiles = [];
  if (writeBundles) {
    for (const bundle of matrix.bundles) {
      const file = bundleFile(bundle.platform);
      await writeJson(file, bundle);
      writtenFiles.push(file);
    }
  }
  return {
    type: 'meeting_platform_runtime_bundle_report',
    ok: matrix.platform_count > 0 && matrix.provider_required_for_realtime_count === 0 && matrix.transcript_blocking_count === 0,
    requirement: 'runtime_bundles_exported_for_nonblocking_realtime_annotation',
    base_url: baseUrl,
    out_dir: writeBundles ? outDir : undefined,
    write_bundles: writeBundles,
    platform_count: matrix.platform_count,
    runtime_ready_count: matrix.runtime_ready_count,
    sdk_wiring_ready_count: matrix.sdk_wiring_ready_count,
    provider_required_for_realtime_count: matrix.provider_required_for_realtime_count,
    transcript_blocking_count: matrix.transcript_blocking_count,
    required_platforms: requiredPlatforms,
    written_files: writtenFiles,
    rows: matrix.rows.map((row) => rowWithFile(row)),
    matrix: includeBundles ? matrix : {
      ...matrix,
      bundles: undefined,
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
    console.log(`meeting_platform_runtime_bundle_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | runtime_ready=${report.runtime_ready_count} | sdk_ready=${report.sdk_wiring_ready_count} | provider_blocks=${report.provider_required_for_realtime_count} | transcript_blocks=${report.transcript_blocking_count} | written=${report.written_files.length}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: runtime=${boolLabel(row.runtime_ready)} sdk=${boolLabel(row.sdk_wiring_ready)} matches=${row.browser_match_count} sample=${row.sample_interval_ms} mutation=${row.mutation_debounce_ms} provider=${row.provider_transport ?? 'none'} bundle=${row.bundle_file}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
  }
  if (!report.ok) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
