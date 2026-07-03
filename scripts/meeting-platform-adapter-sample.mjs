#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  runMeetingPlatformAdapterSampleMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-sample.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const outDir = resolve(String(args.get('out-dir') || args.get('outDir') || args.get('sample-dir') || args.get('sampleDir') || 'data/meeting-platform-adapter-samples'));
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const writeSamples = args.get('write-samples') !== 'false';
const failOnRejected = args.get('fail-on-rejected') === 'true' || args.get('fail-on-sample') === 'true';
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function sampleFile(platform) {
  return resolve(outDir, `${platform}.json`);
}

function summarizeSample(sample = {}) {
  return {
    platform: sample.platform,
    accepted: sample.accepted,
    readiness_status: sample.readiness?.status,
    production_ready: sample.readiness?.production_ready === true,
    ready_for_realtime_annotations: sample.readiness?.ready_for_realtime_annotations === true,
    provider_event_count: sample.provider_event_count,
    timeline_call_count: sample.timeline_call_count,
    timeline_method_counts: sample.timeline_method_counts,
    sample_file: sampleFile(sample.platform),
  };
}

async function buildReport() {
  const matrix = await runMeetingPlatformAdapterSampleMatrix({
    baseUrl,
    env: process.env,
    platforms: requiredPlatforms,
    includeEvidencePackage: args.get('include-evidence-package') === 'true',
  });
  const writtenFiles = [];
  if (writeSamples) {
    for (const sample of matrix.samples) {
      const file = sampleFile(sample.platform);
      await writeJson(file, sample);
      writtenFiles.push(file);
    }
  }
  const rows = matrix.samples.map((sample) => summarizeSample(sample));
  const ok = matrix.platform_count > 0 && (!failOnRejected || matrix.rejected_count === 0);
  return {
    type: 'meeting_platform_adapter_sample_report',
    ok,
    requirement: failOnRejected ? 'all_adapter_samples_accepted' : 'adapter_samples_generated',
    base_url: baseUrl,
    out_dir: writeSamples ? outDir : undefined,
    write_samples: writeSamples,
    platform_count: matrix.platform_count,
    accepted_count: matrix.accepted_count,
    rejected_count: matrix.rejected_count,
    production_ready_count: matrix.production_ready_count,
    realtime_ready_count: matrix.realtime_ready_count,
    required_platforms: requiredPlatforms,
    written_files: writtenFiles,
    rows,
    matrix,
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_adapter_sample_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | accepted=${report.accepted_count}/${report.platform_count} | production_ready=${report.production_ready_count} | realtime_ready=${report.realtime_ready_count} | written=${report.written_files.length}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: accepted=${boolLabel(row.accepted)} readiness=${row.readiness_status} provider_events=${row.provider_event_count} calls=${row.timeline_call_count} sample=${row.sample_file}`);
    }
  }
  if (!report.ok && failOnRejected) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
