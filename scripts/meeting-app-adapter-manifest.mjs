#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingAppAdapterManifestMatrix,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-adapter-manifest.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const outDir = String(args.get('out-dir') || args.get('outDir') || '');
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const includeManifests = args.get('include-manifests') === 'true';
const failOnRejected = args.get('fail-on-rejected') === 'true' || args.get('fail-on-failed') === 'true';
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()).filter(Boolean));

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function manifestFile(platform) {
  return resolve(outDir, `${platform}.json`);
}

const matrix = buildMeetingAppAdapterManifestMatrix({
  baseUrl,
  env: process.env,
  platforms: requiredPlatforms,
});

const writtenFiles = [];
if (outDir) {
  for (const manifest of matrix.manifests) {
    const file = manifestFile(manifest.platform);
    await writeJson(file, manifest);
    writtenFiles.push(file);
  }
}

const report = {
  type: 'meeting_app_adapter_manifest_report',
  ok: matrix.accepted === true,
  base_url: baseUrl,
  out_dir: outDir || undefined,
  platform_count: matrix.platform_count,
  accepted_count: matrix.accepted_count,
  extension_match_count: matrix.extension_match_count,
  mutation_observer_count: matrix.mutation_observer_count,
  candidate_observer_count: matrix.candidate_observer_count,
  participant_selector_ready_count: matrix.participant_selector_ready_count,
  required_platforms: requiredPlatforms,
  written_files: writtenFiles,
  rows: matrix.rows.map((row) => ({
    ...row,
    manifest_file: outDir ? manifestFile(row.platform) : undefined,
  })),
  matrix: includeManifests ? matrix : {
    ...matrix,
    manifests: undefined,
  },
  next_actions: matrix.next_actions,
};

if (reportFile) await writeJson(resolve(reportFile), report);

if (jsonOutput) {
  console.log(JSON.stringify(report, null, 2));
} else {
  console.log(`meeting_app_adapter_manifest_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | accepted=${report.accepted_count} | matches=${report.extension_match_count} | mutation=${report.mutation_observer_count} | candidates=${report.candidate_observer_count} | participant_selectors=${report.participant_selector_ready_count} | written=${report.written_files.length}`);
  for (const row of report.rows) {
    console.log(`${row.platform}: accepted=${boolLabel(row.accepted)} surface=${row.surface} matches=${row.extension_match_count} mutation_selectors=${row.mutation_track_selector_count} participant_selectors=${row.participant_selector_count} observer=${boolLabel(row.observer_sdk_ready)} runtime=${boolLabel(row.runtime_adapter_accepted)} next=${row.first_next_action ?? 'none'} manifest=${row.manifest_file ?? 'n/a'}`);
  }
  if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
}

if (!report.ok && failOnRejected) process.exitCode = 2;
