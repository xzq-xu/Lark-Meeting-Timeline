#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformRegistryAcceptanceReport,
  buildMeetingPlatformRegistryManifest,
} from '../packages/meeting-timeline-sdk/adapters/platform-registry.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const basePath = String(args.get('base-path') || args.get('basePath') || '/api/platform-events');
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const includeEntries = args.get('include-entries') !== 'false';
const platforms = unique(String(
  args.get('platforms') || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function buildReport() {
  const manifest = buildMeetingPlatformRegistryManifest({
    baseUrl,
    basePath,
    env: process.env,
    platforms,
  });
  const acceptance = buildMeetingPlatformRegistryAcceptanceReport(manifest);
  return {
    type: 'meeting_platform_registry_report',
    ok: acceptance.accepted,
    requirement: 'all_selected_platforms_have_nonblocking_registry_entries',
    base_url: baseUrl,
    base_path: basePath,
    platform_count: manifest.platform_count,
    normalizer_count: manifest.normalizer_count,
    runtime_ready_count: manifest.runtime_ready_count,
    contract_accepted_count: manifest.contract_accepted_count,
    provider_required_for_realtime_count: manifest.provider_required_for_realtime_count,
    transcript_blocking_count: manifest.transcript_blocking_count,
    rows: manifest.rows,
    acceptance: {
      ...acceptance,
      manifest: undefined,
    },
    manifest: includeEntries ? manifest : {
      ...manifest,
      entries: undefined,
    },
    next_actions: manifest.next_actions ?? [],
  };
}

try {
  const report = buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_registry_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | normalizers=${report.normalizer_count} | runtime_ready=${report.runtime_ready_count} | contracts=${report.contract_accepted_count} | provider_blocks=${report.provider_required_for_realtime_count} | transcript_blocks=${report.transcript_blocking_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: normalize=${boolLabel(row.normalize_available)} runtime=${boolLabel(row.runtime_ready)} contract=${boolLabel(row.contract_accepted)} matches=${row.browser_match_count} provider=${row.provider_transport ?? 'none'} endpoint=${row.insert_endpoint}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
  }
  if (!report.ok) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
