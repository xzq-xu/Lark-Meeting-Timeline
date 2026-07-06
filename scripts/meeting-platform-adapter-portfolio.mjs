#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdapterPortfolio,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-portfolio.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const outDir = resolve(String(args.get('out-dir') || args.get('outDir') || 'data/meeting-platform-adapter-portfolio'));
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const includeArtifacts = args.get('include-artifacts') === 'true';
const writeItems = args.get('write-items') !== 'false';
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()).filter(Boolean));

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function itemFile(platform) {
  return resolve(outDir, `${platform}.json`);
}

async function buildReport() {
  const portfolio = buildMeetingPlatformAdapterPortfolio({
    baseUrl,
    env: process.env,
    platforms: requiredPlatforms,
    includeArtifacts,
  });
  const writtenFiles = [];
  if (writeItems) {
    for (const item of portfolio.items) {
      const file = itemFile(item.platform);
      await writeJson(file, item);
      writtenFiles.push(file);
    }
  }
  return {
    type: 'meeting_platform_adapter_portfolio_report',
    ok: portfolio.platform_count > 0,
    requirement: 'meeting_platform_adapter_portfolio_available_for_meeting_software_selection_and_handoff',
    base_url: baseUrl,
    out_dir: writeItems ? outDir : undefined,
    write_items: writeItems,
    include_artifacts: includeArtifacts,
    platform_count: portfolio.platform_count,
    built_in_count: portfolio.built_in_count,
    external_authoring_count: portfolio.external_authoring_count,
    browser_surface_ready_count: portfolio.browser_surface_ready_count,
    provider_reconcile_count: portfolio.provider_reconcile_count,
    implementation_ready_count: portfolio.implementation_ready_count,
    pilot_ready_count: portfolio.pilot_ready_count,
    production_ready_count: portfolio.production_ready_count,
    required_platforms: requiredPlatforms,
    written_files: writtenFiles,
    rows: portfolio.rows.map((row) => ({
      ...row,
      item_file: itemFile(row.platform),
    })),
    portfolio: includeArtifacts ? portfolio : {
      ...portfolio,
      items: portfolio.items.map((item) => ({
        ...item,
        artifacts: undefined,
      })),
    },
    next_actions: portfolio.next_actions ?? [],
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_adapter_portfolio_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | builtin=${report.built_in_count} | external=${report.external_authoring_count} | browser_ready=${report.browser_surface_ready_count} | provider=${report.provider_reconcile_count} | pilot=${report.pilot_ready_count} | written=${report.written_files.length}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: status=${row.adapter_status} surface=${row.recommended_first_surface} provider=${row.provider_path} docs=${row.official_doc_count} pilot=${boolLabel(row.pilot_ready)} next=${row.next_action}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
  }
  if (!report.ok) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
