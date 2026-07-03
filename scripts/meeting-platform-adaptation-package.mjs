#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdaptationPackageMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adaptation-package.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const outDir = resolve(String(args.get('out-dir') || args.get('outDir') || 'data/meeting-platform-adaptation-packages'));
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const includePackages = args.get('include-packages') === 'true';
const writePackages = args.get('write-packages') !== 'false';
const failOnContractRejected = args.get('fail-on-rejected') === 'true';
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function packageFile(platform) {
  return resolve(outDir, `${platform}.json`);
}

function rowWithFile(row = {}) {
  return {
    ...row,
    package_file: packageFile(row.platform),
  };
}

async function buildReport() {
  const matrix = buildMeetingPlatformAdaptationPackageMatrix({
    baseUrl,
    env: process.env,
    platforms: requiredPlatforms,
  });
  const writtenFiles = [];
  if (writePackages) {
    for (const item of matrix.packages) {
      const file = packageFile(item.platform);
      await writeJson(file, item);
      writtenFiles.push(file);
    }
  }
  const rejected = matrix.packages.filter((item) => item.readiness.sdk_wiring_ready !== true);
  return {
    type: 'meeting_platform_adaptation_package_report',
    ok: matrix.platform_count > 0 && (!failOnContractRejected || rejected.length === 0),
    requirement: failOnContractRejected
      ? 'all_adaptation_packages_have_accepted_sdk_contracts'
      : 'adaptation_packages_exported',
    base_url: baseUrl,
    out_dir: writePackages ? outDir : undefined,
    write_packages: writePackages,
    platform_count: matrix.platform_count,
    sdk_wiring_ready_count: matrix.sdk_wiring_ready_count,
    browser_observer_count: matrix.browser_observer_count,
    provider_observer_count: matrix.provider_observer_count,
    production_ready_count: matrix.production_ready_count,
    realtime_ready_count: matrix.realtime_ready_count,
    required_platforms: requiredPlatforms,
    written_files: writtenFiles,
    rows: matrix.rows.map((row) => rowWithFile(row)),
    matrix: includePackages ? matrix : {
      ...matrix,
      packages: undefined,
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
    console.log(`meeting_platform_adaptation_package_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | sdk_ready=${report.sdk_wiring_ready_count} | browser_observers=${report.browser_observer_count} | providers=${report.provider_observer_count} | production_ready=${report.production_ready_count} | realtime_ready=${report.realtime_ready_count} | written=${report.written_files.length}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: mode=${row.mode} sdk_ready=${boolLabel(row.sdk_wiring_ready)} browser_matches=${row.browser_match_count} provider=${row.provider_transport ?? 'none'} start_events=${row.provider_start_event_count} end_events=${row.provider_end_event_count} next=${row.first_next_action ?? 'none'} package=${row.package_file}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
  }
  if (!report.ok && failOnContractRejected) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
