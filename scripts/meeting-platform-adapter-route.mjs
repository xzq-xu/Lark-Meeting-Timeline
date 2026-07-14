#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdapterRouteMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-route.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const outDir = resolve(String(args.get('out-dir') || args.get('outDir') || 'data/meeting-platform-adapter-routes'));
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const includeRoutes = args.get('include-routes') === 'true' || args.get('includeRoutes') === 'true';
const writeRoutes = args.get('write-routes') !== 'false';
const failOnBlocking = args.get('fail-on-blocking') === 'true';
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function routeFile(platform) {
  return resolve(outDir, `${platform}.json`);
}

function rowWithFile(row = {}) {
  return {
    ...row,
    route_file: routeFile(row.platform),
  };
}

function blockingRows(matrix = {}) {
  return (matrix.rows ?? []).filter((row) => (
    row.provider_blocks_realtime === true || row.transcript_blocks_realtime === true
  ));
}

async function buildReport() {
  const matrix = buildMeetingPlatformAdapterRouteMatrix({
    baseUrl,
    env: process.env,
    platforms: requiredPlatforms,
  });
  const writtenFiles = [];
  if (writeRoutes) {
    for (const route of matrix.routes) {
      const file = routeFile(route.platform);
      await writeJson(file, route);
      writtenFiles.push(file);
    }
  }
  const blockers = blockingRows(matrix);
  return {
    type: 'meeting_platform_adapter_route_report',
    ok: matrix.platform_count > 0 && blockers.length === 0,
    requirement: 'adapter_routes_keep_local_observer_first_and_provider_nonblocking',
    base_url: baseUrl,
    out_dir: writeRoutes ? outDir : undefined,
    write_routes: writeRoutes,
    platform_count: matrix.platform_count,
    local_observer_first_count: matrix.local_observer_first_count,
    provider_non_blocking_count: matrix.provider_non_blocking_count,
    transcript_non_blocking_count: matrix.transcript_non_blocking_count,
    blocking_count: blockers.length,
    blocking_platforms: blockers.map((row) => row.platform),
    required_platforms: requiredPlatforms,
    written_files: writtenFiles,
    rows: matrix.rows.map((row) => rowWithFile(row)),
    matrix: includeRoutes ? matrix : {
      ...matrix,
      routes: undefined,
    },
    next_actions: unique(matrix.routes.flatMap((route) => route.next_actions ?? [])),
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_adapter_route_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | local_first=${report.local_observer_first_count} | provider_nonblocking=${report.provider_non_blocking_count} | transcript_nonblocking=${report.transcript_non_blocking_count} | blocking=${report.blocking_count} | written=${report.written_files.length}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: mode=${row.recommended_mode} first=${row.first_route} matches=${row.browser_match_count} provider=${row.provider_transport ?? 'none'} production_gate=${row.production_gate} route=${row.route_file}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
  }
  if (!report.ok && failOnBlocking) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
