#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdaptationStrategyMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-strategy.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const includeStrategies = args.get('include-strategies') === 'true' || args.get('includeStrategies') === 'true';
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

function strategyBlockingRows(matrix = {}) {
  return (matrix.rows ?? []).filter((row) => (
    row.provider_blocks_realtime === true || row.transcript_blocks_realtime === true
  ));
}

function buildReport() {
  const matrix = buildMeetingPlatformAdaptationStrategyMatrix({
    baseUrl,
    env: process.env,
    platforms: requiredPlatforms,
  });
  const blockingRows = strategyBlockingRows(matrix);
  return {
    type: 'meeting_platform_adaptation_strategy_report',
    ok: matrix.strategy_count > 0 && blockingRows.length === 0,
    requirement: 'selected_platforms_keep_provider_and_transcript_nonblocking_for_realtime_annotations',
    base_url: baseUrl,
    required_platforms: requiredPlatforms,
    strategy_count: matrix.strategy_count,
    production_ready_count: matrix.production_ready_count,
    realtime_ready_count: matrix.realtime_ready_count,
    local_first_count: matrix.local_first_count,
    non_blocking_provider_count: matrix.non_blocking_provider_count,
    blocking_count: blockingRows.length,
    blocking_platforms: blockingRows.map((row) => row.platform),
    rows: matrix.rows,
    matrix: includeStrategies ? matrix : {
      ...matrix,
      strategies: undefined,
    },
    next_actions: unique(matrix.strategies.flatMap((item) => item.next_actions ?? [])),
  };
}

try {
  const report = buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_adaptation_strategy_report | ok=${boolLabel(report.ok)} | platforms=${report.strategy_count} | realtime_ready=${report.realtime_ready_count} | production_ready=${report.production_ready_count} | local_first=${report.local_first_count} | provider_nonblocking=${report.non_blocking_provider_count} | blocking=${report.blocking_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: primary=${row.primary_axis_source} provider_blocks=${boolLabel(row.provider_blocks_realtime)} transcript_blocks=${boolLabel(row.transcript_blocks_realtime)} speaker=${row.speaker_realtime_primary} recommendation=${row.recommendation}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
  }
  if (!report.ok && failOnBlocking) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
