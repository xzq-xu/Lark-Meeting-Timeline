#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformConformanceReport,
} from '../packages/meeting-timeline-sdk/adapters/platform-conformance.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
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

const report = buildMeetingPlatformConformanceReport({
  baseUrl,
  env: process.env,
  platforms: requiredPlatforms,
});

if (reportFile) await writeJson(resolve(reportFile), report);

if (jsonOutput) {
  console.log(JSON.stringify(report, null, 2));
} else {
  console.log(`meeting_platform_conformance_report | accepted=${boolLabel(report.accepted)} | platforms=${report.platform_count} | accepted_count=${report.accepted_count} | normalizers=${report.normalizer_count} | contracts=${report.adapter_contract_accepted_count} | routes=${report.adapter_route_ready_count} | runtime=${report.runtime_ready_count} | candidates=${report.candidate_observer_count} | blocking=${report.blocking_count}`);
  for (const row of report.rows) {
    console.log(`${row.platform}: accepted=${boolLabel(row.accepted)} normalizer=${boolLabel(row.normalizer_available)} contract=${boolLabel(row.contract_accepted)} route=${boolLabel(row.adapter_route_ready)} runtime=${boolLabel(row.runtime_ready)} candidates=${boolLabel(row.candidate_observation_ready)} first_route=${row.adapter_first_route ?? 'n/a'} issues=${row.blocking_count}`);
  }
  if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
}

if (!report.accepted && failOnRejected) process.exitCode = 2;
