#!/usr/bin/env node

import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformSubscriptionHandoffMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-subscription-handoff.mjs';
import {
  normalizeMeetingPlatform,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const includeHandoffs = args.get('include-handoffs') === 'true';
const failOnNotReady = args.get('fail-on-not-ready') === 'true' || args.get('fail-on-incomplete') === 'true';
const outDir = String(args.get('out-dir') || args.get('outDir') || '');
const writeHandoffs = args.get('write-handoffs') === 'true' || Boolean(outDir);
const subscriptionFile = String(
  args.get('subscriptions-file')
    || args.get('subscription-file')
    || args.get('subscriptions')
    || args.get('subscription')
    || '',
);
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

function normalizePlatformKey(platform) {
  try {
    return normalizeMeetingPlatform(platform);
  } catch {
    return String(platform);
  }
}

async function readJson(file) {
  if (!file) return undefined;
  const text = await readFile(resolve(file), 'utf8');
  return JSON.parse(text);
}

function subscriptionsFromInput(input = {}) {
  if (!input) return {};
  const direct = input.subscriptions ?? input.provider_subscriptions ?? input.providerSubscriptions;
  if (direct && typeof direct === 'object' && !Array.isArray(direct)) {
    return Object.fromEntries(Object.entries(direct).map(([platform, value]) => [normalizePlatformKey(platform), value]));
  }
  if (Array.isArray(input)) {
    return Object.fromEntries(input.map((item) => [
      normalizePlatformKey(item.platform ?? item.provider ?? item.adapter),
      item.subscription ?? item.input ?? item,
    ]).filter(([platform]) => platform && platform !== 'undefined'));
  }
  if (input.platform || input.provider || input.adapter) {
    return {
      [normalizePlatformKey(input.platform ?? input.provider ?? input.adapter)]: input.subscription ?? input.input ?? input,
    };
  }
  return Object.fromEntries(Object.entries(input).map(([platform, value]) => [normalizePlatformKey(platform), value]));
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function handoffFile(platform) {
  return resolve(outDir || 'data/meeting-platform-subscription-handoffs', `${platform}.json`);
}

function summarizeRow(row = {}) {
  return {
    platform: row.platform,
    display_name: row.display_name,
    status: row.status,
    ready_to_create: row.ready_to_create,
    provider_ready: row.provider_ready,
    request_count: row.request_count,
    missing_env: row.missing_env ?? [],
    renewal_due: row.renewal_due,
    renewal_status: row.renewal_status,
    subscription_builders: row.subscription_builders ?? [],
    next_actions: row.next_actions ?? [],
  };
}

async function buildReport() {
  const errors = [];
  let subscriptions = {};
  try {
    subscriptions = subscriptionsFromInput(await readJson(subscriptionFile));
  } catch (error) {
    errors.push({ file: resolve(subscriptionFile), error: String(error?.message ?? error) });
  }
  const matrix = buildMeetingPlatformSubscriptionHandoffMatrix({
    baseUrl,
    env: process.env,
    platforms: requiredPlatforms,
    subscriptions,
  });
  const writtenFiles = [];
  if (writeHandoffs) {
    for (const handoff of matrix.handoffs) {
      const file = handoffFile(handoff.platform);
      await writeJson(file, handoff);
      writtenFiles.push(file);
    }
  }
  const rows = matrix.rows.map((row) => summarizeRow(row));
  const ok = matrix.platform_count > 0
    && matrix.security_blocked_count === 0
    && matrix.parameter_missing_count === 0
    && errors.length === 0;
  return {
    type: 'meeting_platform_subscription_handoff_report',
    ok,
    requirement: 'subscription_requests_or_manual_setup_ready',
    base_url: baseUrl,
    subscriptions_file: subscriptionFile ? resolve(subscriptionFile) : undefined,
    required_platforms: requiredPlatforms,
    platform_count: matrix.platform_count,
    ready_to_create_count: matrix.ready_to_create_count,
    request_count: matrix.request_count,
    security_blocked_count: matrix.security_blocked_count,
    parameter_missing_count: matrix.parameter_missing_count,
    manual_setup_count: matrix.manual_setup_count,
    renewal_due_count: matrix.renewal_due_count,
    written_files: writtenFiles,
    rows,
    matrix: includeHandoffs ? matrix : undefined,
    next_actions: unique(matrix.next_actions ?? []),
    errors,
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_subscription_handoff_report | ok=${boolLabel(report.ok)} | ready_to_create=${report.ready_to_create_count}/${report.platform_count} | requests=${report.request_count} | security_blocked=${report.security_blocked_count} | params_missing=${report.parameter_missing_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: status=${row.status} ready_to_create=${boolLabel(row.ready_to_create)} requests=${row.request_count} missing_env=${row.missing_env.join(',') || '-'} renewal=${row.renewal_status ?? '-'}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
    for (const error of report.errors) console.error(`error ${error.file}: ${error.error}`);
  }
  if (!report.ok && failOnNotReady) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
