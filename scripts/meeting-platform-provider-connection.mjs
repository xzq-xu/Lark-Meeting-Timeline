#!/usr/bin/env node

import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformProviderConnectionMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-provider-connection.mjs';
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
const failOnBlocked = args.get('fail-on-blocked') === 'true' || args.get('fail-on-incomplete') === 'true';
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

function summarizePack(pack = {}) {
  return {
    platform: pack.platform,
    display_name: pack.display_name,
    transport: pack.transport,
    provider_role: pack.provider_role,
    ready: pack.readiness?.ready === true,
    blocking_count: pack.readiness?.blocking_count ?? 0,
    warning_count: pack.readiness?.warning_count ?? 0,
    endpoint: pack.endpoint,
    subscription_builders: pack.subscription?.builders ?? [],
    has_subscription_request: pack.subscription?.request != null,
    required_permissions: pack.permissions?.required_permissions ?? [],
    required_scopes: pack.permissions?.required_scopes ?? [],
    required_env: pack.security?.required_env ?? [],
    missing_env: pack.security?.missing_env ?? [],
    verifier: pack.security?.verifier,
    event_count: pack.event_mapping?.length ?? 0,
    docs: pack.official_docs?.map((doc) => doc.url) ?? [],
    provider_events_block_realtime: pack.realtime_annotation_policy?.provider_events_block_realtime,
    next_actions: pack.next_actions ?? [],
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
  const matrix = buildMeetingPlatformProviderConnectionMatrix({
    baseUrl,
    env: process.env,
    platforms: requiredPlatforms,
    subscriptions,
  });
  const ok = matrix.platform_count > 0 && matrix.blocked_count === 0 && errors.length === 0;
  return {
    type: 'meeting_platform_provider_connection_report',
    ok,
    base_url: baseUrl,
    subscriptions_file: subscriptionFile ? resolve(subscriptionFile) : undefined,
    required_platforms: requiredPlatforms,
    platform_count: matrix.platform_count,
    ready_count: matrix.ready_count,
    blocked_count: matrix.blocked_count,
    rows: matrix.packs.map((pack) => summarizePack(pack)),
    matrix,
    next_actions: unique(matrix.next_actions ?? []),
    errors,
  };
}

try {
  const report = await buildReport();
  if (reportFile) {
    const targetFile = resolve(reportFile);
    await mkdir(dirname(targetFile), { recursive: true });
    await writeFile(targetFile, `${JSON.stringify(report, null, 2)}\n`, 'utf8');
  }
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_provider_connection_report | ok=${boolLabel(report.ok)} | ready=${report.ready_count}/${report.platform_count} | blocked=${report.blocked_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: ready=${boolLabel(row.ready)} transport=${row.transport} verifier=${row.verifier ?? 'none'} missing_env=${row.missing_env.join(',') || '-'}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
    for (const error of report.errors) console.error(`error ${error.file}: ${error.error}`);
  }
  if (!report.ok && failOnBlocked) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
