#!/usr/bin/env node

import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformSessionBinding,
  buildMeetingPlatformSessionBindingMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-session-binding.mjs';
import {
  normalizeMeetingPlatform,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const inputFile = String(args.get('input') || args.get('binding-input') || '');
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const outDir = String(args.get('out-dir') || args.get('outDir') || '');
const writeBindings = args.get('write-bindings') === 'true' || Boolean(outDir);
const includeBindings = args.get('include-bindings') === 'true';
const jsonOutput = args.get('json') === 'true';
const requiredPlatforms = unique(String(
  args.get('platforms') || args.get('required-platforms') || 'google-meet,teams,zoom,webex,lark',
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
  return JSON.parse(await readFile(resolve(file), 'utf8'));
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function bindingFile(platform) {
  return resolve(outDir || 'data/meeting-platform-session-binding', `${platform}.json`);
}

function inputForPlatform(input, platform) {
  if (!input) return {};
  const key = normalizePlatformKey(platform);
  const direct = input[key] ?? input[platform];
  if (direct) return direct;
  const inputPlatform = normalizePlatformKey(
    input.current_meeting?.platform
      ?? input.currentMeeting?.platform
      ?? input.local_observer?.platform
      ?? input.localObserver?.platform
      ?? input.annotation?.platform
      ?? input.platform
      ?? key,
  );
  return inputPlatform === key ? input : {};
}

function optionNumber(name, fallback) {
  const value = args.get(name);
  if (value == null || value === '') return fallback;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function bindingOptions() {
  return {
    acceptScore: optionNumber('accept-score', 70),
    conflictScore: optionNumber('conflict-score', 20),
    maxStartDeltaMs: optionNumber('max-start-delta-ms', 10 * 60_000),
  };
}

function rowSummary(binding = {}) {
  return {
    platform: binding.platform,
    status: binding.status,
    accepted_for_realtime: binding.accepted_for_realtime,
    bind_to_current_axis: binding.bind_to_current_axis,
    should_start_axis: binding.should_start_axis,
    should_store_pending: binding.should_store_pending,
    selected_meeting_id: binding.selected_meeting?.meeting_id,
    candidate_count: binding.candidate_count,
    positive_match_count: binding.positive_match_count,
    conflict_count: binding.conflict_count,
    next_actions: binding.next_actions ?? [],
  };
}

async function buildReport() {
  const errors = [];
  let input;
  try {
    input = await readJson(inputFile);
  } catch (error) {
    errors.push({ file: resolve(inputFile), error: String(error?.message ?? error) });
  }
  const options = bindingOptions();
  const matrix = buildMeetingPlatformSessionBindingMatrix({
    platforms: requiredPlatforms,
    ...options,
  });
  const bindings = matrix.platforms.map((platform) => buildMeetingPlatformSessionBinding(
    platform,
    inputForPlatform(input, platform),
    options,
  ));
  const writtenFiles = [];
  if (writeBindings) {
    for (const binding of bindings) {
      const file = bindingFile(binding.platform);
      await writeJson(file, binding);
      writtenFiles.push(file);
    }
  }
  return {
    type: 'meeting_platform_session_binding_report',
    ok: errors.length === 0,
    requirement: 'bind_local_provider_and_annotation_identity_to_one_meeting_axis',
    input_file: inputFile ? resolve(inputFile) : undefined,
    platform_count: matrix.platform_count,
    bound_count: bindings.filter((binding) => binding.status === 'bound_to_current_axis').length,
    start_axis_count: bindings.filter((binding) => binding.should_start_axis).length,
    pending_count: bindings.filter((binding) => binding.status === 'pending_binding').length,
    conflict_count: bindings.filter((binding) => binding.status === 'binding_conflict').length,
    insufficient_count: bindings.filter((binding) => binding.status === 'insufficient_identity').length,
    provider_blocking_count: matrix.provider_blocking_count,
    transcript_blocking_count: matrix.transcript_blocking_count,
    rows: bindings.map((binding) => rowSummary(binding)),
    plan_rows: matrix.rows,
    written_files: writtenFiles,
    bindings: includeBindings ? bindings : undefined,
    next_actions: unique(bindings.flatMap((binding) => binding.next_actions ?? [])),
    errors,
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_session_binding_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | bound=${report.bound_count} | start_axis=${report.start_axis_count} | pending=${report.pending_count} | conflicts=${report.conflict_count} | insufficient=${report.insufficient_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: status=${row.status} accepted=${boolLabel(row.accepted_for_realtime)} candidates=${row.candidate_count} matches=${row.positive_match_count} conflicts=${row.conflict_count} meeting=${row.selected_meeting_id ?? 'none'}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
    for (const error of report.errors) console.error(`error ${error.file}: ${error.error}`);
  }
  if (!report.ok) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
