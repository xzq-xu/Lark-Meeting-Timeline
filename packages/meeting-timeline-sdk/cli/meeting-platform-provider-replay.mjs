import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { basename, dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformProviderReplayMatrix,
} from '../adapters/platform-ingest.mjs';
import { normalizeMeetingPlatform } from '../adapters/platform-setup.mjs';

function boolLabel(value) {
  return value ? 'yes' : 'no';
}

function boolArg(value, fallback = false) {
  if (value == null || value === '') return fallback;
  const text = String(value).trim().toLowerCase();
  if (['1', 'true', 'yes', 'on'].includes(text)) return true;
  if (['0', 'false', 'no', 'off'].includes(text)) return false;
  return fallback;
}

function unique(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function listArg(value, fallback = []) {
  if (value == null || value === '') return fallback;
  return unique(String(value).split(',').map((item) => item.trim()).filter(Boolean));
}

async function readJson(file) {
  return JSON.parse(await readFile(resolve(file), 'utf8'));
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function recordsFromJson(json, platforms = []) {
  if (json == null) return undefined;
  if (Array.isArray(json)) {
    const platform = platforms.length === 1 ? normalizeMeetingPlatform(platforms[0]) : undefined;
    return platform ? { [platform]: json } : undefined;
  }
  if (typeof json !== 'object') return undefined;
  const explicit = firstNonEmpty(
    json.recordsByPlatform,
    json.records_by_platform,
    json.providerEventsByPlatform,
    json.provider_events_by_platform,
  );
  if (explicit && typeof explicit === 'object' && !Array.isArray(explicit)) return explicit;
  const records = firstNonEmpty(json.records, json.events, json.samples, json.provider_events, json.providerEvents);
  if (Array.isArray(records)) {
    const platform = normalizeMeetingPlatform(firstNonEmpty(json.platform, json.provider, json.adapter, platforms[0]));
    return platform ? { [platform]: records } : undefined;
  }
  return json;
}

export function parseMeetingPlatformProviderReplayCliArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

export function meetingPlatformProviderReplayCliOptionsFromArgs(args = new Map()) {
  return {
    platforms: listArg(args.get('platforms') || args.get('platform-keys'), ['google-meet', 'teams', 'zoom', 'webex', 'lark']),
    recordsFile: String(args.get('records-file') || args.get('provider-events-file') || args.get('input-file') || ''),
    outFile: String(args.get('out-file') || args.get('outFile') || ''),
    reportFile: String(args.get('report-file') || args.get('write-report') || ''),
    jsonOutput: args.get('json') === 'true',
    failOnRejected: boolArg(args.get('fail-on-rejected'), false) || boolArg(args.get('fail-on-failed'), false),
    target: String(args.get('target') || args.get('acceptance-target') || 'pilot'),
    requiredCoverage: listArg(args.get('required-coverage') || args.get('requiredCoverage'), []),
    baseReceivedAtMs: args.has('base-received-at-ms') ? Number(args.get('base-received-at-ms')) : undefined,
  };
}

export async function buildMeetingPlatformProviderReplayCliReport(options = {}) {
  const {
    platforms = ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
    recordsFile = '',
    outFile = '',
    reportFile = '',
    target = 'pilot',
    requiredCoverage = [],
    baseReceivedAtMs,
  } = options;
  let recordsByPlatform;
  let recordsReadError;
  if (recordsFile) {
    try {
      recordsByPlatform = recordsFromJson(await readJson(recordsFile), platforms);
    } catch (error) {
      recordsReadError = error.message;
    }
  }
  const matrix = recordsReadError
    ? {
      schema: 'meeting_platform_provider_replay_matrix',
      accepted: false,
      target,
      platform_count: platforms.length,
      accepted_count: 0,
      runtime_event_count: 0,
      signal_count: 0,
      platforms,
      rows: [],
      issues: ['records_file_read_error'],
      next_actions: ['provide_valid_provider_records_json_or_omit_records_file_to_use_sdk_samples'],
    }
    : buildMeetingPlatformProviderReplayMatrix({
      platforms,
      recordsByPlatform,
      target,
      requiredCoverage: requiredCoverage.length ? requiredCoverage : undefined,
      baseReceivedAtMs,
    });
  const report = {
    type: 'meeting_platform_provider_replay_cli_report',
    ok: matrix.accepted === true && !recordsReadError,
    target,
    records_file: recordsFile || undefined,
    records_read_error: recordsReadError,
    generated_samples: !recordsFile,
    out_file: outFile || undefined,
    platform_count: matrix.platform_count ?? 0,
    accepted_count: matrix.accepted_count ?? 0,
    runtime_event_count: matrix.runtime_event_count ?? 0,
    signal_count: matrix.signal_count ?? 0,
    platforms: matrix.platforms ?? platforms,
    rows: matrix.rows ?? [],
    issues: matrix.issues ?? [],
    next_actions: matrix.next_actions ?? [],
    replay_matrix: matrix,
  };
  if (outFile) await writeJson(resolve(outFile), matrix);
  if (reportFile) await writeJson(resolve(reportFile), report);
  return report;
}

export function formatMeetingPlatformProviderReplayCliReport(report = {}) {
  const lines = [
    `meeting_platform_provider_replay_cli_report | ok=${boolLabel(report.ok)} | target=${report.target ?? 'pilot'} | platforms=${report.platform_count ?? 0} | accepted=${report.accepted_count ?? 0} | runtime_events=${report.runtime_event_count ?? 0} | signals=${report.signal_count ?? 0} | generated_samples=${boolLabel(report.generated_samples)}`,
  ];
  for (const row of report.rows ?? []) {
    lines.push(`${row.platform}: accepted=${boolLabel(row.accepted)} records=${row.record_count ?? 0} start=${boolLabel(row.coverage?.meeting_start)} end=${boolLabel(row.coverage?.meeting_end)} participant=${boolLabel(row.coverage?.participant_track)} artifact=${boolLabel(row.coverage?.artifact_ready)} provider_nonblocking=${boolLabel(row.provider_events_block_realtime === false)} issues=${row.issues?.join(',') || 'none'}`);
  }
  if (report.records_file) lines.push(`records_file=${report.records_file}`);
  if (report.records_read_error) lines.push(`records_read_error=${report.records_read_error}`);
  if (report.next_actions?.length > 0) lines.push(`next_actions=${report.next_actions.join(',')}`);
  if (report.out_file) lines.push(`out_file=${basename(report.out_file)}`);
  return lines.join('\n');
}

export async function runMeetingPlatformProviderReplayCli(argv = process.argv.slice(2), io = console) {
  const options = meetingPlatformProviderReplayCliOptionsFromArgs(
    parseMeetingPlatformProviderReplayCliArgs(argv),
  );
  const report = await buildMeetingPlatformProviderReplayCliReport(options);
  if (options.jsonOutput) {
    io.log(JSON.stringify(report, null, 2));
  } else {
    io.log(formatMeetingPlatformProviderReplayCliReport(report));
  }
  if (!report.ok && options.failOnRejected) process.exitCode = 2;
  return report;
}
