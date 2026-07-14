import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { basename, dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdapterRuntimeManifest,
} from '../adapters/platform-adapter-runtime-recipe.mjs';

function boolLabel(value) {
  return value ? 'yes' : 'no';
}

function compactUndefined(value = {}) {
  return Object.fromEntries(Object.entries(value).filter(([, item]) => item !== undefined));
}

function unique(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function listArg(value, fallback = []) {
  if (value == null || value === '') return fallback;
  return unique(String(value).split(',').map((item) => item.trim()).filter(Boolean));
}

function boolArg(value, fallback = false) {
  if (value == null || value === '') return fallback;
  const text = String(value).trim().toLowerCase();
  if (['1', 'true', 'yes', 'on'].includes(text)) return true;
  if (['0', 'false', 'no', 'off'].includes(text)) return false;
  return fallback;
}

export function parseMeetingPlatformAdapterRuntimeManifestCliArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

export function meetingPlatformAdapterRuntimeManifestCliOptionsFromArgs(args = new Map()) {
  const explicitPlatform = String(args.get('platform') || args.get('platform-key') || '');
  const requiredPlatforms = listArg(
    args.get('required-platforms') || args.get('platforms'),
    explicitPlatform ? [explicitPlatform] : ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  );
  return {
    baseUrl: String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787'),
    outFile: String(args.get('out-file') || args.get('outFile') || ''),
    reportFile: String(args.get('report-file') || args.get('write-report') || ''),
    inputFile: String(args.get('input-file') || args.get('inputFile') || ''),
    jsonOutput: boolArg(args.get('json'), false),
    includeManifest: boolArg(args.get('include-manifest'), false),
    includeRecipes: boolArg(args.get('include-recipes'), false),
    failOnBlocked: boolArg(args.get('fail-on-blocked'), false) || boolArg(args.get('fail-on-failed'), false),
    surface: String(args.get('surface') || args.get('preferred-surface') || ''),
    hostProfile: String(args.get('host-profile') || args.get('hostProfile') || ''),
    filterActiveSpeakerSamples: boolArg(args.get('filter-active-speaker-samples'), true),
    requiredPlatforms,
  };
}

async function readJson(file) {
  return JSON.parse(await readFile(resolve(file), 'utf8'));
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

async function manifestInput(options = {}) {
  return options.inputFile ? await readJson(options.inputFile) : {};
}

function reportRow(row = {}) {
  return compactUndefined({
    platform: row.platform,
    display_name: row.display_name,
    accepted: row.accepted,
    runtime_ready: row.runtime_ready,
    selected_surface: row.selected_surface,
    host_kind: row.host_kind,
    bridge_kind: row.bridge_kind,
    first_required_method: row.first_required_method,
    insert_method: row.insert_method,
    timestamp_field: row.timestamp_field,
    speaker_position_markers: row.speaker_position_markers,
    provider_blocks_realtime: row.provider_reconcile?.blocks_realtime_annotation,
    transcript_blocks_realtime: row.transcript?.blocks_realtime_annotation,
  });
}

export async function buildMeetingPlatformAdapterRuntimeManifestCliReport(options = {}) {
  const {
    baseUrl = 'http://localhost:8787',
    outFile = '',
    reportFile = '',
    includeManifest = false,
    includeRecipes = false,
    requiredPlatforms = ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  } = options;
  const input = await manifestInput(options);
  const manifest = buildMeetingPlatformAdapterRuntimeManifest(input, {
    ...options,
    baseUrl,
    platforms: requiredPlatforms,
    includeRecipes,
  });
  if (outFile) await writeJson(resolve(outFile), manifest);
  const report = compactUndefined({
    type: 'meeting_platform_adapter_runtime_manifest_report',
    ok: manifest.runtime_ready === true,
    base_url: baseUrl,
    out_file: outFile || undefined,
    platform_count: manifest.platform_count,
    accepted_count: manifest.accepted_count,
    runtime_ready_count: manifest.runtime_ready_count,
    browser_surface_count: manifest.browser_surface_count,
    native_surface_count: manifest.native_surface_count,
    provider_reconcile_surface_count: manifest.provider_reconcile_surface_count,
    local_surface_count: manifest.local_surface_count,
    raw_signal_runtime_event_count: manifest.raw_signal_runtime_event_count,
    required_platforms: requiredPlatforms,
    host_endpoints: manifest.host_endpoints,
    runtime_contract: manifest.runtime_contract,
    bridge_groups: manifest.bridge_groups,
    rows: manifest.platform_registry.rows.map((row) => reportRow(row)),
    manifest: includeManifest ? manifest : undefined,
    next_actions: manifest.next_actions,
  });
  if (reportFile) await writeJson(resolve(reportFile), report);
  return report;
}

export function formatMeetingPlatformAdapterRuntimeManifestCliReport(report = {}) {
  const lines = [
    `meeting_platform_adapter_runtime_manifest_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | accepted=${report.accepted_count} | runtime_ready=${report.runtime_ready_count} | browser=${report.browser_surface_count} | native=${report.native_surface_count} | provider_only=${report.provider_reconcile_surface_count} | runtime_events=${report.raw_signal_runtime_event_count}`,
  ];
  for (const group of report.bridge_groups ?? []) {
    lines.push(`bridge=${group.bridge_kind} host=${group.host_kind} platforms=${group.platforms?.join(',') ?? ''}`);
  }
  for (const row of report.rows ?? []) {
    lines.push(`${row.platform}: runtime=${boolLabel(row.runtime_ready)} host=${row.host_kind ?? 'n/a'} bridge=${row.bridge_kind ?? 'n/a'} first=${row.first_required_method ?? 'n/a'} insert=${row.insert_method ?? 'n/a'}`);
  }
  if (report.out_file) lines.push(`out_file=${basename(report.out_file)}`);
  if (report.next_actions?.length > 0) lines.push(`next_actions=${report.next_actions.join(',')}`);
  return lines.join('\n');
}

export async function runMeetingPlatformAdapterRuntimeManifestCli(argv = process.argv.slice(2), io = console) {
  const options = meetingPlatformAdapterRuntimeManifestCliOptionsFromArgs(
    parseMeetingPlatformAdapterRuntimeManifestCliArgs(argv),
  );
  const report = await buildMeetingPlatformAdapterRuntimeManifestCliReport(options);
  if (options.jsonOutput) {
    io.log(JSON.stringify(report, null, 2));
  } else {
    io.log(formatMeetingPlatformAdapterRuntimeManifestCliReport(report));
  }
  if (!report.ok && options.failOnBlocked) process.exitCode = 2;
  return report;
}
