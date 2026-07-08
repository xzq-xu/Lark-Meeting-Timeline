import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { basename, dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdapterRuntimeManifest,
  buildMeetingPlatformAdapterRuntimeTarget,
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

export function parseMeetingPlatformAdapterRuntimeTargetCliArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

export function meetingPlatformAdapterRuntimeTargetCliOptionsFromArgs(args = new Map()) {
  const explicitPlatform = String(args.get('platform') || args.get('platform-key') || '');
  const requiredPlatforms = listArg(
    args.get('required-platforms') || args.get('platforms'),
    explicitPlatform ? [explicitPlatform] : ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  );
  return {
    baseUrl: String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787'),
    manifestFile: String(args.get('manifest-file') || args.get('runtime-manifest-file') || args.get('manifestFile') || ''),
    inputFile: String(args.get('input-file') || args.get('inputFile') || ''),
    outFile: String(args.get('out-file') || args.get('outFile') || ''),
    reportFile: String(args.get('report-file') || args.get('write-report') || ''),
    jsonOutput: boolArg(args.get('json'), false),
    includeTarget: boolArg(args.get('include-target'), false),
    includeManifest: boolArg(args.get('include-manifest'), false),
    failOnBlocked: boolArg(args.get('fail-on-blocked'), false) || boolArg(args.get('fail-on-failed'), false),
    url: String(args.get('url') || args.get('href') || args.get('meeting-url') || ''),
    title: String(args.get('title') || ''),
    platform: explicitPlatform,
    surface: String(args.get('surface') || args.get('preferred-surface') || ''),
    capturedAtMs: args.get('captured-at-ms') || args.get('capturedAtMs') || '',
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

async function targetInput(options = {}) {
  const fromFile = options.inputFile ? await readJson(options.inputFile) : {};
  return compactUndefined({
    ...fromFile,
    url: options.url || fromFile.url,
    title: options.title || fromFile.title,
    platform: options.platform || fromFile.platform,
    surface: options.surface || fromFile.surface,
    captured_at_ms: options.capturedAtMs || fromFile.captured_at_ms || fromFile.capturedAtMs,
  });
}

async function runtimeManifest(options = {}) {
  if (options.manifestFile) return readJson(options.manifestFile);
  return buildMeetingPlatformAdapterRuntimeManifest({}, {
    ...options,
    baseUrl: options.baseUrl,
    platforms: options.requiredPlatforms,
  });
}

function compactTarget(target = {}) {
  return compactUndefined({
    type: target.type,
    schema: target.schema,
    schema_version: target.schema_version,
    accepted: target.accepted,
    status: target.status,
    platform: target.platform,
    detection_reason: target.detection_reason,
    detected_meeting: target.detected_meeting,
    current_url: target.current_url,
    selected_surface: target.selected_surface,
    host_kind: target.host_kind,
    bridge_kind: target.bridge_kind,
    adapter_module: target.adapter_module,
    first_required_method: target.first_required_method,
    insert_method: target.insert_method,
    timestamp_field: target.timestamp_field,
    runtime_event_actions: target.runtime_event_actions,
    host_endpoints: target.host_endpoints,
    runtime_contract: target.runtime_contract,
    dispatch_policy: target.dispatch_policy,
    runtime_actions: target.runtime_actions,
    mark_template: target.mark_template,
    readiness: target.readiness,
    next_actions: target.next_actions,
  });
}

export async function buildMeetingPlatformAdapterRuntimeTargetCliReport(options = {}) {
  const {
    baseUrl = 'http://localhost:8787',
    outFile = '',
    reportFile = '',
    includeTarget = false,
    includeManifest = false,
  } = options;
  const manifest = await runtimeManifest(options);
  const input = await targetInput(options);
  const target = buildMeetingPlatformAdapterRuntimeTarget(manifest, input, {
    ...options,
    baseUrl,
  });
  if (outFile) await writeJson(resolve(outFile), target);
  const report = compactUndefined({
    type: 'meeting_platform_adapter_runtime_target_report',
    ok: target.accepted === true,
    base_url: baseUrl,
    manifest_file: options.manifestFile || undefined,
    input_file: options.inputFile || undefined,
    out_file: outFile || undefined,
    platform: target.platform,
    detection_reason: target.detection_reason,
    selected_surface: target.selected_surface,
    host_kind: target.host_kind,
    bridge_kind: target.bridge_kind,
    first_required_method: target.first_required_method,
    insert_method: target.insert_method,
    timestamp_field: target.timestamp_field,
    runtime_action_count: target.runtime_actions?.length ?? 0,
    issue_count: target.readiness?.issue_count ?? 0,
    issue_codes: target.readiness?.issues?.map((item) => item.code) ?? [],
    mark_template: target.mark_template,
    target: includeTarget ? target : compactTarget(target),
    manifest: includeManifest ? manifest : undefined,
    next_actions: target.next_actions,
  });
  if (reportFile) await writeJson(resolve(reportFile), report);
  return report;
}

export function formatMeetingPlatformAdapterRuntimeTargetCliReport(report = {}) {
  const lines = [
    `meeting_platform_adapter_runtime_target_report | ok=${boolLabel(report.ok)} | platform=${report.platform ?? 'n/a'} | surface=${report.selected_surface ?? 'n/a'} | host=${report.host_kind ?? 'n/a'} | bridge=${report.bridge_kind ?? 'n/a'} | first=${report.first_required_method ?? 'n/a'} | insert=${report.insert_method ?? 'n/a'} | issues=${report.issue_count ?? 0}`,
  ];
  if (report.issue_codes?.length > 0) lines.push(`issue_codes=${report.issue_codes.join(',')}`);
  if (report.out_file) lines.push(`out_file=${basename(report.out_file)}`);
  if (report.next_actions?.length > 0) lines.push(`next_actions=${report.next_actions.join(',')}`);
  return lines.join('\n');
}

export async function runMeetingPlatformAdapterRuntimeTargetCli(argv = process.argv.slice(2), io = console) {
  const options = meetingPlatformAdapterRuntimeTargetCliOptionsFromArgs(
    parseMeetingPlatformAdapterRuntimeTargetCliArgs(argv),
  );
  const report = await buildMeetingPlatformAdapterRuntimeTargetCliReport(options);
  if (options.jsonOutput) {
    io.log(JSON.stringify(report, null, 2));
  } else {
    io.log(formatMeetingPlatformAdapterRuntimeTargetCliReport(report));
  }
  if (!report.ok && options.failOnBlocked) process.exitCode = 2;
  return report;
}
