import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { basename, dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdapterLaunchPlan,
} from '../adapters/platform-adapter-launch-plan.mjs';

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

async function readJson(file) {
  return JSON.parse(await readFile(resolve(file), 'utf8'));
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

export function parseMeetingPlatformAdapterLaunchPlanCliArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

export function meetingPlatformAdapterLaunchPlanCliOptionsFromArgs(args = new Map()) {
  return {
    manifestFile: String(args.get('manifest-file') || args.get('manifestFile') || args.get('install-manifest') || 'data/meeting-platform-adapter-install-manifest.json'),
    outFile: String(args.get('out-file') || args.get('outFile') || ''),
    reportFile: String(args.get('report-file') || args.get('write-report') || ''),
    jsonOutput: args.get('json') === 'true',
    failOnBlocked: boolArg(args.get('fail-on-blocked'), false) || boolArg(args.get('fail-on-failed'), false),
    url: String(args.get('url') || args.get('href') || args.get('meeting-url') || ''),
    title: String(args.get('title') || ''),
    platform: String(args.get('platform') || args.get('platform-key') || ''),
    surface: String(args.get('surface') || args.get('preferred-surface') || ''),
    capturedAtMs: args.has('captured-at-ms') ? Number(args.get('captured-at-ms')) : undefined,
  };
}

export async function buildMeetingPlatformAdapterLaunchPlanCliReport(options = {}) {
  const {
    manifestFile = 'data/meeting-platform-adapter-install-manifest.json',
    outFile = '',
    reportFile = '',
    url = '',
    title = '',
    platform = '',
    surface = '',
    capturedAtMs,
  } = options;
  let manifest = {};
  let manifestReadError;
  try {
    manifest = await readJson(manifestFile);
  } catch (error) {
    manifestReadError = error.message;
  }
  const plan = buildMeetingPlatformAdapterLaunchPlan(manifest, {
    url: url || undefined,
    title: title || undefined,
    platform: platform || undefined,
    surface: surface || undefined,
    capturedAtMs,
  });
  if (outFile) await writeJson(resolve(outFile), plan);
  const report = {
    type: 'meeting_platform_adapter_launch_plan_report',
    ok: plan.accepted === true && !manifestReadError,
    manifest_file: manifestFile,
    out_file: outFile || undefined,
    platform: plan.platform,
    detection_reason: plan.detection_reason,
    selected_surface: plan.selected_surface,
    adapter_blueprint_available: plan.adapter_blueprint?.available === true,
    adapter_blueprint_primary_surface: plan.adapter_blueprint?.primary_surface,
    adapter_blueprint_first_gate: plan.adapter_blueprint?.first_acceptance_gate,
    raw_signal_validation_available: plan.raw_signal_validation?.available === true,
    raw_signal_validation_status: plan.raw_signal_validation?.status,
    raw_signal_validation_runtime_action_count: plan.raw_signal_validation?.runtime_action_count,
    accepted: plan.accepted,
    manifest_read_error: manifestReadError,
    issue_count: plan.readiness?.issue_count ?? 0,
    issues: plan.readiness?.issues ?? [],
    runtime_action_count: plan.runtime_actions?.length ?? 0,
    first_runtime_action: plan.runtime_actions?.[0]?.id,
    current_url: plan.current_url,
    next_actions: plan.next_actions,
    plan,
  };
  if (reportFile) await writeJson(resolve(reportFile), report);
  return report;
}

export function formatMeetingPlatformAdapterLaunchPlanCliReport(report = {}) {
  const lines = [
    `meeting_platform_adapter_launch_plan_report | ok=${boolLabel(report.ok)} | platform=${report.platform ?? 'n/a'} | surface=${report.selected_surface ?? 'n/a'} | blueprint_surface=${report.adapter_blueprint_primary_surface ?? 'n/a'} | raw_signal=${report.raw_signal_validation_status ?? 'n/a'} | reason=${report.detection_reason ?? 'n/a'} | accepted=${boolLabel(report.accepted)} | actions=${report.runtime_action_count} | issues=${report.issue_count}`,
  ];
  if (report.adapter_blueprint_first_gate) lines.push(`adapter_blueprint_first_gate=${report.adapter_blueprint_first_gate}`);
  if (report.first_runtime_action) lines.push(`first_runtime_action=${report.first_runtime_action}`);
  if (report.current_url) lines.push(`url=${report.current_url}`);
  if (report.manifest_read_error) lines.push(`manifest_read_error=${report.manifest_read_error}`);
  if (report.issues?.length > 0) lines.push(`issues=${report.issues.map((issue) => issue.code).join(',')}`);
  if (report.next_actions?.length > 0) lines.push(`next_actions=${report.next_actions.join(',')}`);
  if (report.out_file) lines.push(`out_file=${basename(report.out_file)}`);
  return lines.join('\n');
}

export async function runMeetingPlatformAdapterLaunchPlanCli(argv = process.argv.slice(2), io = console) {
  const options = meetingPlatformAdapterLaunchPlanCliOptionsFromArgs(
    parseMeetingPlatformAdapterLaunchPlanCliArgs(argv),
  );
  const report = await buildMeetingPlatformAdapterLaunchPlanCliReport(options);
  if (options.jsonOutput) {
    io.log(JSON.stringify(report, null, 2));
  } else {
    io.log(formatMeetingPlatformAdapterLaunchPlanCliReport(report));
  }
  if (!report.ok && options.failOnBlocked) process.exitCode = 2;
  return report;
}
