import { readdir, readFile, stat, writeFile, mkdir } from 'node:fs/promises';
import { basename, dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdapterInstallManifest,
} from '../adapters/platform-adapter-install-manifest.mjs';

const PLATFORM_DIR_ALIASES = Object.freeze({
  'google-meet': 'google_meet',
  google_meet: 'google_meet',
  meet: 'google_meet',
  teams: 'microsoft_teams',
  'microsoft-teams': 'microsoft_teams',
  microsoft_teams: 'microsoft_teams',
  zoom: 'zoom',
  webex: 'webex',
  lark: 'lark',
  feishu: 'lark',
});

function boolLabel(value) {
  return value ? 'yes' : 'no';
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

function platformDir(platform) {
  const raw = String(platform ?? '').trim();
  return PLATFORM_DIR_ALIASES[raw] ?? PLATFORM_DIR_ALIASES[raw.toLowerCase()] ?? raw.replace(/[^a-z0-9_-]+/gi, '_').toLowerCase();
}

async function readJson(file) {
  return JSON.parse(await readFile(resolve(file), 'utf8'));
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

async function collectPlanFiles(dir) {
  const resolved = resolve(dir);
  const info = await stat(resolved).catch(() => null);
  if (!info) return [];
  if (info.isFile()) return [resolved];
  const files = [];
  for (const entry of await readdir(resolved, { withFileTypes: true })) {
    const path = resolve(resolved, entry.name);
    if (entry.isDirectory()) {
      const candidate = resolve(path, 'adapter-import-plan.json');
      if (await stat(candidate).catch(() => null)) files.push(candidate);
      else files.push(...await collectPlanFiles(path));
    } else if (entry.isFile() && entry.name === 'adapter-import-plan.json') {
      files.push(path);
    }
  }
  return unique(files);
}

function planFileFor(dir, platform) {
  return resolve(dir, platformDir(platform), 'adapter-import-plan.json');
}

async function readPlans(options = {}) {
  const files = options.planFiles?.length
    ? options.planFiles
    : options.platforms?.length
      ? options.platforms.map((platform) => planFileFor(options.dir, platform))
      : await collectPlanFiles(options.dir);
  const plans = [];
  const missing = [];
  for (const file of files) {
    try {
      plans.push(await readJson(file));
    } catch {
      missing.push(resolve(file));
    }
  }
  return { plans, missing };
}

export function parseMeetingPlatformAdapterInstallManifestCliArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

export function meetingPlatformAdapterInstallManifestCliOptionsFromArgs(args = new Map()) {
  const dir = String(args.get('dir') || args.get('plan-dir') || args.get('planDir') || 'data/meeting-platform-adapter-import-plans');
  return {
    dir,
    outFile: String(args.get('out-file') || args.get('outFile') || args.get('manifest-file') || 'data/meeting-platform-adapter-install-manifest.json'),
    reportFile: String(args.get('report-file') || args.get('write-report') || ''),
    jsonOutput: args.get('json') === 'true',
    target: String(args.get('target') || args.get('acceptance-target') || 'static'),
    baseUrl: String(args.get('base-url') || args.get('baseUrl') || ''),
    installTarget: String(args.get('install-target') || args.get('installTarget') || 'host_project'),
    contentScriptJs: listArg(args.get('content-script-js') || args.get('contentScriptJs'), []),
    allowPartialInstall: boolArg(args.get('allow-partial-install'), false),
    failOnBlocked: args.get('fail-on-blocked') === 'true' || args.get('fail-on-failed') === 'true',
    planFiles: listArg(args.get('plan-file') || args.get('plan-files'), []),
    platforms: listArg(args.get('required-platforms') || args.get('platforms'), []),
  };
}

export async function buildMeetingPlatformAdapterInstallManifestCliReport(options = {}) {
  const {
    dir = 'data/meeting-platform-adapter-import-plans',
    outFile = 'data/meeting-platform-adapter-install-manifest.json',
    reportFile = '',
    target = 'static',
    baseUrl = '',
    installTarget = 'host_project',
    contentScriptJs = [],
    allowPartialInstall = false,
    planFiles = [],
    platforms = [],
  } = options;
  const { plans, missing } = await readPlans({ dir, planFiles, platforms });
  const manifest = buildMeetingPlatformAdapterInstallManifest(plans, {
    target,
    baseUrl: baseUrl || undefined,
    installTarget,
    contentScriptJs: contentScriptJs.length ? contentScriptJs : undefined,
    allowPartialInstall,
  });
  if (outFile) await writeJson(resolve(outFile), manifest);
  const report = {
    type: 'meeting_platform_adapter_install_manifest_report',
    ok: manifest.accepted === true && missing.length === 0,
    target,
    dir,
    out_file: outFile || undefined,
    plan_count: plans.length,
    platform_count: manifest.platform_count,
    ready_platform_count: manifest.ready_platform_count,
    blocked_platform_count: manifest.blocked_platform_count,
    selected_surfaces: manifest.selected_surfaces,
    browser_content_script_count: manifest.browser_extension?.content_scripts?.length ?? 0,
    provider_reconcile_count: manifest.provider_reconcile?.platform_count ?? 0,
    missing_plan_files: missing,
    missing_plan_file_count: missing.length,
    rows: manifest.platform_registry,
    readiness: manifest.readiness,
    next_actions: manifest.next_actions,
  };
  if (reportFile) await writeJson(resolve(reportFile), report);
  return report;
}

export function formatMeetingPlatformAdapterInstallManifestCliReport(report = {}) {
  const lines = [
    `meeting_platform_adapter_install_manifest_report | ok=${boolLabel(report.ok)} | target=${report.target} | plans=${report.plan_count} | platforms=${report.platform_count} | ready=${report.ready_platform_count} | blocked=${report.blocked_platform_count} | browser_content_scripts=${report.browser_content_script_count} | missing_plans=${report.missing_plan_file_count}`,
  ];
  for (const row of report.rows ?? []) {
    lines.push(`${row.platform}: ready=${boolLabel(row.ready)} surface=${row.selected_surface} timestamp=${row.timestamp_field ?? 'n/a'} first_sdk=${row.first_sdk_method ?? 'n/a'} mark_sdk=${row.mark_insert_method ?? 'n/a'} issue=${row.first_issue ?? 'none'}`);
  }
  if (report.missing_plan_files?.length > 0) lines.push(`missing_plan_files=${report.missing_plan_files.join(',')}`);
  if (report.next_actions?.length > 0) lines.push(`next_actions=${report.next_actions.join(',')}`);
  if (report.out_file) lines.push(`out_file=${basename(report.out_file)}`);
  return lines.join('\n');
}

export async function runMeetingPlatformAdapterInstallManifestCli(argv = process.argv.slice(2), io = console) {
  const options = meetingPlatformAdapterInstallManifestCliOptionsFromArgs(
    parseMeetingPlatformAdapterInstallManifestCliArgs(argv),
  );
  const report = await buildMeetingPlatformAdapterInstallManifestCliReport(options);
  if (options.jsonOutput) {
    io.log(JSON.stringify(report, null, 2));
  } else {
    io.log(formatMeetingPlatformAdapterInstallManifestCliReport(report));
  }
  if ((report.blocked_platform_count > 0 || report.missing_plan_file_count > 0) && options.failOnBlocked) process.exitCode = 2;
  return report;
}
