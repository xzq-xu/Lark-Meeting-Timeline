import { readdir, readFile, stat, writeFile, mkdir } from 'node:fs/promises';
import { basename, dirname, relative, resolve } from 'node:path';
import {
  buildMeetingPlatformAdapterImportPlanMatrix,
} from '../adapters/platform-adapter-import-plan.mjs';

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

async function readOptionalJson(file) {
  try {
    return await readJson(file);
  } catch {
    return undefined;
  }
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

async function collectFiles(root, base = root) {
  const resolved = resolve(root);
  const info = await stat(resolved).catch(() => null);
  if (!info) return [];
  if (info.isFile()) return [relative(base, resolved).replace(/\\/g, '/')];
  const rows = [];
  for (const entry of await readdir(resolved, { withFileTypes: true })) {
    const path = resolve(resolved, entry.name);
    if (entry.isDirectory()) rows.push(...await collectFiles(path, base));
    else if (entry.isFile()) rows.push(relative(base, path).replace(/\\/g, '/'));
  }
  return rows;
}

function packageFileFor(dir, platform) {
  return resolve(dir, platformDir(platform), 'adapter-export-package.json');
}

async function readPackages(options = {}) {
  const files = options.packageFiles?.length
    ? options.packageFiles
    : options.platforms.map((platform) => packageFileFor(options.dir, platform));
  const packages = [];
  const missing = [];
  for (const file of files) {
    try {
      const pkg = await readJson(file);
      const blueprintPath = pkg.host_files?.find((hostFile) => hostFile.source === 'adapter_blueprint')?.path;
      const adapterBlueprint = blueprintPath ? await readOptionalJson(resolve(options.dir, blueprintPath)) : undefined;
      packages.push(adapterBlueprint ? {
        ...pkg,
        artifacts: {
          ...(pkg.artifacts ?? {}),
          adapter_blueprint: adapterBlueprint,
        },
      } : pkg);
    } catch {
      missing.push(resolve(file));
    }
  }
  return { packages, missing };
}

export function parseMeetingPlatformAdapterImportPlanCliArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

export function meetingPlatformAdapterImportPlanCliOptionsFromArgs(args = new Map()) {
  const dir = String(args.get('dir') || args.get('package-dir') || args.get('packageDir') || 'data/meeting-platform-adapter-export-packages');
  return {
    dir,
    outDir: String(args.get('out-dir') || args.get('outDir') || 'data/meeting-platform-adapter-import-plans'),
    reportFile: String(args.get('report-file') || args.get('write-report') || ''),
    jsonOutput: args.get('json') === 'true',
    target: String(args.get('target') || args.get('acceptance-target') || 'static'),
    surface: String(args.get('surface') || args.get('preferred-surface') || ''),
    allowCustomAuthoring: boolArg(args.get('allow-custom-authoring'), false),
    failOnBlocked: args.get('fail-on-blocked') === 'true' || args.get('fail-on-failed') === 'true',
    packageFiles: listArg(args.get('package-file') || args.get('package-files'), []),
    platforms: listArg(args.get('required-platforms') || args.get('platforms'), ['google-meet', 'teams', 'zoom', 'webex', 'lark']),
    availableFiles: listArg(args.get('available-files'), []),
  };
}

export async function buildMeetingPlatformAdapterImportPlanCliReport(options = {}) {
  const {
    dir = 'data/meeting-platform-adapter-export-packages',
    outDir = 'data/meeting-platform-adapter-import-plans',
    reportFile = '',
    target = 'static',
    surface = '',
    allowCustomAuthoring = false,
    packageFiles = [],
    platforms = ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
    availableFiles,
  } = options;
  const { packages, missing } = await readPackages({ dir, packageFiles, platforms });
  const collectedFiles = availableFiles?.length ? availableFiles : await collectFiles(dir);
  const matrix = buildMeetingPlatformAdapterImportPlanMatrix(packages, {
    target,
    availableFiles: collectedFiles,
    surface: surface || undefined,
    allowCustomAuthoring,
  });
  const writtenFiles = [];
  if (outDir) {
    for (const plan of matrix.plans) {
      const file = resolve(outDir, plan.platform, 'adapter-import-plan.json');
      await writeJson(file, plan);
      writtenFiles.push(file);
    }
  }
  const report = {
    type: 'meeting_platform_adapter_import_plan_report',
    ok: packages.length > 0,
    target,
    surface: surface || undefined,
    dir,
    out_dir: outDir || undefined,
    package_count: matrix.package_count,
    accepted_count: matrix.accepted_count,
    blocked_count: matrix.blocked_count,
    adapter_preflight_startup_ready_count: matrix.adapter_preflight_startup_ready_count,
    adapter_preflight_realtime_ready_count: matrix.adapter_preflight_realtime_ready_count,
    missing_package_files: missing,
    missing_package_file_count: missing.length,
    available_file_count: collectedFiles.length,
    written_files: writtenFiles,
    rows: matrix.rows.map((row) => ({
      ...row,
      plan_file: outDir ? resolve(outDir, row.platform, 'adapter-import-plan.json') : undefined,
    })),
    matrix: {
      ...matrix,
      plans: undefined,
    },
    next_actions: matrix.next_actions,
  };
  if (reportFile) await writeJson(resolve(reportFile), report);
  return report;
}

export function formatMeetingPlatformAdapterImportPlanCliReport(report = {}) {
  const lines = [
    `meeting_platform_adapter_import_plan_report | ok=${boolLabel(report.ok)} | target=${report.target} | packages=${report.package_count} | accepted=${report.accepted_count} | blocked=${report.blocked_count} | preflight=${report.adapter_preflight_realtime_ready_count}/${report.package_count} | missing_packages=${report.missing_package_file_count} | available_files=${report.available_file_count} | written=${report.written_files?.length ?? 0}`,
  ];
  for (const row of report.rows ?? []) {
    lines.push(`${row.platform}: accepted=${boolLabel(row.accepted)} surface=${row.selected_surface} hard_contract=${boolLabel(row.hard_contract_ready)} surface_ready=${boolLabel(row.selected_surface_ready)} files=${boolLabel(row.file_coverage_ready)} preflight=${row.adapter_preflight_status ?? 'n/a'} realtime=${boolLabel(row.adapter_preflight_realtime_ready)} missing=${row.missing_file_count ?? 0} next=${row.first_next_action ?? 'none'} plan=${row.plan_file ?? 'n/a'}`);
  }
  if (report.missing_package_files?.length > 0) lines.push(`missing_package_files=${report.missing_package_files.join(',')}`);
  if (report.next_actions?.length > 0) lines.push(`next_actions=${report.next_actions.join(',')}`);
  if (report.out_dir) lines.push(`out_dir=${basename(report.out_dir)}`);
  return lines.join('\n');
}

export async function runMeetingPlatformAdapterImportPlanCli(argv = process.argv.slice(2), io = console) {
  const options = meetingPlatformAdapterImportPlanCliOptionsFromArgs(
    parseMeetingPlatformAdapterImportPlanCliArgs(argv),
  );
  const report = await buildMeetingPlatformAdapterImportPlanCliReport(options);
  if (options.jsonOutput) {
    io.log(JSON.stringify(report, null, 2));
  } else {
    io.log(formatMeetingPlatformAdapterImportPlanCliReport(report));
  }
  if (report.blocked_count > 0 && options.failOnBlocked) process.exitCode = 2;
  return report;
}
