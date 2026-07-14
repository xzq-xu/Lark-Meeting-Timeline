import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { basename, dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdapterStartupPlan,
  buildMeetingPlatformAdapterStartupPlanMatrix,
} from '../adapters/platform-adapter-startup.mjs';

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

function stripReports(plan = {}) {
  return {
    ...plan,
    reports: undefined,
  };
}

function compactStartupPlan(plan = {}) {
  if (!plan || typeof plan !== 'object') return plan;
  return compactUndefined({
    type: plan.type,
    schema: plan.schema,
    schema_version: plan.schema_version,
    id: plan.id,
    accepted: plan.accepted,
    realtime_startup_ready: plan.realtime_startup_ready,
    status: plan.status,
    platform: plan.platform,
    display_name: plan.display_name,
    input: plan.input,
    selected_surface: plan.selected_surface,
    install_target: plan.install_target,
    runtime_contract: plan.runtime_contract,
    browser: plan.browser ? compactUndefined({
      matches: plan.browser.matches,
      host_permissions: plan.browser.host_permissions,
      content_script: plan.browser.content_script ? compactUndefined({
        matches: plan.browser.content_script.matches,
        js: plan.browser.content_script.js,
        run_at: plan.browser.content_script.run_at,
      }) : undefined,
    }) : undefined,
    bridge: plan.bridge ? compactUndefined({
      preferred_bridge: plan.bridge.preferred_bridge,
      module: plan.bridge.module,
      create_function: plan.bridge.create_function,
      install_function: plan.bridge.install_function,
      first_method: plan.bridge.first_method,
    }) : undefined,
    runtime: plan.runtime ? compactUndefined({
      preset: plan.runtime.preset,
      observer_scheduler_ready: plan.runtime.observer_scheduler_ready,
      runtime_host_ready: plan.runtime.runtime_host_ready,
    }) : undefined,
    axis: plan.axis,
    actions: plan.actions,
    message_contract: plan.message_contract,
    code_refs: plan.code_refs,
    provider_reconcile: plan.provider_reconcile,
    issues: plan.issues,
    next_actions: plan.next_actions,
  });
}

function compactUndefined(value = {}) {
  return Object.fromEntries(Object.entries(value).filter(([, item]) => item !== undefined));
}

function stripMatrixReports(matrix = {}) {
  return {
    ...matrix,
    plans: (matrix.plans ?? []).map((plan) => stripReports(plan)),
  };
}

function compactMatrix(matrix = {}, options = {}) {
  const includePlans = options.includePlans === true;
  return compactUndefined({
    ...matrix,
    plans: includePlans ? (matrix.plans ?? []).map((plan) => compactStartupPlan(plan)) : undefined,
  });
}

export function parseMeetingPlatformAdapterStartupCliArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

export function meetingPlatformAdapterStartupCliOptionsFromArgs(args = new Map()) {
  const explicitPlatform = String(args.get('platform') || args.get('platform-key') || '');
  const requiredPlatforms = listArg(
    args.get('required-platforms') || args.get('platforms'),
    explicitPlatform ? [explicitPlatform] : ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  );
  return {
    baseUrl: String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787'),
    outDir: String(args.get('out-dir') || args.get('outDir') || ''),
    outFile: String(args.get('out-file') || args.get('outFile') || ''),
    reportFile: String(args.get('report-file') || args.get('write-report') || ''),
    inputFile: String(args.get('input-file') || args.get('inputFile') || ''),
    jsonOutput: boolArg(args.get('json'), false),
    includePlans: boolArg(args.get('include-plans'), false),
    includeReports: boolArg(args.get('include-reports'), false),
    writePlans: boolArg(args.get('write-plans'), true),
    failOnBlocked: boolArg(args.get('fail-on-blocked'), false) || boolArg(args.get('fail-on-failed'), false),
    url: String(args.get('url') || args.get('href') || args.get('meeting-url') || ''),
    title: String(args.get('title') || ''),
    platform: explicitPlatform,
    surface: String(args.get('surface') || args.get('preferred-surface') || ''),
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

function planFile(outDir, platform) {
  return outDir ? resolve(outDir, `${platform}.json`) : undefined;
}

async function writeStartupPlans(outDir, matrix = {}, options = {}) {
  if (!outDir || options.writePlans === false) return [];
  const writtenFiles = [];
  for (const plan of matrix.plans ?? []) {
    const file = planFile(outDir, plan.platform);
    await writeJson(file, options.includeReports ? plan : stripReports(plan));
    writtenFiles.push(file);
  }
  return writtenFiles;
}

async function startupInput(options = {}) {
  const fromFile = options.inputFile ? await readJson(options.inputFile) : {};
  return {
    ...fromFile,
    url: options.url || fromFile.url,
    title: options.title || fromFile.title,
    platform: options.platform || fromFile.platform,
    surface: options.surface || fromFile.surface,
  };
}

export async function buildMeetingPlatformAdapterStartupCliReport(options = {}) {
  const {
    baseUrl = 'http://localhost:8787',
    outDir = '',
    outFile = '',
    reportFile = '',
    includePlans = false,
    includeReports = false,
    writePlans = true,
    requiredPlatforms = ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  } = options;
  const input = await startupInput(options);
  const matrix = buildMeetingPlatformAdapterStartupPlanMatrix(input, {
    ...options,
    baseUrl,
    platforms: requiredPlatforms,
    includeReports,
  });
  const singlePlan = options.platform || options.url || options.inputFile
    ? buildMeetingPlatformAdapterStartupPlan(input, {
      ...options,
      baseUrl,
      includeReports,
    })
    : undefined;
  const writtenFiles = await writeStartupPlans(outDir, matrix, { writePlans, includeReports });
  if (outFile) await writeJson(resolve(outFile), singlePlan ?? (includeReports ? matrix : stripMatrixReports(matrix)));
  const report = {
    type: 'meeting_platform_adapter_startup_plan_report',
    ok: matrix.platform_count > 0 && matrix.accepted_count === matrix.platform_count,
    base_url: baseUrl,
    out_dir: outDir || undefined,
    out_file: outFile || undefined,
    write_plans: writePlans,
    include_reports: includeReports,
    platform_count: matrix.platform_count,
    accepted_count: matrix.accepted_count,
    realtime_startup_ready_count: matrix.realtime_startup_ready_count,
    browser_surface_count: matrix.browser_surface_count,
    native_surface_count: matrix.native_surface_count,
    provider_reconcile_surface_count: matrix.provider_reconcile_surface_count,
    required_platforms: requiredPlatforms,
    written_files: writtenFiles,
    rows: matrix.rows.map((row) => ({
      ...row,
      plan_file: planFile(outDir, row.platform),
    })),
    single_plan: singlePlan ? compactStartupPlan(singlePlan) : undefined,
    matrix: compactMatrix(includeReports ? matrix : stripMatrixReports(matrix), { includePlans }),
    next_actions: matrix.next_actions,
  };
  if (reportFile) await writeJson(resolve(reportFile), report);
  return report;
}

export function formatMeetingPlatformAdapterStartupCliReport(report = {}) {
  const lines = [
    `meeting_platform_adapter_startup_plan_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | accepted=${report.accepted_count} | realtime_ready=${report.realtime_startup_ready_count} | browser=${report.browser_surface_count} | native=${report.native_surface_count} | provider_only=${report.provider_reconcile_surface_count} | written=${report.written_files?.length ?? 0}`,
  ];
  for (const row of report.rows ?? []) {
    lines.push(`${row.platform}: accepted=${boolLabel(row.accepted)} realtime=${boolLabel(row.realtime_startup_ready)} surface=${row.selected_surface ?? 'n/a'} install=${row.install_target ?? 'n/a'} observe=${row.observe_action ?? 'n/a'} insert=${row.insert_action ?? 'n/a'} issues=${row.issue_count ?? 0} plan=${row.plan_file ?? 'n/a'}`);
  }
  if (report.single_plan?.platform) {
    lines.push(`single_plan=${report.single_plan.platform}:${report.single_plan.selected_surface}:${report.single_plan.status}`);
  }
  if (report.next_actions?.length > 0) lines.push(`next_actions=${report.next_actions.join(',')}`);
  if (report.out_dir) lines.push(`out_dir=${basename(report.out_dir)}`);
  if (report.out_file) lines.push(`out_file=${basename(report.out_file)}`);
  return lines.join('\n');
}

export async function runMeetingPlatformAdapterStartupCli(argv = process.argv.slice(2), io = console) {
  const options = meetingPlatformAdapterStartupCliOptionsFromArgs(
    parseMeetingPlatformAdapterStartupCliArgs(argv),
  );
  const report = await buildMeetingPlatformAdapterStartupCliReport(options);
  if (options.jsonOutput) {
    io.log(JSON.stringify(report, null, 2));
  } else {
    io.log(formatMeetingPlatformAdapterStartupCliReport(report));
  }
  if (!report.ok && options.failOnBlocked) process.exitCode = 2;
  return report;
}
