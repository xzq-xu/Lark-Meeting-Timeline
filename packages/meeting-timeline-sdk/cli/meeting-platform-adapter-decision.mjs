import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { basename, dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdapterDecision,
  buildMeetingPlatformAdapterDecisionMatrix,
} from '../adapters/platform-adapter-decision.mjs';

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

function compactUndefined(value = {}) {
  return Object.fromEntries(Object.entries(value).filter(([, item]) => item !== undefined));
}

function stripReports(decision = {}) {
  return {
    ...decision,
    reports: undefined,
  };
}

function compactDecision(decision = {}) {
  if (!decision || typeof decision !== 'object') return decision;
  return compactUndefined({
    type: decision.type,
    schema: decision.schema,
    schema_version: decision.schema_version,
    accepted: decision.accepted,
    realtime_ready: decision.realtime_ready,
    status: decision.status,
    platform: decision.platform,
    display_name: decision.display_name,
    input: decision.input,
    selected_surface: decision.selected_surface,
    surface_source: decision.surface_source,
    selected_route: decision.selected_route,
    recommended_mode: decision.recommended_mode,
    adaptation_strategy: decision.adaptation_strategy,
    adapter_blueprint: decision.adapter_blueprint,
    contracts: decision.contracts,
    route_summary: decision.route_summary,
    runtime_actions: decision.runtime_actions,
    evidence_requirements: decision.evidence_requirements,
    issues: decision.issues,
    next_actions: decision.next_actions,
  });
}

function compactMatrix(matrix = {}, options = {}) {
  const includeDecisions = options.includeDecisions === true;
  return compactUndefined({
    ...matrix,
    decisions: includeDecisions
      ? (matrix.decisions ?? []).map((decision) => compactDecision(stripReports(decision)))
      : undefined,
  });
}

export function parseMeetingPlatformAdapterDecisionCliArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

export function meetingPlatformAdapterDecisionCliOptionsFromArgs(args = new Map()) {
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
    includeDecisions: boolArg(args.get('include-decisions'), false),
    includeReports: boolArg(args.get('include-reports'), false),
    writeDecisions: boolArg(args.get('write-decisions'), true),
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

function decisionFile(outDir, platform) {
  return outDir ? resolve(outDir, `${platform}.json`) : undefined;
}

async function writeDecisionFiles(outDir, matrix = {}, options = {}) {
  if (!outDir || options.writeDecisions === false) return [];
  const writtenFiles = [];
  for (const decision of matrix.decisions ?? []) {
    const file = decisionFile(outDir, decision.platform);
    await writeJson(file, options.includeReports ? decision : compactDecision(stripReports(decision)));
    writtenFiles.push(file);
  }
  return writtenFiles;
}

async function decisionInput(options = {}) {
  const fromFile = options.inputFile ? await readJson(options.inputFile) : {};
  return {
    ...fromFile,
    url: options.url || fromFile.url,
    title: options.title || fromFile.title,
    platform: options.platform || fromFile.platform,
    surface: options.surface || fromFile.surface,
  };
}

export async function buildMeetingPlatformAdapterDecisionCliReport(options = {}) {
  const {
    baseUrl = 'http://localhost:8787',
    outDir = '',
    outFile = '',
    reportFile = '',
    includeDecisions = false,
    includeReports = false,
    writeDecisions = true,
    requiredPlatforms = ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  } = options;
  const input = await decisionInput(options);
  const matrix = buildMeetingPlatformAdapterDecisionMatrix(input, {
    ...options,
    baseUrl,
    platforms: requiredPlatforms,
    includeReports,
  });
  const singleDecision = options.platform || options.url || options.inputFile
    ? buildMeetingPlatformAdapterDecision(input, {
      ...options,
      baseUrl,
      includeReports,
    })
    : undefined;
  const writtenFiles = await writeDecisionFiles(outDir, matrix, { writeDecisions, includeReports });
  if (outFile) {
    await writeJson(resolve(outFile), singleDecision
      ? (includeReports ? singleDecision : compactDecision(stripReports(singleDecision)))
      : compactMatrix(includeReports ? matrix : {
        ...matrix,
        decisions: (matrix.decisions ?? []).map((decision) => stripReports(decision)),
      }, { includeDecisions }));
  }
  const rows = matrix.rows.map((row) => {
    const decision = (matrix.decisions ?? []).find((item) => item.platform === row.platform) ?? {};
    return {
      ...row,
      host_checklist_ready: decision.adaptation_strategy?.host_integration_checklist?.ready_for_realtime_host_wiring,
      first_host_step: decision.adaptation_strategy?.host_integration_checklist?.steps?.[0]?.id,
      decision_file: decisionFile(outDir, row.platform),
    };
  });
  const report = {
    type: 'meeting_platform_adapter_decision_report',
    ok: matrix.platform_count > 0 && matrix.accepted_count === matrix.platform_count,
    base_url: baseUrl,
    out_dir: outDir || undefined,
    out_file: outFile || undefined,
    write_decisions: writeDecisions,
    include_reports: includeReports,
    platform_count: matrix.platform_count,
    accepted_count: matrix.accepted_count,
    realtime_ready_count: matrix.realtime_ready_count,
    browser_surface_count: matrix.browser_surface_count,
    native_surface_count: matrix.native_surface_count,
    provider_reconcile_surface_count: matrix.provider_reconcile_surface_count,
    host_checklist_ready_count: rows.filter((row) => row.host_checklist_ready === true).length,
    required_platforms: requiredPlatforms,
    written_files: writtenFiles,
    rows,
    single_decision: singleDecision ? compactDecision(stripReports(singleDecision)) : undefined,
    matrix: compactMatrix(includeReports ? matrix : {
      ...matrix,
      decisions: (matrix.decisions ?? []).map((decision) => stripReports(decision)),
    }, { includeDecisions }),
    next_actions: matrix.next_actions,
  };
  if (reportFile) await writeJson(resolve(reportFile), report);
  return report;
}

export function formatMeetingPlatformAdapterDecisionCliReport(report = {}) {
  const lines = [
    `meeting_platform_adapter_decision_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | accepted=${report.accepted_count} | realtime_ready=${report.realtime_ready_count} | browser=${report.browser_surface_count} | native=${report.native_surface_count} | provider_only=${report.provider_reconcile_surface_count} | host_ready=${report.host_checklist_ready_count} | written=${report.written_files?.length ?? 0}`,
  ];
  for (const row of report.rows ?? []) {
    lines.push(`${row.platform}: accepted=${boolLabel(row.accepted)} realtime=${boolLabel(row.realtime_ready)} surface=${row.selected_surface ?? 'n/a'} observer=${row.selected_observer_mode ?? 'n/a'} first_evidence=${row.first_evidence_to_collect ?? 'n/a'} host_ready=${boolLabel(row.host_checklist_ready)} decision=${row.decision_file ?? 'n/a'}`);
  }
  if (report.single_decision?.platform) {
    lines.push(`single_decision=${report.single_decision.platform}:${report.single_decision.selected_surface}:${report.single_decision.status}`);
  }
  if (report.next_actions?.length > 0) lines.push(`next_actions=${report.next_actions.join(',')}`);
  if (report.out_dir) lines.push(`out_dir=${basename(report.out_dir)}`);
  if (report.out_file) lines.push(`out_file=${basename(report.out_file)}`);
  return lines.join('\n');
}

export async function runMeetingPlatformAdapterDecisionCli(argv = process.argv.slice(2), io = console) {
  const options = meetingPlatformAdapterDecisionCliOptionsFromArgs(
    parseMeetingPlatformAdapterDecisionCliArgs(argv),
  );
  const report = await buildMeetingPlatformAdapterDecisionCliReport(options);
  if (options.jsonOutput) {
    io.log(JSON.stringify(report, null, 2));
  } else {
    io.log(formatMeetingPlatformAdapterDecisionCliReport(report));
  }
  if (!report.ok && options.failOnBlocked) process.exitCode = 2;
  return report;
}
