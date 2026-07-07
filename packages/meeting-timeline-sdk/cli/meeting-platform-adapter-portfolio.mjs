import { mkdir, writeFile } from 'node:fs/promises';
import { basename, dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdapterPortfolio,
} from '../adapters/platform-adapter-portfolio.mjs';

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

function stripArtifacts(portfolio = {}) {
  return {
    ...portfolio,
    items: (portfolio.items ?? []).map((item) => ({
      ...item,
      artifacts: undefined,
    })),
  };
}

export function parseMeetingPlatformAdapterPortfolioCliArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

export function meetingPlatformAdapterPortfolioCliOptionsFromArgs(args = new Map()) {
  return {
    baseUrl: String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787'),
    outDir: String(args.get('out-dir') || args.get('outDir') || ''),
    reportFile: String(args.get('report-file') || args.get('write-report') || ''),
    jsonOutput: boolArg(args.get('json'), false),
    includeArtifacts: boolArg(args.get('include-artifacts'), false),
    writeItems: boolArg(args.get('write-items'), true),
    failOnEmpty: boolArg(args.get('fail-on-empty'), true) || boolArg(args.get('fail-on-failed'), false),
    requiredPlatforms: listArg(
      args.get('required-platforms') || args.get('platforms'),
      ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
    ),
  };
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function itemFile(outDir, platform) {
  return outDir ? resolve(outDir, `${platform}.json`) : undefined;
}

async function writePortfolioItems(outDir, portfolio = {}, options = {}) {
  if (!outDir || options.writeItems === false) return [];
  const writtenFiles = [];
  for (const item of portfolio.items ?? []) {
    const file = itemFile(outDir, item.platform);
    await writeJson(file, item);
    writtenFiles.push(file);
  }
  return writtenFiles;
}

export async function buildMeetingPlatformAdapterPortfolioCliReport(options = {}) {
  const {
    baseUrl = 'http://localhost:8787',
    outDir = '',
    reportFile = '',
    includeArtifacts = false,
    writeItems = true,
    requiredPlatforms = ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  } = options;
  const portfolio = buildMeetingPlatformAdapterPortfolio({
    ...options,
    baseUrl,
    env: options.env ?? process.env,
    platforms: requiredPlatforms,
    includeArtifacts,
  });
  const writtenFiles = await writePortfolioItems(outDir, portfolio, { writeItems });
  const report = {
    type: 'meeting_platform_adapter_portfolio_report',
    ok: portfolio.platform_count > 0,
    requirement: 'meeting_platform_adapter_portfolio_available_for_meeting_software_selection_and_handoff',
    base_url: baseUrl,
    out_dir: outDir || undefined,
    write_items: writeItems,
    include_artifacts: includeArtifacts,
    platform_count: portfolio.platform_count,
    built_in_count: portfolio.built_in_count,
    external_authoring_count: portfolio.external_authoring_count,
    browser_surface_ready_count: portfolio.browser_surface_ready_count,
    provider_reconcile_count: portfolio.provider_reconcile_count,
    implementation_ready_count: portfolio.implementation_ready_count,
    pilot_ready_count: portfolio.pilot_ready_count,
    production_ready_count: portfolio.production_ready_count,
    required_platforms: requiredPlatforms,
    written_files: writtenFiles,
    rows: portfolio.rows.map((row) => ({
      ...row,
      item_file: itemFile(outDir, row.platform),
    })),
    portfolio: includeArtifacts ? portfolio : stripArtifacts(portfolio),
    next_actions: portfolio.next_actions ?? [],
  };
  if (reportFile) await writeJson(resolve(reportFile), report);
  return report;
}

export function formatMeetingPlatformAdapterPortfolioCliReport(report = {}) {
  const lines = [
    `meeting_platform_adapter_portfolio_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | builtin=${report.built_in_count} | external=${report.external_authoring_count} | browser_ready=${report.browser_surface_ready_count} | provider=${report.provider_reconcile_count} | pilot=${report.pilot_ready_count} | written=${report.written_files?.length ?? 0}`,
  ];
  for (const row of report.rows ?? []) {
    lines.push(`${row.platform}: status=${row.adapter_status} surface=${row.recommended_first_surface} provider=${row.provider_path} docs=${row.official_doc_count} pilot=${boolLabel(row.pilot_ready)} next=${row.next_action} item=${row.item_file ?? 'n/a'}`);
  }
  if (report.next_actions?.length > 0) lines.push(`next_actions=${report.next_actions.join(',')}`);
  if (report.out_dir) lines.push(`out_dir=${basename(report.out_dir)}`);
  return lines.join('\n');
}

export async function runMeetingPlatformAdapterPortfolioCli(argv = process.argv.slice(2), io = console) {
  const options = meetingPlatformAdapterPortfolioCliOptionsFromArgs(
    parseMeetingPlatformAdapterPortfolioCliArgs(argv),
  );
  const report = await buildMeetingPlatformAdapterPortfolioCliReport(options);
  if (options.jsonOutput) {
    io.log(JSON.stringify(report, null, 2));
  } else {
    io.log(formatMeetingPlatformAdapterPortfolioCliReport(report));
  }
  if (!report.ok && options.failOnEmpty) process.exitCode = 2;
  return report;
}
