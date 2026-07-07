import { mkdir, writeFile } from 'node:fs/promises';
import { basename, dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdapterBlueprintMatrix,
} from '../adapters/platform-adapter-blueprint.mjs';

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

export function parseMeetingPlatformAdapterBlueprintCliArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

export function meetingPlatformAdapterBlueprintCliOptionsFromArgs(args = new Map()) {
  return {
    baseUrl: String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787'),
    outDir: String(args.get('out-dir') || args.get('outDir') || ''),
    reportFile: String(args.get('report-file') || args.get('write-report') || ''),
    jsonOutput: boolArg(args.get('json'), false),
    includeBlueprints: boolArg(args.get('include-blueprints'), false),
    writeBlueprints: boolArg(args.get('write-blueprints'), true),
    failOnBlocked: boolArg(args.get('fail-on-blocked'), false) || boolArg(args.get('fail-on-failed'), false),
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

function blueprintFile(outDir, platform) {
  return outDir ? resolve(outDir, platform, 'adapter-blueprint.json') : undefined;
}

async function writeBlueprintFiles(outDir, blueprints = [], options = {}) {
  if (!outDir || options.writeBlueprints === false) return [];
  const writtenFiles = [];
  for (const blueprint of blueprints) {
    const file = blueprintFile(outDir, blueprint.platform);
    await writeJson(file, blueprint);
    writtenFiles.push(file);
  }
  return writtenFiles;
}

export async function buildMeetingPlatformAdapterBlueprintCliReport(options = {}) {
  const {
    baseUrl = 'http://localhost:8787',
    outDir = '',
    reportFile = '',
    includeBlueprints = false,
    writeBlueprints = true,
    requiredPlatforms = ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  } = options;
  const matrix = buildMeetingPlatformAdapterBlueprintMatrix({
    baseUrl,
    platforms: requiredPlatforms,
  });
  const writtenFiles = await writeBlueprintFiles(outDir, matrix.blueprints, { writeBlueprints });
  const blocked = matrix.blueprints.filter((blueprint) => blueprint.readiness?.ready !== true);
  const report = {
    type: 'meeting_platform_adapter_blueprint_report',
    ok: matrix.platform_count > 0 && blocked.length === 0,
    requirement: 'adapter_blueprints_keep_local_realtime_axis_and_provider_nonblocking',
    base_url: baseUrl,
    out_dir: outDir || undefined,
    write_blueprints: writeBlueprints,
    platform_count: matrix.platform_count,
    ready_count: matrix.ready_count,
    browser_primary_count: matrix.browser_primary_count,
    native_primary_count: matrix.native_primary_count,
    provider_non_blocking_count: matrix.provider_non_blocking_count,
    transcript_non_blocking_count: matrix.transcript_non_blocking_count,
    blocking_count: blocked.length,
    blocking_platforms: blocked.map((blueprint) => blueprint.platform),
    required_platforms: requiredPlatforms,
    written_files: writtenFiles,
    rows: matrix.rows.map((row) => ({
      ...row,
      blueprint_file: blueprintFile(outDir, row.platform),
    })),
    matrix: includeBlueprints ? matrix : {
      ...matrix,
      blueprints: undefined,
    },
    next_actions: unique(matrix.blueprints.flatMap((blueprint) => blueprint.next_actions ?? [])),
  };
  if (reportFile) await writeJson(resolve(reportFile), report);
  return report;
}

export function formatMeetingPlatformAdapterBlueprintCliReport(report = {}) {
  const lines = [
    `meeting_platform_adapter_blueprint_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | ready=${report.ready_count} | browser_primary=${report.browser_primary_count} | native_primary=${report.native_primary_count} | provider_nonblocking=${report.provider_non_blocking_count} | transcript_nonblocking=${report.transcript_non_blocking_count} | blocking=${report.blocking_count} | written=${report.written_files?.length ?? 0}`,
  ];
  for (const row of report.rows ?? []) {
    lines.push(`${row.platform}: ready=${boolLabel(row.ready)} primary=${row.primary_surface ?? 'none'} browser=${boolLabel(row.browser_recommended)} native=${boolLabel(row.native_recommended)} provider_blocks=${boolLabel(row.provider_blocks_realtime)} start=${row.start_create_on ?? 'n/a'} file=${row.blueprint_file ?? 'n/a'}`);
  }
  if (report.next_actions?.length > 0) lines.push(`next_actions=${report.next_actions.join(',')}`);
  if (report.out_dir) lines.push(`out_dir=${basename(report.out_dir)}`);
  return lines.join('\n');
}

export async function runMeetingPlatformAdapterBlueprintCli(argv = process.argv.slice(2), io = console) {
  const options = meetingPlatformAdapterBlueprintCliOptionsFromArgs(
    parseMeetingPlatformAdapterBlueprintCliArgs(argv),
  );
  const report = await buildMeetingPlatformAdapterBlueprintCliReport(options);
  if (options.jsonOutput) {
    io.log(JSON.stringify(report, null, 2));
  } else {
    io.log(formatMeetingPlatformAdapterBlueprintCliReport(report));
  }
  if (!report.ok && options.failOnBlocked) process.exitCode = 2;
  return report;
}
