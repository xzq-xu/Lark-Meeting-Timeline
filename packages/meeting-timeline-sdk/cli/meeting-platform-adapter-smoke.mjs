import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { basename, dirname, resolve } from 'node:path';
import {
  runMeetingPlatformAdapterSmoke,
} from '../adapters/platform-adapter-smoke.mjs';

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

export function parseMeetingPlatformAdapterSmokeCliArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

export function meetingPlatformAdapterSmokeCliOptionsFromArgs(args = new Map()) {
  return {
    manifestFile: String(args.get('manifest-file') || args.get('manifestFile') || args.get('install-manifest') || ''),
    outFile: String(args.get('out-file') || args.get('outFile') || ''),
    reportFile: String(args.get('report-file') || args.get('write-report') || ''),
    jsonOutput: args.get('json') === 'true',
    failOnBlocked: boolArg(args.get('fail-on-blocked'), false) || boolArg(args.get('fail-on-failed'), false),
    platforms: listArg(args.get('platforms') || args.get('platform-keys'), []),
    target: String(args.get('target') || args.get('acceptance-target') || 'static'),
    baseUrl: String(args.get('base-url') || args.get('baseUrl') || ''),
    capturedAtMs: args.has('captured-at-ms') ? Number(args.get('captured-at-ms')) : undefined,
  };
}

export async function buildMeetingPlatformAdapterSmokeCliReport(options = {}) {
  const {
    manifestFile = '',
    outFile = '',
    reportFile = '',
    platforms = [],
    target = 'static',
    baseUrl = '',
    capturedAtMs,
  } = options;
  let manifestOrOptions = {
    platforms: platforms.length ? platforms : undefined,
    target,
    baseUrl: baseUrl || undefined,
    capturedAtMs,
  };
  let manifestReadError;
  let generatedManifest = true;
  if (manifestFile) {
    try {
      manifestOrOptions = await readJson(manifestFile);
      generatedManifest = false;
    } catch (error) {
      manifestReadError = error.message;
      generatedManifest = false;
    }
  }
  const smokeReport = manifestReadError
    ? {
      schema: 'meeting_platform_adapter_smoke_report',
      accepted: false,
      platform_count: 0,
      accepted_count: 0,
      failed_count: 0,
      platforms,
      rows: [],
      next_actions: ['provide_adapter_install_manifest_or_omit_manifest_file_to_generate_static_manifest'],
    }
    : await runMeetingPlatformAdapterSmoke(manifestOrOptions, {
      platforms: platforms.length ? platforms : undefined,
      target,
      baseUrl: baseUrl || undefined,
      capturedAtMs,
    });
  const report = {
    type: 'meeting_platform_adapter_smoke_cli_report',
    ok: smokeReport.accepted === true && !manifestReadError,
    target,
    base_url: baseUrl || undefined,
    manifest_file: manifestFile || undefined,
    generated_manifest: generatedManifest,
    manifest_read_error: manifestReadError,
    out_file: outFile || undefined,
    platform_count: smokeReport.platform_count ?? 0,
    accepted_count: smokeReport.accepted_count ?? 0,
    failed_count: smokeReport.failed_count ?? 0,
    install_manifest_accepted: smokeReport.install_manifest_accepted === true,
    platforms: smokeReport.platforms ?? platforms,
    rows: smokeReport.rows ?? [],
    next_actions: smokeReport.next_actions ?? [],
    smoke_report: smokeReport,
  };
  if (outFile) await writeJson(resolve(outFile), smokeReport);
  if (reportFile) await writeJson(resolve(reportFile), report);
  return report;
}

export function formatMeetingPlatformAdapterSmokeCliReport(report = {}) {
  const lines = [
    `meeting_platform_adapter_smoke_cli_report | ok=${boolLabel(report.ok)} | target=${report.target ?? 'static'} | platforms=${report.platform_count ?? 0} | accepted=${report.accepted_count ?? 0} | failed=${report.failed_count ?? 0} | generated_manifest=${boolLabel(report.generated_manifest)}`,
  ];
  for (const row of report.rows ?? []) {
    lines.push(`${row.platform}: accepted=${boolLabel(row.accepted)} surface=${row.selected_surface ?? 'n/a'} blueprint_surface=${row.adapter_blueprint_primary_surface ?? 'n/a'} observe_before_insert=${boolLabel(row.observe_before_insert)} captured_at_ms=${boolLabel(row.captured_at_ms_preserved)} speaker=${boolLabel(row.speaker_track_inserted)} participant=${boolLabel(row.participant_track_inserted)} provider_nonblocking=${boolLabel(row.provider_reconcile_nonblocking)} issues=${row.issues?.join(',') || 'none'}`);
  }
  if (report.manifest_file) lines.push(`manifest_file=${report.manifest_file}`);
  if (report.manifest_read_error) lines.push(`manifest_read_error=${report.manifest_read_error}`);
  if (report.next_actions?.length > 0) lines.push(`next_actions=${report.next_actions.join(',')}`);
  if (report.out_file) lines.push(`out_file=${basename(report.out_file)}`);
  return lines.join('\n');
}

export async function runMeetingPlatformAdapterSmokeCli(argv = process.argv.slice(2), io = console) {
  const options = meetingPlatformAdapterSmokeCliOptionsFromArgs(
    parseMeetingPlatformAdapterSmokeCliArgs(argv),
  );
  const report = await buildMeetingPlatformAdapterSmokeCliReport(options);
  if (options.jsonOutput) {
    io.log(JSON.stringify(report, null, 2));
  } else {
    io.log(formatMeetingPlatformAdapterSmokeCliReport(report));
  }
  if (!report.ok && options.failOnBlocked) process.exitCode = 2;
  return report;
}
