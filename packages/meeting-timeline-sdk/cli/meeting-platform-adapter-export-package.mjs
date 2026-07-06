import { mkdir, writeFile } from 'node:fs/promises';
import { basename, dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdapterExportPackageMatrix,
} from '../adapters/platform-adapter-export-package.mjs';

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

export function parseMeetingPlatformAdapterExportPackageCliArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

export function meetingPlatformAdapterExportPackageCliOptionsFromArgs(args = new Map()) {
  return {
    baseUrl: String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787'),
    outDir: String(args.get('out-dir') || args.get('outDir') || ''),
    reportFile: String(args.get('report-file') || args.get('write-report') || ''),
    target: String(args.get('target') || args.get('acceptance-target') || 'static'),
    jsonOutput: args.get('json') === 'true',
    includePackages: boolArg(args.get('include-packages'), false),
    includeArtifactsInPackage: boolArg(args.get('include-artifacts-in-package'), false),
    writeArtifacts: boolArg(args.get('write-artifacts'), true),
    failOnBlocked: args.get('fail-on-blocked') === 'true' || args.get('fail-on-failed') === 'true',
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

function stripArtifacts(pkg = {}, includeArtifactsInPackage = false) {
  if (includeArtifactsInPackage) return pkg;
  return {
    ...pkg,
    authoring_plan: undefined,
    artifacts: undefined,
  };
}

function artifactFor(pkg = {}, source) {
  if (source === 'adapter_export_package') return pkg;
  return pkg.artifacts?.[source];
}

async function writeExportPackageFiles(outDir, pkg = {}, options = {}) {
  if (!outDir) return [];
  const root = resolve(outDir);
  const writtenFiles = [];
  const mainFile = resolve(root, pkg.platform, 'adapter-export-package.json');
  await writeJson(mainFile, stripArtifacts(pkg, options.includeArtifactsInPackage));
  writtenFiles.push(mainFile);
  if (options.writeArtifacts === false) return writtenFiles;
  for (const hostFile of pkg.host_files ?? []) {
    if (hostFile.source === 'adapter_export_package') continue;
    const artifact = artifactFor(pkg, hostFile.source);
    if (!artifact) continue;
    const file = resolve(root, hostFile.path);
    await writeJson(file, artifact);
    writtenFiles.push(file);
  }
  return writtenFiles;
}

export async function buildMeetingPlatformAdapterExportPackageCliReport(options = {}) {
  const {
    baseUrl = 'http://localhost:8787',
    outDir = '',
    reportFile = '',
    target = 'static',
    includePackages = false,
    includeArtifactsInPackage = false,
    writeArtifacts = true,
    requiredPlatforms = ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  } = options;
  const matrix = buildMeetingPlatformAdapterExportPackageMatrix({
    platforms: requiredPlatforms,
  }, {
    baseUrl,
    target,
    includeArtifacts: true,
  });
  const writtenFiles = [];
  for (const pkg of matrix.packages) {
    writtenFiles.push(...await writeExportPackageFiles(outDir, pkg, {
      includeArtifactsInPackage,
      writeArtifacts,
    }));
  }
  const report = {
    type: 'meeting_platform_adapter_export_package_report',
    ok: matrix.platform_count > 0,
    target,
    base_url: baseUrl,
    out_dir: outDir || undefined,
    platform_count: matrix.platform_count,
    accepted_count: matrix.accepted_count,
    export_ready_count: matrix.export_ready_count,
    built_in_count: matrix.built_in_count,
    custom_authoring_count: matrix.custom_authoring_count,
    browser_extension_ready_count: matrix.browser_extension_ready_count,
    provider_reconcile_count: matrix.provider_reconcile_count,
    required_platforms: requiredPlatforms,
    written_files: writtenFiles,
    rows: matrix.rows.map((row) => ({
      ...row,
      package_file: outDir ? resolve(outDir, row.platform, 'adapter-export-package.json') : undefined,
    })),
    matrix: includePackages ? matrix : {
      ...matrix,
      packages: undefined,
      portfolio: undefined,
      acceptance_checklist_matrix: undefined,
    },
    next_actions: matrix.next_actions,
  };
  if (reportFile) await writeJson(resolve(reportFile), report);
  return report;
}

export function formatMeetingPlatformAdapterExportPackageCliReport(report = {}) {
  const lines = [
    `meeting_platform_adapter_export_package_report | ok=${boolLabel(report.ok)} | target=${report.target} | platforms=${report.platform_count} | accepted=${report.accepted_count} | export_ready=${report.export_ready_count} | browser_ready=${report.browser_extension_ready_count} | providers=${report.provider_reconcile_count} | written=${report.written_files?.length ?? 0}`,
  ];
  for (const row of report.rows ?? []) {
    lines.push(`${row.platform}: accepted=${boolLabel(row.accepted)} export_ready=${boolLabel(row.export_ready)} built_in=${boolLabel(row.built_in)} surface=${row.recommended_first_surface ?? 'none'} browser=${boolLabel(row.browser_extension_ready)} provider=${boolLabel(row.provider_reconcile_ready)} files=${row.host_file_count} next=${row.first_next_action ?? 'none'} package=${row.package_file ?? 'n/a'}`);
  }
  if (report.next_actions?.length > 0) lines.push(`next_actions=${report.next_actions.join(',')}`);
  if (report.out_dir) lines.push(`out_dir=${basename(report.out_dir)}`);
  return lines.join('\n');
}

export async function runMeetingPlatformAdapterExportPackageCli(argv = process.argv.slice(2), io = console) {
  const options = meetingPlatformAdapterExportPackageCliOptionsFromArgs(
    parseMeetingPlatformAdapterExportPackageCliArgs(argv),
  );
  const report = await buildMeetingPlatformAdapterExportPackageCliReport(options);
  if (options.jsonOutput) {
    io.log(JSON.stringify(report, null, 2));
  } else {
    io.log(formatMeetingPlatformAdapterExportPackageCliReport(report));
  }
  if (report.accepted_count !== report.platform_count && options.failOnBlocked) process.exitCode = 2;
  return report;
}
