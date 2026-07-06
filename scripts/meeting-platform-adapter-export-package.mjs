#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdapterExportPackageMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-export-package.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const outDir = resolve(String(args.get('out-dir') || args.get('outDir') || 'data/meeting-platform-adapter-export-packages'));
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const target = String(args.get('target') || args.get('acceptance-target') || 'static');
const jsonOutput = args.get('json') === 'true';
const includePackages = args.get('include-packages') === 'true';
const includeArtifactsInPackage = args.get('include-artifacts-in-package') === 'true';
const writePackages = args.get('write-packages') !== 'false';
const writeArtifacts = args.get('write-artifacts') !== 'false';
const failOnBlocked = args.get('fail-on-blocked') === 'true';
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function stripArtifacts(pkg = {}) {
  if (includeArtifactsInPackage) return pkg;
  return {
    ...pkg,
    authoring_plan: undefined,
    artifacts: undefined,
  };
}

function packageFile(platform) {
  return resolve(outDir, platform, 'adapter-export-package.json');
}

function artifactFile(pkg = {}, hostFile = {}) {
  return resolve(outDir, hostFile.path);
}

function artifactFor(pkg = {}, source) {
  if (source === 'adapter_export_package') return stripArtifacts(pkg);
  return pkg.artifacts?.[source];
}

async function writeExportPackage(pkg = {}) {
  const written = [];
  if (!writePackages) return written;
  const mainFile = packageFile(pkg.platform);
  await writeJson(mainFile, stripArtifacts(pkg));
  written.push(mainFile);
  if (!writeArtifacts) return written;
  for (const hostFile of pkg.host_files ?? []) {
    if (hostFile.source === 'adapter_export_package') continue;
    const artifact = artifactFor(pkg, hostFile.source);
    if (!artifact) continue;
    const file = artifactFile(pkg, hostFile);
    await writeJson(file, artifact);
    written.push(file);
  }
  return written;
}

async function buildReport() {
  const matrix = buildMeetingPlatformAdapterExportPackageMatrix({
    platforms: requiredPlatforms,
  }, {
    baseUrl,
    env: process.env,
    target,
    includeArtifacts: true,
  });
  const writtenFiles = [];
  for (const pkg of matrix.packages) {
    writtenFiles.push(...await writeExportPackage(pkg));
  }
  const rows = matrix.rows.map((row) => ({
    ...row,
    package_file: packageFile(row.platform),
  }));
  const blocked = matrix.packages.filter((pkg) => pkg.accepted !== true);
  return {
    type: 'meeting_platform_adapter_export_package_report',
    ok: matrix.platform_count > 0 && (!failOnBlocked || blocked.length === 0),
    requirement: failOnBlocked
      ? `all_adapter_export_packages_pass_${target}_target`
      : 'adapter_export_packages_generated',
    base_url: baseUrl,
    target,
    out_dir: writePackages ? outDir : undefined,
    write_packages: writePackages,
    write_artifacts: writeArtifacts,
    platform_count: matrix.platform_count,
    accepted_count: matrix.accepted_count,
    export_ready_count: matrix.export_ready_count,
    built_in_count: matrix.built_in_count,
    custom_authoring_count: matrix.custom_authoring_count,
    browser_extension_ready_count: matrix.browser_extension_ready_count,
    provider_reconcile_count: matrix.provider_reconcile_count,
    required_platforms: requiredPlatforms,
    written_files: writtenFiles,
    rows,
    matrix: includePackages ? matrix : {
      ...matrix,
      packages: undefined,
      portfolio: undefined,
      acceptance_checklist_matrix: undefined,
    },
    next_actions: matrix.next_actions ?? [],
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_adapter_export_package_report | ok=${boolLabel(report.ok)} | target=${report.target} | platforms=${report.platform_count} | accepted=${report.accepted_count} | export_ready=${report.export_ready_count} | browser_ready=${report.browser_extension_ready_count} | providers=${report.provider_reconcile_count} | written=${report.written_files.length}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: accepted=${boolLabel(row.accepted)} export_ready=${boolLabel(row.export_ready)} built_in=${boolLabel(row.built_in)} surface=${row.recommended_first_surface ?? 'none'} browser=${boolLabel(row.browser_extension_ready)} provider=${boolLabel(row.provider_reconcile_ready)} files=${row.host_file_count} next=${row.first_next_action ?? 'none'} package=${row.package_file}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
  }
  if (!report.ok && failOnBlocked) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
