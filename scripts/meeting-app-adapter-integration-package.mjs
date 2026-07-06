#!/usr/bin/env node

import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { basename, dirname, resolve } from 'node:path';
import {
  buildMeetingAppAdapterIntegrationPackageMatrix,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-adapter-integration-package.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const outDir = String(args.get('out-dir') || args.get('outDir') || '');
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const includePackages = args.get('include-packages') === 'true';
const includePackageDescriptor = args.get('include-package-descriptor') !== 'false';
const failOnRejected = args.get('fail-on-rejected') === 'true' || args.get('fail-on-failed') === 'true';
const target = String(
  args.get('target')
    || args.get('acceptance-target')
    || args.get('acceptanceTarget')
    || (args.get('require-production-ready') === 'true' ? 'production' : '')
    || (args.get('require-realtime-ready') === 'true' ? 'realtime' : '')
    || 'pilot',
);
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()).filter(Boolean));

async function readJsonArg(name) {
  const value = args.get(name);
  if (!value) return undefined;
  return JSON.parse(await readFile(resolve(String(value)), 'utf8'));
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

async function writeContent(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, typeof value === 'string' ? value : `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function packageDir(pkg = {}) {
  return resolve(outDir, pkg.platform || 'adapter');
}

function targetAccepted(pkg = {}) {
  switch (target) {
    case 'static':
      return pkg.static_ready === true;
    case 'realtime':
      return pkg.accepted === true && pkg.realtime_ready === true;
    case 'production':
      return pkg.production_ready === true;
    case 'pilot':
    default:
      return pkg.accepted === true && pkg.pilot_ready === true;
  }
}

function stripFileContents(files = []) {
  return files.map(({ content, ...file }) => file);
}

function descriptor(pkg = {}, writtenFiles = []) {
  return {
    ...pkg,
    files: stripFileContents(pkg.files),
    handoff_package: pkg.handoff_package
      ? {
        ...pkg.handoff_package,
        files: stripFileContents(pkg.handoff_package.files),
      }
      : undefined,
    written_files: writtenFiles,
  };
}

const inputs = await readJsonArg('input-file') ?? await readJsonArg('snapshot-file');
const evidenceByAdapter = await readJsonArg('evidence-file');
const matrix = buildMeetingAppAdapterIntegrationPackageMatrix({
  baseUrl,
  platforms: requiredPlatforms,
  inputs,
  evidenceByAdapter,
  includeVerification: evidenceByAdapter != null || args.get('include-verification') === 'true',
});

const writtenFiles = [];
if (outDir) {
  for (const pkg of matrix.packages) {
    const pkgDir = packageDir(pkg);
    const packageWrittenFiles = [];
    for (const file of pkg.files) {
      const outFile = resolve(pkgDir, file.path);
      await writeContent(outFile, file.content);
      writtenFiles.push(outFile);
      packageWrittenFiles.push(outFile);
    }
    if (includePackageDescriptor) {
      const descriptorFile = resolve(pkgDir, 'integration-package.json');
      await writeJson(descriptorFile, descriptor(pkg, packageWrittenFiles));
      writtenFiles.push(descriptorFile);
    }
  }
}

const acceptedCountForTarget = matrix.packages.filter((pkg) => targetAccepted(pkg)).length;
const report = {
  type: 'meeting_app_adapter_integration_package_report',
  ok: acceptedCountForTarget === matrix.platform_count,
  target,
  base_url: baseUrl,
  out_dir: outDir || undefined,
  platform_count: matrix.platform_count,
  accepted_count: matrix.accepted_count,
  target_accepted_count: acceptedCountForTarget,
  static_ready_count: matrix.static_ready_count,
  pilot_ready_count: matrix.pilot_ready_count,
  realtime_ready_count: matrix.realtime_ready_count,
  production_ready_count: matrix.production_ready_count,
  runtime_ready_count: matrix.packages.filter((pkg) => pkg.runtime_delivery?.readiness?.runtime_ready === true).length,
  local_observer_first_count: matrix.packages.filter((pkg) => pkg.runtime_delivery?.adapter_route?.local_observer_first === true).length,
  required_platforms: requiredPlatforms,
  written_files: writtenFiles,
  rows: matrix.rows.map((row) => {
    const pkg = matrix.packages.find((item) => item.platform === row.platform);
    return {
      ...row,
      target_accepted: pkg ? targetAccepted(pkg) : false,
      runtime_ready: pkg?.runtime_delivery?.readiness?.runtime_ready === true,
      first_runtime_route: pkg?.runtime_delivery?.adapter_route?.first_route,
      package_dir: outDir ? packageDir(row) : undefined,
    };
  }),
  matrix: includePackages ? matrix : {
    ...matrix,
    packages: undefined,
  },
  next_actions: matrix.next_actions,
};

if (reportFile) await writeJson(resolve(reportFile), report);

if (jsonOutput) {
  console.log(JSON.stringify(report, null, 2));
} else {
  console.log(`meeting_app_adapter_integration_package_report | ok=${boolLabel(report.ok)} | target=${target} | platforms=${report.platform_count} | accepted=${report.accepted_count} | target_accepted=${report.target_accepted_count} | pilot=${report.pilot_ready_count} | realtime=${report.realtime_ready_count} | production=${report.production_ready_count} | runtime_ready=${report.runtime_ready_count} | written=${report.written_files.length}`);
  for (const row of report.rows) {
    console.log(`${row.platform}: target=${boolLabel(row.target_accepted)} pilot=${boolLabel(row.pilot_ready)} realtime=${boolLabel(row.realtime_ready)} production=${boolLabel(row.production_ready)} runtime=${boolLabel(row.runtime_ready)} route=${row.first_runtime_route ?? 'none'} next=${row.first_next_action ?? 'none'} package=${row.package_dir ? basename(row.package_dir) : 'n/a'}`);
  }
  if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
}

if (!report.ok && failOnRejected) process.exitCode = 2;
