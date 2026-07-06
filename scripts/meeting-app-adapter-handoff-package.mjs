#!/usr/bin/env node

import { mkdir, readFile, readdir, stat, writeFile } from 'node:fs/promises';
import { basename, dirname, extname, join, resolve } from 'node:path';
import {
  buildMeetingAppAdapterHandoffPackageMatrix,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-adapter-handoff-package.mjs';
import {
  buildMeetingAppAdapterSpecTemplate,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-adapter-spec.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const outDir = String(args.get('out-dir') || args.get('outDir') || '');
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const templateFile = String(args.get('template-file') || args.get('templateFile') || '');
const jsonOutput = args.get('json') === 'true';
const includePackages = args.get('include-packages') === 'true';
const includeBuiltIns = args.get('include-built-ins') !== 'false' && args.get('no-built-ins') !== 'true';
const includeTemplate = args.get('include-template') !== 'false';
const includePackageDescriptor = args.get('include-package-descriptor') !== 'false';
const failOnRejected = args.get('fail-on-rejected') === 'true' || args.get('fail-on-failed') === 'true';
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()).filter(Boolean));

function splitList(value) {
  return unique(String(value || '').split(',').map((item) => item.trim()).filter(Boolean));
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

async function writeContent(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, typeof value === 'string' ? value : `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

async function readJson(file) {
  return JSON.parse(await readFile(file, 'utf8'));
}

function specEntriesFromJson(value) {
  if (Array.isArray(value)) return value;
  if (Array.isArray(value?.specs)) return value.specs;
  if (Array.isArray(value?.adapters)) return value.adapters;
  if (Array.isArray(value?.adapter_specs)) return value.adapter_specs;
  if (Array.isArray(value?.adapterSpecs)) return value.adapterSpecs;
  return value ? [value] : [];
}

async function specFilesFromDir(dir) {
  const entries = await readdir(dir);
  const files = [];
  for (const entry of entries) {
    const file = join(dir, entry);
    const itemStat = await stat(file);
    if (itemStat.isFile() && extname(file).toLowerCase() === '.json') files.push(file);
  }
  return files.sort();
}

async function loadCustomSpecs() {
  const files = [
    ...splitList(args.get('spec-file') || args.get('specFile')),
    ...splitList(args.get('spec-files') || args.get('specFiles')),
  ];
  const dirs = splitList(args.get('spec-dir') || args.get('specDir'));
  for (const dir of dirs) files.push(...await specFilesFromDir(resolve(dir)));
  const specs = [];
  const loadedFiles = [];
  for (const file of files) {
    const resolved = resolve(file);
    const value = await readJson(resolved);
    specs.push(...specEntriesFromJson(value));
    loadedFiles.push(resolved);
  }
  return { specs, loadedFiles };
}

function templateInput() {
  const key = args.get('template-adapter-key') || args.get('templateAdapterKey') || args.get('template');
  if (!key) return null;
  const matches = splitList(args.get('template-matches') || args.get('templateMatches') || args.get('matches') || args.get('url-matches'));
  const hostPermissions = splitList(args.get('template-host-permissions') || args.get('templateHostPermissions') || args.get('host-permissions'));
  return {
    adapter_key: key,
    display_name: args.get('template-display-name') || args.get('templateDisplayName') || args.get('display-name') || args.get('displayName'),
    matches: matches.length ? matches : undefined,
    host_permissions: hostPermissions.length ? hostPermissions : undefined,
  };
}

function packageDir(pkg = {}) {
  return resolve(outDir, pkg.adapter_key || 'adapter');
}

function descriptor(pkg = {}, writtenFiles = []) {
  return {
    ...pkg,
    files: pkg.files.map(({ content, ...file }) => file),
    adapter_spec: undefined,
    adapter_manifest: undefined,
    runtime_config: undefined,
    extension_manifest_fragment: undefined,
    written_files: writtenFiles,
  };
}

const custom = await loadCustomSpecs();
const template = templateInput() ? buildMeetingAppAdapterSpecTemplate(templateInput()) : null;
if (templateFile && template) await writeJson(resolve(templateFile), template);

const specs = [
  ...custom.specs,
  ...(template && includeTemplate ? [template] : []),
];
const matrix = buildMeetingAppAdapterHandoffPackageMatrix({
  baseUrl,
  platforms: includeBuiltIns ? requiredPlatforms : [],
  specs,
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
      const descriptorFile = resolve(pkgDir, 'handoff-package.json');
      await writeJson(descriptorFile, descriptor(pkg, packageWrittenFiles));
      writtenFiles.push(descriptorFile);
    }
  }
}

const report = {
  type: 'meeting_app_adapter_handoff_package_report',
  ok: matrix.accepted === true,
  base_url: baseUrl,
  out_dir: outDir || undefined,
  package_count: matrix.package_count,
  accepted_count: matrix.accepted_count,
  custom_count: matrix.custom_count,
  built_in_count: matrix.built_in_count,
  runtime_config_ready_count: matrix.runtime_config_ready_count,
  content_script_ready_count: matrix.content_script_ready_count,
  live_evidence_required_count: matrix.live_evidence_required_count,
  required_platforms: includeBuiltIns ? requiredPlatforms : [],
  loaded_spec_files: custom.loadedFiles,
  template_file: templateFile && template ? resolve(templateFile) : undefined,
  written_files: writtenFiles,
  rows: matrix.rows.map((row) => ({
    ...row,
    package_dir: outDir ? packageDir(row) : undefined,
  })),
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
  console.log(`meeting_app_adapter_handoff_package_report | ok=${boolLabel(report.ok)} | packages=${report.package_count} | accepted=${report.accepted_count} | built_in=${report.built_in_count} | custom=${report.custom_count} | runtime_ready=${report.runtime_config_ready_count} | content_script_ready=${report.content_script_ready_count} | live_evidence_required=${report.live_evidence_required_count} | written=${report.written_files.length}`);
  for (const row of report.rows) {
    console.log(`${row.adapter_key}: accepted=${boolLabel(row.accepted)} source=${row.source} files=${row.file_count} runtime=${boolLabel(row.runtime_config_accepted)} content_script=${boolLabel(row.content_script_ready)} capture=${boolLabel(row.capture_ready)} mutation=${boolLabel(row.mutation_ready)} next=${row.first_next_action ?? 'none'} package=${row.package_dir ? basename(row.package_dir) : 'n/a'}`);
  }
  if (report.template_file) console.log(`template_file=${report.template_file}`);
  if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
}

if (!report.ok && failOnRejected) process.exitCode = 2;
