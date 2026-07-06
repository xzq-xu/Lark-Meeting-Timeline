#!/usr/bin/env node

import { mkdir, readFile, readdir, stat, writeFile } from 'node:fs/promises';
import { dirname, extname, join, resolve } from 'node:path';
import {
  buildMeetingAppAdapterVerificationReportMatrix,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-adapter-handoff-package.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const includeReports = args.get('include-reports') === 'true';
const failOnRejected = args.get('fail-on-rejected') === 'true' || args.get('fail-on-failed') === 'true';
const target = String(args.get('target') || args.get('acceptance-target') || args.get('acceptanceTarget') || 'production');
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || '',
).split(',').map((item) => item.trim()).filter(Boolean));

function splitList(value) {
  return unique(String(value || '').split(',').map((item) => item.trim()).filter(Boolean));
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

async function readJson(file) {
  return JSON.parse(await readFile(file, 'utf8'));
}

function keyFromValue(value = {}, fallback = '') {
  return String(value.adapter_key ?? value.platform ?? value.provider ?? value.key ?? fallback)
    .trim()
    .replace(/-/g, '_');
}

function specEntriesFromJson(value) {
  if (Array.isArray(value)) return value;
  if (Array.isArray(value?.specs)) return value.specs;
  if (Array.isArray(value?.adapters)) return value.adapters;
  if (Array.isArray(value?.adapter_specs)) return value.adapter_specs;
  if (Array.isArray(value?.adapterSpecs)) return value.adapterSpecs;
  if (Array.isArray(value?.packages)) return value.packages;
  return value ? [value] : [];
}

async function jsonFilesFromDir(dir) {
  const entries = await readdir(dir);
  const files = [];
  for (const entry of entries) {
    const file = join(dir, entry);
    const itemStat = await stat(file);
    if (itemStat.isFile() && extname(file).toLowerCase() === '.json') files.push(file);
  }
  return files.sort();
}

async function packageSpecsFromDir(dir) {
  const entries = await readdir(dir);
  const specs = [];
  const files = [];
  for (const entry of entries) {
    const child = join(dir, entry);
    const itemStat = await stat(child);
    if (itemStat.isDirectory()) {
      const specFile = join(child, 'adapter-spec.json');
      try {
        await stat(specFile);
        specs.push(await readJson(specFile));
        files.push(specFile);
      } catch {
        // Ignore directories that are not handoff packages.
      }
    } else if (itemStat.isFile() && extname(child).toLowerCase() === '.json') {
      const value = await readJson(child);
      specs.push(...specEntriesFromJson(value));
      files.push(child);
    }
  }
  return { specs, files };
}

async function loadSpecs() {
  const specs = [];
  const files = [];
  for (const dir of splitList(args.get('package-dir') || args.get('packageDir'))) {
    const loaded = await packageSpecsFromDir(resolve(dir));
    specs.push(...loaded.specs);
    files.push(...loaded.files);
  }
  for (const file of [
    ...splitList(args.get('package-file') || args.get('packageFile')),
    ...splitList(args.get('spec-file') || args.get('specFile')),
    ...splitList(args.get('spec-files') || args.get('specFiles')),
  ]) {
    const resolved = resolve(file);
    const value = await readJson(resolved);
    specs.push(...specEntriesFromJson(value));
    files.push(resolved);
  }
  for (const dir of splitList(args.get('spec-dir') || args.get('specDir'))) {
    for (const file of await jsonFilesFromDir(resolve(dir))) {
      const value = await readJson(file);
      specs.push(...specEntriesFromJson(value));
      files.push(file);
    }
  }
  return { specs, files };
}

async function loadEvidenceByAdapter() {
  const evidenceByAdapter = {};
  const files = [
    ...splitList(args.get('evidence-file') || args.get('evidenceFile')),
    ...splitList(args.get('evidence-files') || args.get('evidenceFiles')),
  ];
  for (const dir of splitList(args.get('evidence-dir') || args.get('evidenceDir'))) {
    files.push(...await jsonFilesFromDir(resolve(dir)));
  }
  const loadedFiles = [];
  for (const file of files) {
    const resolved = resolve(file);
    const value = await readJson(resolved);
    const entries = Array.isArray(value) ? value : specEntriesFromJson(value.evidence_by_adapter ?? value.evidenceByAdapter ?? value);
    if (value.evidence_by_adapter || value.evidenceByAdapter) {
      for (const [key, evidence] of Object.entries(value.evidence_by_adapter ?? value.evidenceByAdapter)) {
        evidenceByAdapter[key.replace(/-/g, '_')] = evidence;
      }
    } else {
      for (const evidence of entries) {
        evidenceByAdapter[keyFromValue(evidence, resolved.split('/').pop()?.replace(/\.json$/i, ''))] = evidence;
      }
    }
    loadedFiles.push(resolved);
  }
  return { evidenceByAdapter, files: loadedFiles };
}

const loadedSpecs = await loadSpecs();
const loadedEvidence = await loadEvidenceByAdapter();
const matrix = buildMeetingAppAdapterVerificationReportMatrix({
  target,
  platforms: requiredPlatforms,
  specs: loadedSpecs.specs,
  evidenceByAdapter: loadedEvidence.evidenceByAdapter,
});

const report = {
  type: 'meeting_app_adapter_verification_cli_report',
  ok: matrix.accepted === true,
  target,
  report_count: matrix.report_count,
  accepted_count: matrix.accepted_count,
  pilot_ready_count: matrix.pilot_ready_count,
  production_ready_count: matrix.production_ready_count,
  missing_evidence_count: matrix.missing_evidence_count,
  required_platforms: requiredPlatforms,
  loaded_spec_files: loadedSpecs.files,
  loaded_evidence_files: loadedEvidence.files,
  rows: matrix.rows,
  matrix: includeReports ? matrix : {
    ...matrix,
    reports: undefined,
  },
  next_actions: matrix.next_actions,
};

if (reportFile) await writeJson(resolve(reportFile), report);

if (jsonOutput) {
  console.log(JSON.stringify(report, null, 2));
} else {
  console.log(`meeting_app_adapter_verification_report | ok=${boolLabel(report.ok)} | target=${target} | accepted=${report.accepted_count}/${report.report_count} | pilot_ready=${report.pilot_ready_count} | production_ready=${report.production_ready_count} | missing_evidence=${report.missing_evidence_count}`);
  for (const row of report.rows) {
    console.log(`${row.adapter_key}: accepted=${boolLabel(row.accepted)} static=${boolLabel(row.static_ready)} live=${boolLabel(row.live_evidence_ready)} pilot=${boolLabel(row.pilot_ready)} production=${boolLabel(row.production_ready)} missing=${row.missing_evidence_count} next=${row.first_next_action ?? 'none'}`);
  }
  if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
}

if (!report.ok && failOnRejected) process.exitCode = 2;
