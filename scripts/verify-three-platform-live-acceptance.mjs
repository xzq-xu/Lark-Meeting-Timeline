#!/usr/bin/env node

import { mkdir, readdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { normalizeMeetingPlatform } from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';
import {
  THREE_PLATFORM_LIVE_ACCEPTANCE_PLATFORMS,
  buildThreePlatformLiveAcceptanceReport,
  evaluateThreePlatformLiveEvidence,
} from './three-platform-live-acceptance-core.mjs';

function parseArgs(argv = process.argv.slice(2)) {
  return new Map(argv.map((raw) => {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    return [key, rest.length ? rest.join('=') : 'true'];
  }));
}

async function jsonFiles(root) {
  const files = [];
  async function walk(dir) {
    let entries = [];
    try { entries = await readdir(dir, { withFileTypes: true }); } catch { return; }
    for (const entry of entries) {
      const path = resolve(dir, entry.name);
      if (entry.isDirectory()) await walk(path);
      else if (entry.isFile() && entry.name.endsWith('.json')) files.push(path);
    }
  }
  await walk(root);
  return files.sort();
}

export async function verifyThreePlatformLiveAcceptance(options = {}) {
  const inputDir = resolve(String(options.inputDir ?? options.input_dir ?? 'data/meeting-platform-field-evidence'));
  const platforms = (options.platforms ?? THREE_PLATFORM_LIVE_ACCEPTANCE_PLATFORMS)
    .map((platform) => normalizeMeetingPlatform(platform));
  const files = await jsonFiles(inputDir);
  const requireSpeaker = options.requireSpeaker === true || options.require_speaker === true;
  const evaluations = [];
  const errors = [];
  for (const file of files) {
    try {
      const input = JSON.parse(await readFile(file, 'utf8'));
      if (input.schema !== 'meeting_platform_field_evidence_input') continue;
      const platform = normalizeMeetingPlatform(input.platform);
      if (!platforms.includes(platform)) continue;
      evaluations.push(evaluateThreePlatformLiveEvidence(input, {
        platform,
        file,
        sinceMs: options.sinceMs ?? options.since_ms,
        annotationSource: options.annotationSource ?? options.annotation_source,
        requireSpeaker,
      }));
    } catch (error) {
      errors.push({ file, error: String(error?.message ?? error) });
    }
  }
  const report = {
    ...buildThreePlatformLiveAcceptanceReport(evaluations, { platforms, requireSpeaker }),
    input_dir: inputDir,
    source_file_count: files.length,
    evaluated_file_count: evaluations.length,
    errors,
  };
  report.accepted = report.accepted && errors.length === 0;
  report.core_accepted = report.core_accepted && errors.length === 0;
  report.production_ready = report.core_accepted;
  report.full_production_ready = report.full_production_ready && errors.length === 0;
  return report;
}

async function main() {
  const args = parseArgs();
  const platforms = String(args.get('platforms') ?? 'google-meet,teams,zoom')
    .split(',')
    .map((value) => value.trim())
    .filter(Boolean);
  const report = await verifyThreePlatformLiveAcceptance({
    inputDir: args.get('dir') ?? args.get('input-dir'),
    platforms,
    sinceMs: args.get('since-ms'),
    annotationSource: args.get('annotation-source'),
    requireSpeaker: args.get('require-speaker') === 'true',
  });
  const reportFile = resolve(String(args.get('report-file') ?? 'data/three-platform-live-acceptance/report.json'));
  await mkdir(dirname(reportFile), { recursive: true });
  await writeFile(reportFile, `${JSON.stringify(report, null, 2)}\n`, 'utf8');
  if (args.get('json') === 'true') {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`three_platform_live_acceptance | accepted=${report.accepted ? 'yes' : 'no'} | profile=${report.acceptance_profile} | core=${report.core_accepted_platform_count}/${report.platform_count} | speaker=${report.speaker_accepted_platform_count}/${report.platform_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: accepted=${row.accepted ? 'yes' : 'no'} core=${row.core_accepted ? 'yes' : 'no'} speaker=${row.speaker_accepted ? 'yes' : 'no'} meeting=${row.meeting_id ?? '-'} speaker_identity=${row.speaker_identity || '-'} failed=${row.failed_check_ids?.join(',') || '-'} warnings=${row.warning_check_ids?.join(',') || '-'}`);
    }
    console.log(`report=${reportFile}`);
  }
  const failOnIncomplete = args.get('fail-on-incomplete') !== 'false';
  if (!report.accepted && failOnIncomplete) process.exitCode = 2;
}

if (process.argv[1] && resolve(process.argv[1]) === resolve(fileURLToPath(import.meta.url))) {
  main().catch((error) => {
    console.error(`three_platform_live_acceptance | accepted=no | error=${String(error?.message ?? error)}`);
    process.exitCode = 2;
  });
}
