#!/usr/bin/env node

import { mkdir, readdir, readFile, stat, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingAppLaunchGate,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-gate.mjs';
import {
  buildMeetingAppDomAdaptationDiagnosis,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-profile.mjs';
import {
  normalizeMeetingPlatform,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';
import {
  boolLabel,
  extractRecordSet,
  parseCliArgs,
  platformsFor,
  summarizeGate,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();

const evidenceDir = resolve(String(args.get('dir') || args.get('evidence-dir') || 'data/meeting-app-evidence'));
const inputList = String(args.get('input') || args.get('inputs') || '').split(',').map((item) => item.trim()).filter(Boolean);
const jsonOutput = args.get('json') === 'true';
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const allowFixtureEvidence = args.get('allow-fixture') === 'true' || args.get('allow-fixture-evidence') === 'true';
const requireProductionReady = args.get('require-production-ready') !== 'false';
const failOnIncomplete = args.get('fail-on-incomplete') === 'true';
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

function normalizePlatformKey(platform) {
  try {
    return normalizeMeetingPlatform(platform);
  } catch {
    return String(platform);
  }
}

function scoreGate(gate = {}) {
  return [
    gate.production_ready ? 4 : 0,
    gate.passed ? 2 : 0,
    gate.evidence_level === 'captured_dom' ? 1 : 0,
    Number(gate.evidence_count ?? 0) / 1000,
  ].reduce((sum, value) => sum + value, 0);
}

async function collectJsonFiles(dir) {
  const files = [];
  async function walk(current) {
    let entries = [];
    try {
      entries = await readdir(current, { withFileTypes: true });
    } catch {
      return;
    }
    for (const entry of entries) {
      const path = resolve(current, entry.name);
      if (entry.isDirectory()) await walk(path);
      else if (entry.isFile() && entry.name.endsWith('.json')) files.push(path);
    }
  }
  await walk(dir);
  return files.sort();
}

async function evidenceFiles() {
  if (inputList.length > 0) return inputList.map((item) => resolve(item));
  return collectJsonFiles(evidenceDir);
}

async function readEvidenceFile(file) {
  const text = await readFile(file, 'utf8');
  const input = JSON.parse(text);
  const recordSet = extractRecordSet(input);
  const platforms = platformsFor(recordSet, input);
  const stats = await stat(file);
  return {
    file,
    input,
    recordSet,
    platforms,
    bytes: Buffer.byteLength(text, 'utf8'),
    modified_at_ms: stats.mtimeMs,
  };
}

function evaluateEvidence(evidence) {
  const gates = evidence.platforms.map((platform) => buildMeetingAppLaunchGate(platform, {
    recordSet: evidence.recordSet,
    allowFixtureEvidence,
    requireProductionReady,
  }));
  const diagnoses = evidence.platforms.map((platform) => buildMeetingAppDomAdaptationDiagnosis(platform, {
    recordSet: evidence.recordSet,
    requireMeetingEnd: requireProductionReady,
  }));
  return {
    file: evidence.file,
    record_set_id: evidence.recordSet.id,
    record_count: evidence.recordSet.record_count,
    platforms: evidence.platforms,
    bytes: evidence.bytes,
    modified_at_ms: Math.round(evidence.modified_at_ms),
    gates,
    diagnoses,
    rows: gates.map((gate) => summarizeGate(gate)),
  };
}

function bestRows(evaluations = []) {
  const byPlatform = new Map();
  for (const evaluation of evaluations) {
    for (const gate of evaluation.gates) {
      const current = byPlatform.get(gate.platform);
      const candidate = {
        platform: gate.platform,
        source_file: evaluation.file,
        record_set_id: evaluation.record_set_id,
        record_count: evaluation.record_count,
        modified_at_ms: evaluation.modified_at_ms,
        ...summarizeGate(gate),
      };
      const diagnosis = evaluation.diagnoses.find((item) => item.platform === gate.platform);
      if (diagnosis) {
        candidate.dom_diagnosis = {
          accepted: diagnosis.accepted,
          production_ready: diagnosis.production_ready,
          controls_matched: diagnosis.selector_probe?.matched?.controls === true,
          participants_matched: diagnosis.selector_probe?.matched?.participants === true,
          active_speaker_matched: diagnosis.selector_probe?.matched?.active_speaker === true,
          meeting_started: diagnosis.observer_probe?.coverage?.meeting_started === true,
          speaker_started: diagnosis.observer_probe?.coverage?.speaker_started === true,
          meeting_ended: diagnosis.observer_probe?.coverage?.meeting_ended === true,
          issue_codes: (diagnosis.issues ?? []).map((item) => item.code),
          next_actions: diagnosis.next_actions ?? [],
        };
      }
      if (!current || scoreGate(gate) > current.score) {
        byPlatform.set(gate.platform, {
          ...candidate,
          score: scoreGate(gate),
        });
      }
    }
  }
  return [...byPlatform.values()].sort((a, b) => a.platform.localeCompare(b.platform));
}

function buildSummary(files, evaluations) {
  const rows = bestRows(evaluations);
  const rowByPlatform = new Map(rows.map((row) => [row.platform, row]));
  const requiredRows = requiredPlatforms.map((platform) => {
    const platformKey = normalizePlatformKey(platform);
    const matched = rowByPlatform.get(platformKey);
    return matched ?? {
      platform: platformKey,
      status: 'missing',
      passed: false,
      production_ready: false,
      evidence_level: 'none',
      evidence_count: 0,
      missing_required_coverage: ['captured_dom_evidence'],
      blocking_issue_codes: ['missing_platform_evidence_file'],
      warning_codes: [],
      dom_diagnosis: {
        accepted: false,
        production_ready: false,
        issue_codes: ['missing_platform_evidence_file'],
        next_actions: ['capture_live_dom_snapshots_for_this_platform'],
      },
      next_actions: ['capture_live_dom_snapshots_for_this_platform'],
    };
  });
  const ok = requiredRows.length > 0
    && requiredRows.every((row) => row.passed)
    && (!requireProductionReady || requiredRows.every((row) => row.production_ready));
  return {
    type: 'meeting_app_evidence_matrix',
    ok,
    production_ready: requiredRows.length > 0 && requiredRows.every((row) => row.production_ready),
    evidence_dir: inputList.length > 0 ? null : evidenceDir,
    input_files: files,
    evidence_file_count: files.length,
    evaluated_file_count: evaluations.length,
    required_platforms: requiredPlatforms,
    allow_fixture_evidence: allowFixtureEvidence,
    require_production_ready: requireProductionReady,
    production_ready_count: requiredRows.filter((row) => row.production_ready).length,
    dom_diagnosis_accepted_count: requiredRows.filter((row) => row.dom_diagnosis?.accepted).length,
    active_speaker_ready_count: requiredRows.filter((row) => row.dom_diagnosis?.active_speaker_matched).length,
    meeting_end_ready_count: requiredRows.filter((row) => row.dom_diagnosis?.meeting_ended).length,
    failed_count: requiredRows.filter((row) => row.status === 'failed' || row.status === 'missing').length,
    missing_platforms: requiredRows.filter((row) => row.status === 'missing').map((row) => row.platform),
    rows: requiredRows,
    all_rows: rows,
    evaluations,
    next_actions: unique(requiredRows.flatMap((row) => [
      ...(row.next_actions ?? []),
      ...(row.dom_diagnosis?.next_actions ?? []),
    ])),
  };
}

try {
  const files = await evidenceFiles();
  const evaluations = [];
  const errors = [];
  for (const file of files) {
    try {
      const evidence = await readEvidenceFile(file);
      if (evidence.platforms.length === 0) {
        errors.push({ file, error: 'no_platforms_in_evidence' });
      } else {
        evaluations.push(evaluateEvidence(evidence));
      }
    } catch (error) {
      errors.push({ file, error: String(error?.message ?? error) });
    }
  }
  const summary = {
    ...buildSummary(files, evaluations),
    errors,
  };
  if (errors.length > 0 && evaluations.length === 0) summary.ok = false;

  if (reportFile) {
    const target = resolve(reportFile);
    await mkdir(dirname(target), { recursive: true });
    await writeFile(target, `${JSON.stringify(summary, null, 2)}\n`, 'utf8');
  }

  if (jsonOutput) {
    console.log(JSON.stringify(summary, null, 2));
  } else {
    console.log(`meeting_app_evidence_matrix | ok=${boolLabel(summary.ok)} | production_ready=${boolLabel(summary.production_ready)} | files=${summary.evaluated_file_count}/${summary.evidence_file_count}`);
    for (const row of summary.rows) {
      console.log(`${row.platform}: status=${row.status} production_ready=${boolLabel(row.production_ready)} observer=${boolLabel(row.dom_diagnosis?.accepted)} speaker=${boolLabel(row.dom_diagnosis?.active_speaker_matched)} end=${boolLabel(row.dom_diagnosis?.meeting_ended)} evidence=${row.evidence_level}/${row.evidence_count} file=${row.source_file ?? '-'}`);
    }
    if (summary.next_actions.length > 0) console.log(`next_actions=${summary.next_actions.join(',')}`);
  }
  if (!summary.ok && failOnIncomplete) process.exitCode = 2;
} catch (error) {
  const payload = {
    type: 'meeting_app_evidence_matrix',
    ok: false,
    evidence_dir: evidenceDir,
    error: String(error?.message ?? error),
  };
  if (jsonOutput) console.error(JSON.stringify(payload, null, 2));
  else console.error(`meeting_app_evidence_matrix | ok=no | error=${payload.error}`);
  process.exitCode = 2;
}
