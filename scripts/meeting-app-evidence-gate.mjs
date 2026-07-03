#!/usr/bin/env node

import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingAppLaunchGate,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-gate.mjs';
import {
  buildMeetingAppSnapshotRecordSet,
  meetingAppSnapshotRecords,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';

const args = new Map();
for (const raw of process.argv.slice(2)) {
  const [key, ...rest] = raw.replace(/^--/, '').split('=');
  args.set(key, rest.length ? rest.join('=') : 'true');
}

const inputFile = String(args.get('input') || args.get('file') || args.get('evidence') || '');
const jsonOutput = args.get('json') === 'true';
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const allowFixtureEvidence = args.get('allow-fixture') === 'true' || args.get('allow-fixture-evidence') === 'true';
const requireProductionReady = args.get('require-production-ready') !== 'false';
const platformArg = String(args.get('platforms') || args.get('platform') || '').trim();

function asArray(value) {
  if (Array.isArray(value)) return value;
  return value == null ? [] : [value];
}

function boolLabel(value) {
  return value ? 'yes' : 'no';
}

function unique(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function extractRecordSet(input = {}) {
  if (input?.schema === 'meeting_app_snapshot_record_set') return input;
  if (input?.type === 'meeting_app_live_evidence_package' && input.record_set) return input.record_set;
  if (input?.record_set) return input.record_set;
  if (input?.recordSet) return input.recordSet;
  if (input?.snapshot_records) return buildMeetingAppSnapshotRecordSet(input.snapshot_records, {
    id: input.id,
    label: input.label,
    source: input.source,
    createdAtMs: input.createdAtMs ?? input.created_at_ms,
  });
  if (input?.snapshotRecords) return buildMeetingAppSnapshotRecordSet(input.snapshotRecords, {
    id: input.id,
    label: input.label,
    source: input.source,
    createdAtMs: input.createdAtMs ?? input.created_at_ms,
  });
  if (input?.records) return buildMeetingAppSnapshotRecordSet(input.records, {
    id: input.id,
    label: input.label,
    source: input.source,
    createdAtMs: input.createdAtMs ?? input.created_at_ms,
  });
  if (Array.isArray(input)) return buildMeetingAppSnapshotRecordSet(input);
  return buildMeetingAppSnapshotRecordSet([]);
}

function platformsFor(recordSet = {}, input = {}) {
  if (platformArg) return unique(platformArg.split(',').map((item) => item.trim()));
  return unique([
    ...asArray(input.platforms),
    ...asArray(input.platform),
    ...asArray(recordSet.platforms),
    ...meetingAppSnapshotRecords(recordSet).map((record) => record.platform ?? record.provider),
  ]);
}

function summarizeGate(gate = {}) {
  return {
    platform: gate.platform,
    status: gate.status,
    passed: gate.passed,
    production_ready: gate.production_ready,
    evidence_level: gate.evidence_level,
    evidence_count: gate.evidence_count,
    missing_required_coverage: gate.missing_required_coverage ?? [],
    blocking_issue_codes: (gate.blocking_issues ?? []).map((item) => item.code),
    warning_codes: (gate.warnings ?? []).map((item) => item.code),
    next_actions: gate.next_actions ?? [],
  };
}

function buildSummary({ input, recordSet, platforms, gates }) {
  const rows = gates.map((gate) => summarizeGate(gate));
  const ok = gates.length > 0
    && gates.every((gate) => gate.passed)
    && (!requireProductionReady || gates.every((gate) => gate.production_ready));
  return {
    type: 'meeting_app_evidence_gate',
    ok,
    production_ready: gates.length > 0 && gates.every((gate) => gate.production_ready),
    accepted_count: gates.filter((gate) => gate.passed).length,
    production_ready_count: gates.filter((gate) => gate.production_ready).length,
    failed_count: gates.filter((gate) => gate.status === 'failed').length,
    input_file: inputFile || null,
    record_set_id: recordSet.id,
    record_count: recordSet.record_count,
    platforms,
    allow_fixture_evidence: allowFixtureEvidence,
    require_production_ready: requireProductionReady,
    source_type: input?.type ?? input?.schema ?? (Array.isArray(input) ? 'array' : 'object'),
    rows,
    gates,
    next_actions: unique(rows.flatMap((row) => row.next_actions)),
  };
}

try {
  if (!inputFile) {
    throw new Error('Missing --input=<path>. Save window.__meetingTimelineLiveCapture.exportRecords() or evidencePackage() as JSON first.');
  }
  const input = JSON.parse(await readFile(resolve(inputFile), 'utf8'));
  const recordSet = extractRecordSet(input);
  const platforms = platformsFor(recordSet, input);
  if (platforms.length === 0) {
    throw new Error('No meeting app platform was found in evidence. Pass --platform=google-meet or include platform in captured records.');
  }
  const gates = platforms.map((platform) => buildMeetingAppLaunchGate(platform, {
    recordSet,
    allowFixtureEvidence,
    requireProductionReady,
  }));
  const summary = buildSummary({ input, recordSet, platforms, gates });

  if (reportFile) {
    const target = resolve(reportFile);
    await mkdir(dirname(target), { recursive: true });
    await writeFile(target, `${JSON.stringify(summary, null, 2)}\n`, 'utf8');
  }

  if (jsonOutput) {
    console.log(JSON.stringify(summary, null, 2));
  } else {
    console.log(`meeting_app_evidence_gate | ok=${boolLabel(summary.ok)} | production_ready=${boolLabel(summary.production_ready)} | records=${summary.record_count}`);
    for (const row of summary.rows) {
      console.log(`${row.platform}: status=${row.status} production_ready=${boolLabel(row.production_ready)} evidence=${row.evidence_level}/${row.evidence_count} missing=${row.missing_required_coverage.join(',') || '-'}`);
    }
    if (summary.next_actions.length > 0) console.log(`next_actions=${summary.next_actions.join(',')}`);
  }
  if (!summary.ok) process.exitCode = 2;
} catch (error) {
  const payload = {
    type: 'meeting_app_evidence_gate',
    ok: false,
    input_file: inputFile || null,
    error: String(error?.message ?? error),
  };
  if (jsonOutput) {
    console.error(JSON.stringify(payload, null, 2));
  } else {
    console.error(`meeting_app_evidence_gate | ok=no | error=${payload.error}`);
  }
  process.exitCode = 2;
}
