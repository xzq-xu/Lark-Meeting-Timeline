import {
  buildMeetingAppSnapshotRecordSet,
  meetingAppSnapshotRecords,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';

export function asArray(value) {
  if (Array.isArray(value)) return value;
  return value == null ? [] : [value];
}

export function boolLabel(value) {
  return value ? 'yes' : 'no';
}

export function unique(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

export function parseCliArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

export function extractRecordSet(input = {}) {
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

export function platformsFor(recordSet = {}, input = {}, platformArg = '') {
  if (platformArg) return unique(String(platformArg).split(',').map((item) => item.trim()));
  return unique([
    ...asArray(input.platforms),
    ...asArray(input.platform),
    ...asArray(recordSet.platforms),
    ...meetingAppSnapshotRecords(recordSet).map((record) => record.platform ?? record.provider),
  ]);
}

export function summarizeGate(gate = {}) {
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
