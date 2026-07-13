#!/usr/bin/env node

import { mkdir, readdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformPilotMeasurementContract,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-acceptance-checklist.mjs';
import {
  normalizeMeetingPlatform,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const inputDir = resolve(String(args.get('dir') || args.get('input-dir') || 'data/meeting-platform-field-evidence'));
const reportFile = resolve(String(args.get('report-file') || 'data/meeting-platform-p0-field-report.json'));
const jsonOutput = args.get('json') === 'true';
const failOnIncomplete = args.get('fail-on-incomplete') === 'true' || args.get('fail-on-not-ready') === 'true';
const requiredPlatforms = unique(String(
  args.get('platforms') || args.get('required-platforms') || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()).filter(Boolean)).map((platform) => normalizeMeetingPlatform(platform));

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

function numeric(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function max(values = []) {
  const rows = values.map(numeric).filter((value) => value != null);
  return rows.length ? Math.max(...rows) : null;
}

function p95(values = []) {
  const rows = values.map(numeric).filter((value) => value != null).sort((a, b) => a - b);
  if (!rows.length) return null;
  return rows[Math.max(0, Math.ceil(rows.length * 0.95) - 1)];
}

function absoluteMax(values = []) {
  return max(values.map((value) => Math.abs(Number(value))));
}

function realSnapshotRecords(input = {}) {
  return (input.meetingAppRecords ?? input.meeting_app_records ?? []).filter((record) => {
    const source = String(record?.source ?? record?.snapshot?.source ?? '').toLowerCase();
    const title = String(record?.snapshot?.title ?? '').toLowerCase();
    return record?.snapshot && !source.includes('fixture') && !title.includes('sdk fixture');
  });
}

function speakerIdentity(row = {}) {
  return String(row.speaker_id ?? row.speaker_name ?? row.participant_id ?? row.participant_name ?? row.label ?? '').trim();
}

function phaseCoverage(records = []) {
  const phases = new Set(records.map((record) => String(record.phase ?? '').toLowerCase()));
  return {
    active: phases.has('active'),
    ended: phases.has('ended'),
  };
}

function observerSurface(input = {}, records = []) {
  const explicit = String(
    input.observer_surface
      ?? input.observerSurface
      ?? input.measurements?.observer_surface
      ?? records.find((record) => record.observer_surface || record.observerSurface)?.observer_surface
      ?? records.find((record) => record.observer_surface || record.observerSurface)?.observerSurface
      ?? '',
  ).trim().toLowerCase().replace(/[-\s]+/g, '_');
  if (['native', 'native_detector', 'desktop', 'desktop_observer', 'accessibility'].includes(explicit)) {
    return 'native_detector';
  }
  if (['browser', 'browser_extension', 'content_script', 'web'].includes(explicit)) {
    return 'browser_extension';
  }
  const sourceCorpus = records.map((record) => String(record.source ?? record.snapshot?.source ?? '')).join(' ').toLowerCase();
  if (/extension|content[_ -]?script|browser/.test(sourceCorpus)) return 'browser_extension';
  if (/native|desktop|accessibility|window[_ -]?observer/.test(sourceCorpus)) return 'native_detector';
  return null;
}

function evaluateRun(input = {}, file, contract = {}) {
  const thresholds = contract.thresholds ?? {};
  const records = realSnapshotRecords(input);
  const phases = phaseCoverage(records);
  const annotations = input.annotations ?? [];
  const uniqueAnnotationIds = new Set(annotations.map((row) => String(row.id ?? '')));
  const speakerStarts = (input.speaker_markers ?? []).filter((row) => row.kind === 'speaker_started');
  const distinctSpeakers = new Set(speakerStarts.map(speakerIdentity).filter(Boolean));
  const startLatency = numeric(input.measurements?.start_detection_latency_ms);
  const endLatency = numeric(input.measurements?.end_detection_latency_ms);
  const annotationLatencies = annotations.map((row) => numeric(row.visible_latency_ms));
  const annotationTimelineErrors = annotations.map((row) => numeric(row.timeline_error_ms));
  const speakerMarkerLatencies = speakerStarts.map((row) => numeric(row.visible_latency_ms));
  const annotationLatencyComplete = annotations.length > 0
    && annotationLatencies.every((value) => value != null && value >= 0);
  const annotationTimelineComplete = annotations.length > 0
    && annotationTimelineErrors.every((value) => value != null);
  const speakerLatencyComplete = speakerStarts.length > 0
    && speakerMarkerLatencies.every((value) => value != null && value >= 0);
  const annotationP95 = p95(annotationLatencies);
  const annotationMax = max(annotationLatencies);
  const timelineErrorMax = absoluteMax(annotationTimelineErrors);
  const speakerMarkerMax = max(speakerMarkerLatencies);
  const duplicateCount = annotations.reduce((count, row) => (
    count + Math.max(0, Number(row.visible_instance_count ?? (row.duplicate_visible ? 2 : 1)) - 1)
  ), 0);
  const expectedAnnotations = Number(contract.annotations_per_meeting ?? 5);
  const lostCount = Math.max(0, expectedAnnotations - uniqueAnnotationIds.size);
  const leakageCount = numeric(input.measurements?.previous_meeting_annotation_count_on_new_axis);
  const checks = [
    { id: 'run_identity', passed: Boolean(input.run_id ?? input.runId) },
    { id: 'real_active_snapshot', passed: phases.active },
    { id: 'real_ended_snapshot', passed: phases.ended },
    { id: 'operator_join_reference', passed: numeric(input.measurements?.operator_join_at_ms) != null },
    { id: 'meeting_start_detection_latency', passed: startLatency != null && startLatency <= Number(thresholds.meeting_start_detection_latency_ms?.lte) },
    { id: 'annotation_count', passed: uniqueAnnotationIds.size >= expectedAnnotations },
    { id: 'annotation_visible_p95', passed: annotationLatencyComplete && annotationP95 != null && annotationP95 <= Number(thresholds.annotation_visible_latency_ms?.lte) },
    { id: 'annotation_visible_hard_max', passed: annotationLatencyComplete && annotationMax != null && annotationMax <= Number(thresholds.annotation_visible_latency_ms?.hard_max_lte) },
    { id: 'annotation_timeline_error', passed: annotationTimelineComplete && timelineErrorMax != null && timelineErrorMax <= Number(thresholds.annotation_timeline_error_ms?.lte) },
    { id: 'speaker_switches', passed: speakerStarts.length >= Number(contract.speaker_switches_per_meeting ?? 2) && distinctSpeakers.size >= 2 },
    { id: 'speaker_marker_visible_latency', passed: speakerLatencyComplete && speakerMarkerMax != null && speakerMarkerMax <= Number(thresholds.speaker_marker_visible_latency_ms?.lte) },
    { id: 'operator_leave_reference', passed: numeric(input.measurements?.operator_leave_at_ms) != null },
    { id: 'meeting_end_detection_latency', passed: endLatency != null && endLatency <= Number(thresholds.meeting_end_detection_latency_ms?.lte) },
    { id: 'lost_annotations', passed: lostCount === Number(thresholds.lost_annotation_count?.equals ?? 0) },
    { id: 'duplicate_annotations', passed: duplicateCount === Number(thresholds.duplicate_annotation_count?.equals ?? 0) },
    { id: 'cross_meeting_isolation', passed: leakageCount === Number(thresholds.previous_meeting_annotation_count_on_new_axis?.equals ?? 0) },
  ];
  const startedAtMs = numeric(input.measurements?.active_observer_captured_at_ms)
    ?? numeric(records.find((record) => record.phase === 'active')?.captured_at_ms)
    ?? 0;
  return {
    platform: normalizeMeetingPlatform(input.platform),
    observer_surface: contract.observer_surface ?? observerSurface(input, records),
    run_id: input.run_id ?? input.runId,
    meeting_id: input.meeting_id ?? input.meeting?.meeting_id,
    file,
    started_at_ms: startedAtMs,
    accepted: checks.every((check) => check.passed),
    real_snapshot_count: records.length,
    annotation_count: uniqueAnnotationIds.size,
    speaker_start_count: speakerStarts.length,
    distinct_speaker_count: distinctSpeakers.size,
    speaker_marker_visible_max_ms: speakerMarkerMax,
    start_detection_latency_ms: startLatency,
    annotation_visible_p95_ms: annotationP95,
    annotation_visible_max_ms: annotationMax,
    annotation_timeline_error_max_ms: timelineErrorMax,
    end_detection_latency_ms: endLatency,
    lost_annotation_count: lostCount,
    duplicate_annotation_count: duplicateCount,
    previous_meeting_annotation_count_on_new_axis: leakageCount,
    checks,
    failed_check_ids: checks.filter((check) => !check.passed).map((check) => check.id),
  };
}

async function loadRuns() {
  const files = await collectJsonFiles(inputDir);
  const rows = [];
  const errors = [];
  for (const file of files) {
    try {
      const input = JSON.parse(await readFile(file, 'utf8'));
      if (input?.schema !== 'meeting_platform_field_evidence_input') continue;
      const platform = normalizeMeetingPlatform(input.platform);
      if (!requiredPlatforms.includes(platform)) continue;
      const surface = observerSurface(input, realSnapshotRecords(input));
      rows.push(evaluateRun(input, file, buildMeetingPlatformPilotMeasurementContract(platform, {
        observerSurface: surface ?? undefined,
      })));
    } catch (error) {
      errors.push({ file, error: String(error?.message ?? error) });
    }
  }
  return { files, rows, errors };
}

function summarizePlatform(platform, runs = []) {
  const contract = buildMeetingPlatformPilotMeasurementContract(platform);
  const required = Number(contract.required_consecutive_real_meetings ?? 3);
  const sorted = runs.filter((run) => run.platform === platform).sort((a, b) => a.started_at_ms - b.started_at_ms);
  const latest = sorted.slice(-required);
  const accepted = latest.length === required && latest.every((run) => run.accepted);
  return {
    platform,
    display_name: contract.display_name,
    required_consecutive_real_meetings: required,
    run_count: sorted.length,
    evaluated_run_count: latest.length,
    accepted_run_count: latest.filter((run) => run.accepted).length,
    accepted,
    status: accepted ? 'p0_field_accepted' : latest.length < required ? 'needs_more_real_meetings' : 'latest_consecutive_runs_failed',
    latest_runs: latest,
    failed_check_ids: unique(latest.flatMap((run) => run.failed_check_ids)),
    next_action: accepted ? 'proceed_to_provider_reconcile_p1' : latest.length < required ? 'run_more_real_meetings' : 'fix_failed_p0_checks_and_repeat_until_three_consecutive_pass',
  };
}

const loaded = await loadRuns();
const platforms = requiredPlatforms.map((platform) => summarizePlatform(platform, loaded.rows));
const report = {
  type: 'meeting_platform_p0_field_report',
  schema: 'meeting_platform_p0_field_report',
  schema_version: 1,
  generated_at: new Date().toISOString(),
  input_dir: inputDir,
  required_platforms: requiredPlatforms,
  platform_count: platforms.length,
  accepted_platform_count: platforms.filter((row) => row.accepted).length,
  accepted: platforms.length > 0 && platforms.every((row) => row.accepted) && loaded.errors.length === 0,
  platforms,
  run_count: loaded.rows.length,
  source_file_count: loaded.files.length,
  errors: loaded.errors,
};

await mkdir(dirname(reportFile), { recursive: true });
await writeFile(reportFile, `${JSON.stringify(report, null, 2)}\n`, 'utf8');

if (jsonOutput) {
  console.log(JSON.stringify(report, null, 2));
} else {
  console.log(`meeting_platform_p0_field_report | accepted=${boolLabel(report.accepted)} | platforms=${report.accepted_platform_count}/${report.platform_count} | runs=${report.run_count}`);
  for (const row of platforms) {
    console.log(`${row.platform}: accepted=${boolLabel(row.accepted)} runs=${row.accepted_run_count}/${row.required_consecutive_real_meetings} status=${row.status} failed=${row.failed_check_ids.join(',') || '-'}`);
  }
}

if (!report.accepted && failOnIncomplete) process.exitCode = 2;
