#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname } from 'node:path';
import {
  buildMeetingPlatformAdaptationPackageMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adaptation-package.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

function scoreForRow(row = {}) {
  let score = 0;
  if (row.sdk_wiring_ready === true) score += 2;
  if (row.ready_for_realtime_annotations === true) score += 3;
  if (row.candidate_observation_ready === true) score += 2;
  if (row.provider_reconcilable_start_axis === true) score += 1;
  if (row.provider_reconcilable_end_axis === true) score += 1;
  if (row.provider_reconcilable_participants === true) score += 1;
  if (row.provider_reconcilable_artifacts === true) score += 1;
  if (row.provider_reconcilable_recording === true) score += 1;
  if (row.provider_reconcilable_speaker_markers === true) score += 1;
  return score;
}

function readinessLevel(score) {
  if (score >= 10) return 'ready_for_rollout';
  if (score >= 7) return 'ready_for_pilot';
  if (score >= 4) return 'pre_pilot';
  return 'bootstrap_required';
}

function rowCoverage(row = {}) {
  return {
    platform: row.platform,
    display_name: row.display_name,
    primary_surface: row.primary_surface,
    recommended_mode: row.recommended_mode,
    compatibility_score: scoreForRow(row),
    readiness: readinessLevel(scoreForRow(row)),
    realtime: row.ready_for_realtime_annotations === true,
    adapter_surface: row.adapter_selection_axis_surface,
    candidate_observer_ready: row.candidate_observation_ready === true,
    provider_path: row.provider_path,
    axis_events: {
      start_ok: row.provider_reconcilable_start_axis === true,
      end_ok: row.provider_reconcilable_end_axis === true,
      start_count: row.provider_start_event_count,
      end_count: row.provider_end_event_count,
      lifecycle_count: row.provider_lifecycle_event_count,
    },
    participant_events: {
      track_ok: row.provider_reconcilable_participants === true,
      count: row.provider_participant_event_count,
    },
    artifact_events: {
      transcript_or_backfill_ok: row.provider_reconcilable_artifacts === true,
      recording_ok: row.provider_reconcilable_recording === true,
      count: row.provider_artifact_event_count,
    },
    speaker_support: row.provider_reconcilable_speaker_markers === true,
    realtime_blocked_by_provider: row.adapter_selection_provider_blocks_realtime === true,
    transcript_blocks_realtime: row.adapter_selection_transcript_blocks_realtime === true,
    runtime_actions: row.runtime_event_action_count,
    first_next_action: row.first_next_action,
  };
}

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const outFile = String(args.get('out-file') || args.get('outFile') || '');
const jsonOutput = args.get('json') === 'true';
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()).filter(Boolean));

async function writeJson(file, value) {
  if (!file) return;
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

const matrix = buildMeetingPlatformAdaptationPackageMatrix({
  baseUrl,
  platforms: requiredPlatforms,
});

const rows = matrix.rows
  .map(rowCoverage)
  .sort((a, b) => b.compatibility_score - a.compatibility_score || a.platform.localeCompare(b.platform));

const report = {
  type: 'meeting_platform_adapter_crosswalk_report',
  base_url: baseUrl,
  platforms: requiredPlatforms,
  platform_count: rows.length,
  crosswalk_count: rows.length,
  rows,
  totals: {
    sdk_wiring_ready_count: matrix.sdk_wiring_ready_count ?? 0,
    ready_for_realtime_count: matrix.realtime_ready_count ?? 0,
    provider_axis_reconcile_count: matrix.provider_axis_reconcile_count,
    provider_participant_reconcile_count: matrix.provider_participant_reconcile_count,
    provider_artifact_reconcile_count: matrix.provider_artifact_reconcile_count,
  },
};

await writeJson(outFile, report);

if (jsonOutput) {
  console.log(JSON.stringify(report, null, 2));
} else {
  console.log(`meeting_platform_adapter_crosswalk_report | platforms=${report.platform_count} | sdk_wiring_ready=${report.totals.sdk_wiring_ready_count} | realtime_ready=${report.totals.ready_for_realtime_count} | axis_reconcile=${report.totals.provider_axis_reconcile_count} | participant_reconcile=${report.totals.provider_participant_reconcile_count} | artifact_reconcile=${report.totals.provider_artifact_reconcile_count}`);
  for (const row of rows) {
    console.log(`${row.platform}: score=${row.compatibility_score} readiness=${row.readiness} realtime=${boolLabel(row.realtime)} surface=${row.primary_surface} provider=${row.provider_path} axis(${row.axis_events.start_ok}/${row.axis_events.end_ok}, start=${row.axis_events.start_count}, end=${row.axis_events.end_count}) participants=${row.participant_events.track_ok ? 'ok' : 'no'}(${row.participant_events.count}) artifacts=${row.artifact_events.transcript_or_backfill_ok ? 'ok' : 'no'}(${row.artifact_events.count}) speaker=${boolLabel(row.speaker_support)} candidate=${boolLabel(row.candidate_observer_ready)} next=${row.first_next_action ?? 'none'}`);
  }
}
