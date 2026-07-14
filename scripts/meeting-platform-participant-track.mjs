#!/usr/bin/env node

import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformParticipantTrack,
  buildMeetingPlatformParticipantTrackMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-participant-track.mjs';
import {
  normalizeMeetingPlatform,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const outDir = String(args.get('out-dir') || args.get('outDir') || '');
const writeTracks = args.get('write-tracks') === 'true' || Boolean(outDir);
const inputFile = String(args.get('snapshots-file') || args.get('signals-file') || args.get('input') || '');
const requiredPlatforms = unique(String(
  args.get('platforms') || args.get('required-platforms') || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

function normalizePlatformKey(platform) {
  try {
    return normalizeMeetingPlatform(platform);
  } catch {
    return String(platform);
  }
}

async function readJson(file) {
  if (!file) return undefined;
  return JSON.parse(await readFile(resolve(file), 'utf8'));
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function arraysFromInput(input = {}) {
  return {
    snapshots: input.snapshots ?? input.participantSnapshots ?? input.participant_snapshots ?? input.rosterSnapshots ?? input.roster_snapshots,
    signals: input.signals ?? input.events,
  };
}

function inputForPlatform(input, platform) {
  if (!input) return {};
  const key = normalizePlatformKey(platform);
  if (Array.isArray(input)) {
    return {
      signals: input.filter((item) => normalizePlatformKey(item.platform ?? item.meeting?.platform ?? key) === key),
    };
  }
  const direct = input[key] ?? input[platform];
  if (direct) return direct;
  const { snapshots, signals } = arraysFromInput(input);
  if (snapshots && typeof snapshots === 'object' && !Array.isArray(snapshots)) {
    return {
      snapshots: snapshots[key] ?? snapshots[platform] ?? [],
      signals: signals && typeof signals === 'object' && !Array.isArray(signals) ? (signals[key] ?? signals[platform] ?? []) : undefined,
    };
  }
  if (signals && typeof signals === 'object' && !Array.isArray(signals)) {
    return {
      snapshots: [],
      signals: signals[key] ?? signals[platform] ?? [],
    };
  }
  return {
    snapshots: Array.isArray(snapshots) ? snapshots.filter((item) => normalizePlatformKey(item.platform ?? item.meeting?.platform ?? key) === key) : [],
    signals: Array.isArray(signals) ? signals.filter((item) => normalizePlatformKey(item.platform ?? item.meeting?.platform ?? key) === key) : [],
  };
}

function trackFile(platform) {
  return resolve(outDir || 'data/meeting-platform-participant-tracks', `${platform}.json`);
}

function trackSummary(track = {}) {
  return {
    platform: track.platform,
    status: track.status,
    input_snapshot_count: track.input_snapshot_count,
    signal_count: track.signal_count,
    mark_count: track.mark_count,
    dropped_duplicate_count: track.diagnostics?.dropped_duplicate_count ?? 0,
    suppressed_reconnect_pair_count: track.diagnostics?.suppressed_reconnect_pair_count ?? 0,
    pending_leave_count: track.diagnostics?.pending_leave_count ?? 0,
    next_actions: track.next_actions ?? [],
  };
}

async function buildReport() {
  const errors = [];
  let input;
  try {
    input = await readJson(inputFile);
  } catch (error) {
    errors.push({ file: resolve(inputFile), error: String(error?.message ?? error) });
  }
  const matrix = buildMeetingPlatformParticipantTrackMatrix({
    platforms: requiredPlatforms,
  });
  const tracks = [];
  const writtenFiles = [];
  for (const platform of matrix.platforms) {
    const trackInput = inputForPlatform(input, platform);
    const track = buildMeetingPlatformParticipantTrack(platform, trackInput, {
      duplicateWindowMs: args.get('duplicate-window-ms') ?? args.get('duplicateWindowMs') ?? undefined,
      suppressReconnectGapMs: args.get('suppress-reconnect-gap-ms') ?? args.get('suppressReconnectGapMs') ?? undefined,
      leaveStableMs: args.get('leave-stable-ms') ?? args.get('leaveStableMs') ?? undefined,
      emitInitialRoster: args.get('emit-initial-roster') ?? args.get('emitInitialRoster') ?? undefined,
      closePendingLeavesAtMs: args.get('close-pending-leaves-at-ms') ?? args.get('closePendingLeavesAtMs') ?? undefined,
    });
    tracks.push(track);
    if (writeTracks) {
      const file = trackFile(track.platform);
      await writeJson(file, track);
      writtenFiles.push(file);
    }
  }
  return {
    type: 'meeting_platform_participant_track_report',
    ok: errors.length === 0,
    requirement: 'participant_positions_without_transcript_content',
    input_file: inputFile ? resolve(inputFile) : undefined,
    platform_count: matrix.platform_count,
    mark_count: tracks.reduce((total, track) => total + track.mark_count, 0),
    rows: tracks.map((track) => trackSummary(track)),
    plans: matrix.rows,
    written_files: writtenFiles,
    tracks: args.get('include-tracks') === 'true' ? tracks : undefined,
    next_actions: unique(tracks.flatMap((track) => track.next_actions ?? [])),
    errors,
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_participant_track_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | marks=${report.mark_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: status=${row.status} snapshots=${row.input_snapshot_count} signals=${row.signal_count} marks=${row.mark_count} dropped_duplicate=${row.dropped_duplicate_count} suppressed_reconnect=${row.suppressed_reconnect_pair_count} pending_leave=${row.pending_leave_count}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
    for (const error of report.errors) console.error(`error ${error.file}: ${error.error}`);
  }
  if (!report.ok) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
