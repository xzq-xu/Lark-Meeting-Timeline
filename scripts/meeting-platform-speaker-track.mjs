#!/usr/bin/env node

import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformSpeakerTrack,
  buildMeetingPlatformSpeakerTrackMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-speaker-track.mjs';
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
const samplesFile = String(args.get('samples-file') || args.get('input') || '');
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

function inputForPlatform(input, platform) {
  if (!input) return {};
  const key = normalizePlatformKey(platform);
  if (Array.isArray(input)) return input.filter((item) => normalizePlatformKey(item.platform ?? item.meeting?.platform ?? key) === key);
  const direct = input[key] ?? input[platform];
  if (direct) return direct;
  const samples = input.samples ?? input.activeSpeakerSamples ?? input.active_speaker_samples;
  const signals = input.signals ?? input.events;
  if (samples && typeof samples === 'object' && !Array.isArray(samples)) {
    return {
      samples: samples[key] ?? samples[platform] ?? [],
      signals: signals && typeof signals === 'object' && !Array.isArray(signals) ? (signals[key] ?? signals[platform] ?? []) : undefined,
    };
  }
  if (signals && typeof signals === 'object' && !Array.isArray(signals)) {
    return {
      samples: [],
      signals: signals[key] ?? signals[platform] ?? [],
    };
  }
  return {
    samples: Array.isArray(samples) ? samples.filter((item) => normalizePlatformKey(item.platform ?? item.meeting?.platform ?? key) === key) : [],
    signals: Array.isArray(signals) ? signals.filter((item) => normalizePlatformKey(item.platform ?? item.meeting?.platform ?? key) === key) : [],
  };
}

function trackFile(platform) {
  return resolve(outDir || 'data/meeting-platform-speaker-tracks', `${platform}.json`);
}

function trackSummary(track = {}) {
  return {
    platform: track.platform,
    status: track.status,
    input_sample_count: track.input_sample_count,
    signal_count: track.signal_count,
    segment_count: track.segment_count,
    mark_count: track.mark_count,
    dropped_duplicate_count: track.diagnostics?.dropped_duplicate_count ?? 0,
    dropped_short_segment_count: track.diagnostics?.dropped_short_segment_count ?? 0,
    open_segment_count: track.diagnostics?.open_segment_count ?? 0,
    next_actions: track.next_actions ?? [],
  };
}

async function buildReport() {
  const errors = [];
  let input;
  try {
    input = await readJson(samplesFile);
  } catch (error) {
    errors.push({ file: resolve(samplesFile), error: String(error?.message ?? error) });
  }
  const matrix = buildMeetingPlatformSpeakerTrackMatrix({
    platforms: requiredPlatforms,
  });
  const tracks = [];
  const writtenFiles = [];
  for (const platform of matrix.platforms) {
    const trackInput = inputForPlatform(input, platform);
    const track = buildMeetingPlatformSpeakerTrack(platform, trackInput, {
      minStableMs: args.get('min-stable-ms') ?? args.get('minStableMs') ?? undefined,
      switchStableMs: args.get('switch-stable-ms') ?? args.get('switchStableMs') ?? undefined,
      endIdleMs: args.get('end-idle-ms') ?? args.get('endIdleMs') ?? undefined,
      minSegmentMs: args.get('min-segment-ms') ?? args.get('minSegmentMs') ?? undefined,
      mergeGapMs: args.get('merge-gap-ms') ?? args.get('mergeGapMs') ?? undefined,
      duplicateWindowMs: args.get('duplicate-window-ms') ?? args.get('duplicateWindowMs') ?? undefined,
    });
    tracks.push(track);
    if (writeTracks) {
      const file = trackFile(track.platform);
      await writeJson(file, track);
      writtenFiles.push(file);
    }
  }
  const markCount = tracks.reduce((total, track) => total + track.mark_count, 0);
  return {
    type: 'meeting_platform_speaker_track_report',
    ok: errors.length === 0,
    requirement: 'speaker_positions_without_transcript_content',
    samples_file: samplesFile ? resolve(samplesFile) : undefined,
    platform_count: matrix.platform_count,
    mark_count: markCount,
    segment_count: tracks.reduce((total, track) => total + track.segment_count, 0),
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
    console.log(`meeting_platform_speaker_track_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | marks=${report.mark_count} | segments=${report.segment_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: status=${row.status} samples=${row.input_sample_count} signals=${row.signal_count} segments=${row.segment_count} marks=${row.mark_count} dropped_duplicate=${row.dropped_duplicate_count} dropped_short=${row.dropped_short_segment_count}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
    for (const error of report.errors) console.error(`error ${error.file}: ${error.error}`);
  }
  if (!report.ok) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
