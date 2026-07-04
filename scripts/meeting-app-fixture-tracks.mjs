#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingAppFixtureTrackReadinessReport,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixture-tracks.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const requiredPlatforms = unique(String(
  args.get('platforms') || args.get('required-platforms') || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

try {
  const report = buildMeetingAppFixtureTrackReadinessReport({
    platforms: requiredPlatforms,
  });
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_app_fixture_track_readiness_report | accepted=${boolLabel(report.accepted)} | platforms=${report.platform_count} | accepted=${report.accepted_count}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: ready=${boolLabel(row.ready)} roster=${row.coverage.participant_roster_count} speaker_mark=${boolLabel(row.coverage.speaker_track_mark)} participant_mark=${boolLabel(row.coverage.participant_track_mark)}`);
    }
    if (report.missing.length > 0) console.log(`missing=${report.missing.join(',')}`);
  }
  if (!report.accepted) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
