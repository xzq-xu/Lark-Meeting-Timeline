#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformConsumerHandoff,
} from '../packages/meeting-timeline-sdk/adapters/platform-consumer-handoff.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const basePath = String(args.get('base-path') || args.get('basePath') || '/api/platform-events');
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const failOnRejected = args.get('fail-on-rejected') === 'true' || args.get('fail-on-failed') === 'true';
const requireHandoffReady = args.get('require-handoff-ready') === 'true' || args.get('requireHandoffReady') === 'true';
const requireProductionReady = args.get('require-production-ready') === 'true' || args.get('requireProductionReady') === 'true';
const includeDetails = args.get('include-details') === 'true' || args.get('includeDetails') === 'true';
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()).filter(Boolean));

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

const handoff = buildMeetingPlatformConsumerHandoff({
  baseUrl,
  basePath,
  env: process.env,
  platforms: requiredPlatforms,
  requireHandoffReady,
  requireProductionReady,
  includeDetails,
});

if (reportFile) await writeJson(resolve(reportFile), handoff);

if (jsonOutput) {
  console.log(JSON.stringify(handoff, null, 2));
} else {
  console.log(`meeting_platform_consumer_handoff | accepted=${boolLabel(handoff.accepted)} | platforms=${handoff.platform_count} | consumer_ready=${handoff.consumer_ready_count} | conformance=${handoff.conformance_accepted_count} | runtime=${handoff.runtime_ready_count} | routes=${handoff.adapter_route_ready_count} | candidates=${handoff.candidate_observer_count} | speaker=${handoff.speaker_track_ready_count} | participant=${handoff.participant_track_ready_count} | handoff=${handoff.handoff_ready_count} | production=${handoff.production_ready_count} | blocking=${handoff.blocking_count} | warnings=${handoff.warning_count}`);
  for (const row of handoff.rows) {
    console.log(`${row.platform}: consumer_ready=${boolLabel(row.consumer_ready)} route=${boolLabel(row.adapter_route_ready)} candidate=${boolLabel(row.candidate_observation_ready)} speaker=${boolLabel(row.speaker_track_ready)} participant=${boolLabel(row.participant_track_ready)} handoff=${boolLabel(row.handoff_ready)} production=${boolLabel(row.production_ready)} first_route=${row.adapter_first_route ?? 'n/a'} next=${row.first_next_action ?? 'none'}`);
  }
  if (handoff.next_actions.length > 0) console.log(`next_actions=${handoff.next_actions.join(',')}`);
}

if (!handoff.accepted && failOnRejected) process.exitCode = 2;
