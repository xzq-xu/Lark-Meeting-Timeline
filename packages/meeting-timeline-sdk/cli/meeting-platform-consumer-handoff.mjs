import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformConsumerHandoff,
} from '../adapters/platform-consumer-handoff.mjs';

function boolLabel(value) {
  return value ? 'yes' : 'no';
}

function unique(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function listArg(value, fallback = []) {
  if (value == null || value === '') return fallback;
  return unique(String(value).split(',').map((item) => item.trim()).filter(Boolean));
}

function boolArg(value, fallback = false) {
  if (value == null || value === '') return fallback;
  const text = String(value).trim().toLowerCase();
  if (['1', 'true', 'yes', 'on'].includes(text)) return true;
  if (['0', 'false', 'no', 'off'].includes(text)) return false;
  return fallback;
}

export function parseMeetingPlatformConsumerHandoffCliArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

export function meetingPlatformConsumerHandoffCliOptionsFromArgs(args = new Map()) {
  return {
    baseUrl: String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787'),
    basePath: String(args.get('base-path') || args.get('basePath') || '/api/platform-events'),
    reportFile: String(args.get('report-file') || args.get('write-report') || ''),
    jsonOutput: boolArg(args.get('json'), false),
    failOnRejected: boolArg(args.get('fail-on-rejected'), false) || boolArg(args.get('fail-on-failed'), false),
    requireHandoffReady: boolArg(args.get('require-handoff-ready'), false)
      || boolArg(args.get('requireHandoffReady'), false),
    requireProductionReady: boolArg(args.get('require-production-ready'), false)
      || boolArg(args.get('requireProductionReady'), false),
    includeDetails: boolArg(args.get('include-details'), false) || boolArg(args.get('includeDetails'), false),
    requiredPlatforms: listArg(
      args.get('required-platforms') || args.get('platforms'),
      ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
    ),
  };
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

export async function buildMeetingPlatformConsumerHandoffCliReport(options = {}) {
  const {
    baseUrl = 'http://localhost:8787',
    basePath = '/api/platform-events',
    reportFile = '',
    requiredPlatforms = ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
    requireHandoffReady = false,
    requireProductionReady = false,
    includeDetails = false,
  } = options;
  const handoff = buildMeetingPlatformConsumerHandoff({
    ...options,
    baseUrl,
    basePath,
    env: options.env ?? process.env,
    platforms: requiredPlatforms,
    requireHandoffReady,
    requireProductionReady,
    includeDetails,
  });
  if (reportFile) await writeJson(resolve(reportFile), handoff);
  return handoff;
}

export function formatMeetingPlatformConsumerHandoffCliReport(handoff = {}) {
  const lines = [
    `meeting_platform_consumer_handoff | accepted=${boolLabel(handoff.accepted)} | platforms=${handoff.platform_count} | consumer_ready=${handoff.consumer_ready_count} | conformance=${handoff.conformance_accepted_count} | runtime=${handoff.runtime_ready_count} | routes=${handoff.adapter_route_ready_count} | candidates=${handoff.candidate_observer_count} | speaker=${handoff.speaker_track_ready_count} | participant=${handoff.participant_track_ready_count} | startup=${handoff.adapter_startup_ready_count} | preflight=${handoff.adapter_preflight_realtime_ready_count}/${handoff.adapter_preflight_platform_count} | handoff=${handoff.handoff_ready_count} | production=${handoff.production_ready_count} | blocking=${handoff.blocking_count} | warnings=${handoff.warning_count}`,
  ];
  for (const row of handoff.rows ?? []) {
    lines.push(`${row.platform}: consumer_ready=${boolLabel(row.consumer_ready)} route=${boolLabel(row.adapter_route_ready)} candidate=${boolLabel(row.candidate_observation_ready)} speaker=${boolLabel(row.speaker_track_ready)} participant=${boolLabel(row.participant_track_ready)} startup=${boolLabel(row.adapter_startup_ready)} preflight=${row.adapter_preflight_status ?? 'n/a'} realtime=${boolLabel(row.adapter_preflight_realtime_ready)} handoff=${boolLabel(row.handoff_ready)} production=${boolLabel(row.production_ready)} first_route=${row.adapter_first_route ?? 'n/a'} startup_surface=${row.adapter_startup_selected_surface ?? 'n/a'} next=${row.adapter_preflight_first_next_action ?? row.first_next_action ?? 'none'}`);
  }
  if (handoff.next_actions?.length > 0) lines.push(`next_actions=${handoff.next_actions.join(',')}`);
  return lines.join('\n');
}

export async function runMeetingPlatformConsumerHandoffCli(argv = process.argv.slice(2), io = console) {
  const options = meetingPlatformConsumerHandoffCliOptionsFromArgs(
    parseMeetingPlatformConsumerHandoffCliArgs(argv),
  );
  const handoff = await buildMeetingPlatformConsumerHandoffCliReport(options);
  if (options.jsonOutput) {
    io.log(JSON.stringify(handoff, null, 2));
  } else {
    io.log(formatMeetingPlatformConsumerHandoffCliReport(handoff));
  }
  if (!handoff.accepted && options.failOnRejected) process.exitCode = 2;
  return handoff;
}
