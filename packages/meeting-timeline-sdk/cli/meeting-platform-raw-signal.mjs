import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { basename, dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformRawSignalBatch,
  buildMeetingPlatformRawSignalExampleBatch,
} from '../adapters/platform-raw-signal.mjs';

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

function numberArg(value, fallback = undefined) {
  if (value == null || value === '') return fallback;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function compactUndefined(value = {}) {
  return Object.fromEntries(Object.entries(value).filter(([, item]) => item !== undefined));
}

export function parseMeetingPlatformRawSignalCliArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

export function meetingPlatformRawSignalCliOptionsFromArgs(args = new Map()) {
  return {
    inputFile: String(args.get('input-file') || args.get('inputFile') || ''),
    outFile: String(args.get('out-file') || args.get('outFile') || ''),
    reportFile: String(args.get('report-file') || args.get('write-report') || ''),
    jsonOutput: boolArg(args.get('json'), false),
    examples: boolArg(args.get('examples') || args.get('example-batch'), false),
    stdin: boolArg(args.get('stdin'), false),
    includeSignals: boolArg(args.get('include-signals'), false),
    includeBatch: boolArg(args.get('include-batch'), false),
    includeRuntimeEvents: boolArg(args.get('include-runtime-events'), true),
    filterActiveSpeakerSamples: boolArg(
      args.get('filter-active-speaker-samples') || args.get('filterActiveSpeakerSamples'),
      false,
    ),
    minStableMs: numberArg(args.get('min-stable-ms') || args.get('minStableMs'), undefined),
    switchStableMs: numberArg(args.get('switch-stable-ms') || args.get('switchStableMs'), undefined),
    endIdleMs: numberArg(args.get('end-idle-ms') || args.get('endIdleMs'), undefined),
    platforms: listArg(args.get('platforms') || args.get('platform-keys'), ['google-meet', 'teams', 'zoom', 'webex', 'lark']),
  };
}

async function readJson(file) {
  return JSON.parse(await readFile(resolve(file), 'utf8'));
}

async function readStdinJson() {
  const text = await readFile(0, 'utf8');
  return JSON.parse(text);
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function eventRows(events = []) {
  return events.map((event, index) => compactUndefined({
    index,
    action: event.action,
    platform: event.platform,
    source: event.source,
    sent_at_ms: event.sent_at_ms,
    signal_count: Array.isArray(event.signals) ? event.signals.length : undefined,
    annotation_label: event.annotation?.label,
    meeting_id: event.current_meeting?.meeting_id ?? event.meeting?.meeting_id,
  }));
}

function actionCounts(events = []) {
  const counts = {};
  for (const event of events) counts[event.action] = (counts[event.action] ?? 0) + 1;
  return counts;
}

function compactBatch(batch = {}, options = {}) {
  return compactUndefined({
    type: batch.type,
    schema: batch.schema,
    schema_version: batch.schema_version,
    signal_count: batch.signal_count,
    runtime_event_count: batch.runtime_event_count,
    filtered_speaker_event_count: batch.filtered_speaker_event_count,
    platform_count: batch.platform_count,
    platforms: batch.platforms,
    kinds: batch.kinds,
    rows: batch.rows,
    signals: options.includeSignals ? batch.signals : undefined,
    runtime_events: options.includeRuntimeEvents ? batch.runtime_events : undefined,
  });
}

async function inputForOptions(options = {}) {
  if (options.inputFile) return readJson(options.inputFile);
  if (options.stdin) return readStdinJson();
  return undefined;
}

export async function buildMeetingPlatformRawSignalCliReport(options = {}) {
  const input = await inputForOptions(options);
  const useExamples = options.examples === true || input == null;
  const signalOptions = compactUndefined({
    ...options,
    platforms: options.platforms,
    filterActiveSpeakerSamples: options.filterActiveSpeakerSamples,
    minStableMs: options.minStableMs,
    switchStableMs: options.switchStableMs,
    endIdleMs: options.endIdleMs,
  });
  const batch = useExamples
    ? buildMeetingPlatformRawSignalExampleBatch(signalOptions)
    : buildMeetingPlatformRawSignalBatch(input, signalOptions);
  const runtimeEvents = batch.runtime_events ?? [];
  const report = compactUndefined({
    type: 'meeting_platform_raw_signal_report',
    schema: 'meeting_platform_raw_signal_report',
    ok: batch.signal_count > 0 && batch.runtime_event_count > 0,
    mode: useExamples ? 'examples' : 'input',
    input_file: options.inputFile || undefined,
    out_file: options.outFile || undefined,
    report_file: options.reportFile || undefined,
    filter_active_speaker_samples: options.filterActiveSpeakerSamples,
    signal_count: batch.signal_count,
    runtime_event_count: batch.runtime_event_count,
    filtered_speaker_event_count: batch.filtered_speaker_event_count,
    platform_count: batch.platform_count,
    platforms: batch.platforms,
    kinds: batch.kinds,
    runtime_actions: unique(runtimeEvents.map((event) => event.action)),
    action_counts: actionCounts(runtimeEvents),
    rows: batch.rows,
    event_rows: eventRows(runtimeEvents),
    runtime_events: options.includeRuntimeEvents ? runtimeEvents : undefined,
    batch: options.includeBatch ? compactBatch(batch, options) : undefined,
    next_actions: [
      'wire_runtime_events_to_sdk_sendRuntimeEvent_or_server_runtime_event_endpoint',
      'enable_filter_active_speaker_samples_for_noisy_speaker_tiles',
      'keep_provider_events_as_reconcile_not_realtime_gate',
    ],
  });
  if (options.outFile) await writeJson(resolve(options.outFile), options.includeBatch ? report : compactBatch(batch, options));
  if (options.reportFile) await writeJson(resolve(options.reportFile), report);
  return report;
}

export function formatMeetingPlatformRawSignalCliReport(report = {}) {
  const lines = [
    `meeting_platform_raw_signal_report | ok=${boolLabel(report.ok)} | mode=${report.mode} | signals=${report.signal_count} | runtime_events=${report.runtime_event_count} | filtered_speaker=${report.filtered_speaker_event_count ?? 0} | platforms=${(report.platforms ?? []).join(',')}`,
  ];
  for (const row of report.rows ?? []) {
    lines.push(`${row.platform ?? 'n/a'}:${row.kind} source=${row.source ?? 'n/a'} events=${row.runtime_event_count} actions=${(row.runtime_actions ?? []).join(',')}`);
  }
  const counts = Object.entries(report.action_counts ?? {}).map(([action, count]) => `${action}=${count}`);
  if (counts.length > 0) lines.push(`actions=${counts.join(',')}`);
  if (report.out_file) lines.push(`out_file=${basename(report.out_file)}`);
  if (report.report_file) lines.push(`report_file=${basename(report.report_file)}`);
  return lines.join('\n');
}

export async function runMeetingPlatformRawSignalCli(argv = process.argv.slice(2), io = console) {
  const options = meetingPlatformRawSignalCliOptionsFromArgs(
    parseMeetingPlatformRawSignalCliArgs(argv),
  );
  const report = await buildMeetingPlatformRawSignalCliReport(options);
  if (options.jsonOutput) {
    io.log(JSON.stringify(report, null, 2));
  } else {
    io.log(formatMeetingPlatformRawSignalCliReport(report));
  }
  if (!report.ok) process.exitCode = 2;
  return report;
}
